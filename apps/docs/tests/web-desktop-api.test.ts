import { describe, expect, it, vi } from 'vitest'
import type { OfficeFile, OfficeHostApi, SaveDocumentInput } from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { createDocsWebDesktopController } from '../src/web/desktop-api'
import { installWebHostPolicy } from '../src/web/host-policy'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

function bytesOf(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer
}

function createHarness(
  saveImpl?: (input: SaveDocumentInput) => ReturnType<OfficeHostApi['saveDocument']>,
) {
  const initialFile: OfficeFile = {
    id: 'doc-1',
    name: '测试文档.docx',
    mimeType: DOCX_MIME,
    size: 6,
    version: 'v1',
    transport: 'buffer',
    bytes: bytesOf('source'),
  }

  const saveDocument = vi.fn(
    saveImpl ??
      (async (input: SaveDocumentInput) => ({
        ok: true,
        file: { ...input.file, version: 'v2' },
      })),
  )
  const host: OfficeHostApi = {
    getLocale: vi.fn(async () => 'zh-CN'),
    saveDocument,
    pickDocument: vi.fn(async () => ({
      status: 'cancelled' as const,
      selectionId: null,
      file: null,
    })),
    confirmDocumentOpened: vi.fn(async () => ({ ok: true })),
    releasePickedDocument: vi.fn(async () => {}),
    pickAssets: vi.fn(async () => ({ status: 'cancelled' as const, files: [] })),
    downloadDocument: vi.fn(async () => ({ ok: true })),
    approveClose: vi.fn(async () => {}),
    cancelClose: vi.fn(async () => {}),
    pickFile: vi.fn(async () => null),
    readFile: vi.fn(async () => initialFile),
    setDirty: vi.fn(),
    setTitle: vi.fn(),
  }

  const incoming = new Set<(message: HostToEditorMessage) => void>()
  const send = vi.fn()
  const bridge = {
    send,
    subscribe: vi.fn((handler: (message: HostToEditorMessage) => void) => {
      incoming.add(handler)
      return () => incoming.delete(handler)
    }),
  } as unknown as EditorIframeBridge

  const policy = installWebHostPolicy('embedded', bridge)
  const controller = createDocsWebDesktopController(host, bridge)
  const emit = (message: HostToEditorMessage) => {
    if (incoming.size === 0) throw new Error('Bridge handler is not registered')
    for (const handler of incoming) handler(message)
  }
  const destroy = () => {
    policy.destroy()
    controller.destroy()
  }

  return { controller, policy, host, saveDocument, send, emit, initialFile, destroy }
}

describe('Docs web desktop adapter', () => {
  it('serializes initial ready/init, applies parent mode, and applies Web UI capabilities', async () => {
    const { controller, policy, host, send, emit, initialFile, destroy } = createHarness()

    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()

    expect(send).toHaveBeenCalledWith({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:ready',
      payload: { kind: 'docx' },
    })

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: 'init-1',
      payload: {
        kind: 'docx',
        mode: 'view',
        locale: 'zh-CN',
        capabilities: {
          ai: false,
          autoSave: 'host',
          pageCropMarks: true,
        },
        file: initialFile,
      },
    })

    const opened = await pendingOpen
    expect(opened?.name).toBe(initialFile.name)
    expect(opened?.path).toContain('web-office://files/doc-1/')
    expect(new TextDecoder().decode(opened?.data)).toBe('source')
    expect(await controller.desktopApi.getHostEditorMode?.()).toBe('view')
    expect(await controller.desktopApi.getLanguage()).toBe('zh')
    expect(policy.getCapabilities()).toMatchObject({
      ai: false,
      open: true,
      save: true,
      saveAs: true,
      autoSave: 'host',
      pageCropMarks: true,
    })
    expect(document.documentElement.classList.contains('office-web')).toBe(true)
    expect(document.documentElement.classList.contains('office-ai-enabled')).toBe(false)
    expect(document.documentElement.classList.contains('office-page-crop-marks')).toBe(true)
    expect(host.setTitle).toHaveBeenCalledWith(initialFile.name)

    const modeChanged = vi.fn()
    const offMode = controller.desktopApi.onHostEditorModeChanged?.(modeChanged)
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:set-mode',
      payload: { mode: 'edit' },
    })
    expect(modeChanged).toHaveBeenCalledWith('edit')
    expect(await controller.desktopApi.getHostEditorMode?.()).toBe('edit')
    offMode?.()

    controller.desktopApi.reportDirtyChange?.(true)
    expect(host.setDirty).toHaveBeenCalledWith(true)

    destroy()
    expect(document.documentElement.classList.contains('office-web')).toBe(false)
  })

  it('starts a blank embedded document, switches locale live, and marks its first save', async () => {
    const { controller, emit, saveDocument, destroy } = createHarness()
    const changed = vi.fn()
    const offLanguage = controller.desktopApi.onLanguageChanged(changed)
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:new',
      requestId: 'new-1',
      payload: { kind: 'docx', mode: 'edit', locale: 'pt-BR', capabilities: { open: true } },
    })

    expect(await pendingOpen).toBeNull()
    expect(await controller.desktopApi.getLanguage()).toBe('pt')
    expect(changed).toHaveBeenCalledWith('pt')

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:set-locale',
      payload: { locale: 'zh-TW' },
    })
    expect(await controller.desktopApi.getLanguage()).toBe('zh-TW')
    expect(changed).toHaveBeenLastCalledWith('zh-TW')

    await controller.desktopApi.saveDocxNew('新建文档.docx', bytesOf('draft'))
    expect(saveDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        newDocument: true,
        file: expect.objectContaining({ name: '新建文档.docx' }),
      }),
    )

    offLanguage()
    destroy()
  })

  it('routes parent save through the host and acknowledges the original request', async () => {
    const { controller, saveDocument, send, emit, initialFile, destroy } = createHarness()
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: 'init-2',
      payload: { kind: 'docx', mode: 'edit', file: initialFile },
    })
    const opened = await pendingOpen
    expect(opened).not.toBeNull()

    controller.desktopApi.onMenuCommand((command) => {
      if (command === 'save') {
        void controller.desktopApi.saveDocx(opened!.path, bytesOf('edited'))
      }
    })

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save',
      requestId: 'parent-save-1',
    })

    expect(saveDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        baseVersion: 'v1',
        mode: 'save',
        file: expect.objectContaining({ id: 'doc-1', name: initialFile.name }),
      }),
    )
    const saveInput = saveDocument.mock.calls[0]?.[0]
    expect(new TextDecoder().decode(saveInput?.bytes)).toBe('edited')

    await vi.waitFor(() => {
      expect(send).toHaveBeenCalledWith({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-result',
        requestId: 'parent-save-1',
        payload: { ok: true, error: undefined },
      })
    })

    destroy()
  })

  it('marks Save As explicitly so the host can create a new system file', async () => {
    const { controller, saveDocument, emit, initialFile, host, destroy } = createHarness(
      async (input) => ({
        ok: true,
        file: {
          ...input.file,
          id: 'doc-copy',
          name: '副本.docx',
          version: 'v1',
        },
      }),
    )
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: 'init-save-as',
      payload: { kind: 'docx', mode: 'edit', file: initialFile },
    })
    await pendingOpen

    const result = await controller.desktopApi.saveDocxAs('副本.docx', bytesOf('copy'))

    expect(saveDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        mode: 'saveAs',
        baseVersion: 'v1',
        file: expect.objectContaining({ id: 'doc-1', name: '副本.docx' }),
      }),
    )
    expect(result.ok).toBe(true)
    expect(result.path).toContain('web-office://files/doc-copy/')
    expect(host.setTitle).toHaveBeenLastCalledWith('副本.docx')

    destroy()
  })

  it('routes the parent window close request through the guarded editor lifecycle', async () => {
    const { controller, host, emit, destroy } = createHarness()
    const requested = vi.fn()
    const off = controller.desktopApi.onHostCloseRequest?.(requested)

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:request-close',
      requestId: 'window-close-1',
      payload: { reason: 'window-close' },
    })
    expect(requested).toHaveBeenCalledTimes(1)

    await controller.desktopApi.requestHostClose?.()
    expect(host.approveClose).toHaveBeenCalledWith('window-close-1')

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:request-close',
      requestId: 'window-close-2',
      payload: { reason: 'window-close' },
    })
    controller.desktopApi.cancelHostCloseRequest?.()
    expect(host.cancelClose).toHaveBeenCalledWith('window-close-2')

    off?.()
    destroy()
  })

  it('keeps a picked document pending until the renderer confirms it opened', async () => {
    const { controller, host, destroy } = createHarness()
    const candidate: OfficeFile = {
      id: 'doc-2',
      nodeId: 'doc-2',
      name: '候选文档.docx',
      mimeType: DOCX_MIME,
      size: 9,
      version: 7,
      bytes: bytesOf('candidate'),
      transport: 'buffer',
    }
    vi.mocked(host.pickDocument!).mockResolvedValue({
      status: 'selected',
      selectionId: 'selection-2',
      file: candidate,
    })
    vi.mocked(host.confirmDocumentOpened!).mockResolvedValue({
      ok: true,
      file: {
        id: candidate.id,
        nodeId: candidate.nodeId,
        name: candidate.name,
        mimeType: candidate.mimeType,
        size: candidate.size,
        version: candidate.version,
        transport: 'buffer',
      },
    })

    const picked = await controller.desktopApi.openDocx()
    expect(picked?.selectionId).toBe('selection-2')
    expect(host.confirmDocumentOpened).not.toHaveBeenCalled()
    expect(host.setTitle).not.toHaveBeenCalledWith(candidate.name)

    const confirmed = await controller.desktopApi.confirmOpenDocx?.('selection-2')
    expect(confirmed).toEqual({ ok: true })
    expect(host.confirmDocumentOpened).toHaveBeenCalledWith('selection-2')
    expect(host.setTitle).toHaveBeenCalledWith(candidate.name)

    destroy()
  })

  it('propagates a failed document picker without reading selection fields', async () => {
    const { controller, host, destroy } = createHarness()
    vi.mocked(host.pickDocument!).mockResolvedValue({
      status: 'failed',
      code: 'PERMISSION_DENIED',
      error: 'The Host denied document selection.',
    })

    await expect(controller.desktopApi.openDocx()).rejects.toMatchObject({
      code: 'PERMISSION_DENIED',
      message: 'The Host denied document selection.',
    })
    expect(host.confirmDocumentOpened).not.toHaveBeenCalled()

    destroy()
  })

  it('rejects a Host save success that omits the authoritative file descriptor', async () => {
    const { controller, emit, initialFile, host, destroy } = createHarness(async () => ({
      ok: true,
    }))
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: 'init-missing-save-file',
      payload: { kind: 'docx', mode: 'edit', file: initialFile },
    })
    const opened = await pendingOpen

    const result = await controller.desktopApi.saveDocx(opened!.path, bytesOf('edited'))
    expect(result.ok).toBe(false)
    expect(result.error).toContain('latest file descriptor')
    expect(host.setTitle).toHaveBeenCalledTimes(1)

    destroy()
  })

  it('maps host version conflicts to the existing external-modified save contract', async () => {
    const { controller, emit, initialFile, destroy } = createHarness(async () => ({
      ok: false,
      code: 'VERSION_CONFLICT',
      error: 'The document changed on the server.',
    }))
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: 'init-3',
      payload: { kind: 'docx', mode: 'edit', file: initialFile },
    })
    const opened = await pendingOpen

    const result = await controller.desktopApi.saveDocx(opened!.path, bytesOf('edited'))
    expect(result).toEqual({
      ok: false,
      error: 'The document changed on the server.',
      reason: 'external-modified',
    })

    destroy()
  })
})
