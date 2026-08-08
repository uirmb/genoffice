import { describe, expect, it, vi } from 'vitest'
import type { OfficeFile, OfficeHostApi, SaveDocumentInput } from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { createDocsWebDesktopController } from '../src/web/desktop-api'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

function bytesOf(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer
}

function createHarness(saveImpl?: (input: SaveDocumentInput) => ReturnType<OfficeHostApi['saveDocument']>) {
  const initialFile: OfficeFile = {
    id: 'doc-1',
    name: '测试文档.docx',
    mimeType: DOCX_MIME,
    size: 6,
    version: 'v1',
    bytes: bytesOf('source'),
  }

  const host: OfficeHostApi = {
    getLocale: vi.fn(async () => 'zh-CN'),
    saveDocument: vi.fn(
      saveImpl ??
        (async (input) => ({
          ok: true,
          file: { ...input.file, version: 'v2' },
        })),
    ),
    pickFile: vi.fn(async () => null),
    readFile: vi.fn(async () => initialFile),
    setDirty: vi.fn(),
    setTitle: vi.fn(),
  }

  let incoming: ((message: HostToEditorMessage) => void) | null = null
  const send = vi.fn()
  const bridge = {
    send,
    subscribe: vi.fn((handler: (message: HostToEditorMessage) => void) => {
      incoming = handler
      return () => {
        incoming = null
      }
    }),
  } as unknown as EditorIframeBridge

  const controller = createDocsWebDesktopController(host, bridge)
  const emit = (message: HostToEditorMessage) => {
    if (!incoming) throw new Error('Bridge handler is not registered')
    incoming(message)
  }

  return { controller, host, send, emit, initialFile }
}

async function flushAsync(): Promise<void> {
  await Promise.resolve()
  await Promise.resolve()
  await new Promise((resolve) => setTimeout(resolve, 0))
}

describe('Docs web desktop adapter', () => {
  it('serializes the initial ready/init handshake and applies parent mode', async () => {
    const { controller, host, send, emit, initialFile } = createHarness()

    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()

    expect(send).toHaveBeenCalledWith({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:ready',
      payload: { kind: 'docx' },
    })

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      payload: {
        kind: 'docx',
        mode: 'view',
        locale: 'zh-CN',
        file: initialFile,
      },
    })

    const opened = await pendingOpen
    expect(opened?.name).toBe(initialFile.name)
    expect(opened?.path).toContain('web-office://files/doc-1/')
    expect(new TextDecoder().decode(opened?.data)).toBe('source')
    expect(await controller.desktopApi.getHostEditorMode?.()).toBe('view')
    expect(await controller.desktopApi.getLanguage()).toBe('zh')
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

    controller.destroy()
  })

  it('routes parent save through the host and acknowledges the original request', async () => {
    const { controller, host, send, emit, initialFile } = createHarness()
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
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
    await flushAsync()

    expect(host.saveDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        baseVersion: 'v1',
        file: expect.objectContaining({ id: 'doc-1', name: initialFile.name }),
      }),
    )
    const saveInput = vi.mocked(host.saveDocument).mock.calls[0]?.[0]
    expect(new TextDecoder().decode(saveInput?.bytes)).toBe('edited')
    expect(send).toHaveBeenCalledWith({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-result',
      requestId: 'parent-save-1',
      payload: { ok: true, error: undefined },
    })

    controller.destroy()
  })

  it('maps host version conflicts to the existing external-modified save contract', async () => {
    const { controller, emit, initialFile } = createHarness(async () => ({
      ok: false,
      code: 'VERSION_CONFLICT',
      error: 'The document changed on the server.',
    }))
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      payload: { kind: 'docx', mode: 'edit', file: initialFile },
    })
    const opened = await pendingOpen

    const result = await controller.desktopApi.saveDocx(opened!.path, bytesOf('edited'))
    expect(result).toEqual({
      ok: false,
      error: 'The document changed on the server.',
      reason: 'external-modified',
    })

    controller.destroy()
  })
})
