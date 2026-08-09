import { webcrypto } from 'node:crypto'
import { beforeAll, describe, expect, it, vi } from 'vitest'
import type { OfficeHostApi, SaveDocumentInput } from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { createSlidesWebController } from '../src/web/slides-api'

beforeAll(() => {
  if (!globalThis.crypto?.subtle) {
    Object.defineProperty(globalThis, 'crypto', {
      configurable: true,
      value: webcrypto,
    })
  }
})

function createHarness() {
  const saveDocument = vi.fn(async (input: SaveDocumentInput) => ({
    ok: true as const,
    file: {
      ...input.file,
      id: input.newDocument ? 'ppt-created' : input.file.id,
      version: 'v2',
    },
  }))
  const saveHistoryVersion = vi.fn(async () => ({ ok: true as const }))
  const exportDocument = vi.fn(async () => ({ ok: true as const }))
  const requestClose = vi.fn(async () => {})
  const host: OfficeHostApi = {
    getLocale: vi.fn(async () => 'zh-CN'),
    saveDocument,
    saveHistoryVersion,
    exportDocument,
    requestClose,
    pickFile: vi.fn(async () => null),
    readFile: vi.fn(async () => {
      throw new Error('not used')
    }),
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

  const controller = createSlidesWebController(host, bridge)
  const emit = (message: HostToEditorMessage) => {
    for (const handler of incoming) handler(message)
  }

  return {
    controller,
    host,
    saveDocument,
    saveHistoryVersion,
    exportDocument,
    requestClose,
    send,
    emit,
  }
}

describe('Slides Web host adapter', () => {
  it('starts an App Center blank deck, persists it as a new document, and switches mode/locale live', async () => {
    const { controller, host, saveDocument, send, emit } = createHarness()

    const pendingOpen = controller.slidesApi.consumePendingOpen(960)
    expect(send).toHaveBeenCalledWith({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:ready',
      payload: { kind: 'pptx' },
    })

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:new',
      requestId: 'new-1',
      payload: {
        kind: 'pptx',
        mode: 'edit',
        locale: 'pt-BR',
      },
    })

    expect(await pendingOpen).toBeNull()
    expect(await controller.slidesApi.getLanguage()).toBe('pt')
    expect(await controller.slidesApi.getHostEditorMode?.()).toBe('edit')

    const blank = await controller.slidesApi.newBlank(960)
    expect(blank.slides).toHaveLength(1)
    const editablePlaceholder = blank.slides[0]?.nodes.find(
      (node: any) => node.placeholder && (node.type === 'text' || node.type === 'shape'),
    )
    expect(editablePlaceholder).toBeTruthy()

    const added = await controller.slidesApi.editText({
      slideIndex: 0,
      sourceId: editablePlaceholder!.sourceId,
      paragraphs: [{ runs: [{ text: '新建 PPT Web' }] }],
    })
    expect(added).not.toBeNull()
    expect(host.setDirty).toHaveBeenCalledWith(true)

    const saved = await controller.slidesApi.save()
    expect(saved.ok).toBe(true)
    expect(saveDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        mode: 'save',
        newDocument: true,
        file: expect.objectContaining({ mimeType: expect.stringContaining('presentationml') }),
      }),
    )

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:set-locale',
      payload: { locale: 'zh-TW' },
    })
    expect(await controller.slidesApi.getLanguage()).toBe('zh-TW')

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:set-mode',
      payload: { mode: 'view' },
    })
    expect(await controller.slidesApi.getHostEditorMode?.()).toBe('view')
    expect(document.documentElement.classList.contains('office-view-mode')).toBe(true)

    const blocked = await controller.slidesApi.addElement({
      slideIndex: 0,
      kind: 'textbox',
      xPx: 10,
      yPx: 10,
      wPx: 100,
      hPx: 40,
      fitWidthPx: 960,
      text: 'should not edit',
    })
    expect(blocked).toBeNull()

    controller.destroy()
  })

  it('delegates history, PPTX export, and close to the Host without replacing the current file identity', async () => {
    const { controller, saveHistoryVersion, exportDocument, requestClose, emit } = createHarness()
    const pendingOpen = controller.slidesApi.consumePendingOpen(960)

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:new',
      requestId: 'new-lifecycle',
      payload: { kind: 'pptx', mode: 'edit', locale: 'zh-CN' },
    })
    await pendingOpen
    await controller.slidesApi.newBlank(960)
    await controller.slidesApi.save()

    expect(await controller.slidesApi.saveHistoryVersion?.()).toEqual({
      ok: true,
      error: undefined,
    })
    expect(saveHistoryVersion).toHaveBeenCalledWith(
      expect.objectContaining({
        file: expect.objectContaining({ id: 'ppt-created' }),
        baseVersion: 'v2',
      }),
    )

    expect(await controller.slidesApi.exportPptx?.()).toEqual({
      ok: true,
      error: undefined,
    })
    expect(exportDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        format: 'pptx',
        file: expect.objectContaining({ id: 'ppt-created' }),
      }),
    )

    await controller.slidesApi.requestHostClose?.()
    expect(requestClose).toHaveBeenCalledTimes(1)

    const render = await controller.slidesApi.getRenderSlides()
    expect(render).toHaveLength(1)

    controller.destroy()
  })
})
