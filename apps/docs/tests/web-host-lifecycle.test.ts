import { describe, expect, it, vi } from 'vitest'
import type { OfficeFile, OfficeHostApi } from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { createDocsWebDesktopController } from '../src/web/desktop-api'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

function bytesOf(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer
}

describe('Docs Web host lifecycle actions', () => {
  it('delegates history, DOCX export, and close without replacing the current document identity', async () => {
    const initialFile: OfficeFile = {
      id: 'doc-1',
      name: '当前文档.docx',
      mimeType: DOCX_MIME,
      size: 6,
      version: 'v7',
      bytes: bytesOf('source'),
    }

    const saveHistoryVersion = vi.fn(async () => ({ ok: true as const }))
    const exportDocument = vi.fn(async () => ({ ok: true as const }))
    const requestClose = vi.fn(async () => {})
    const host: OfficeHostApi = {
      getLocale: vi.fn(async () => 'zh-CN'),
      saveDocument: vi.fn(async (input) => ({ ok: true, file: input.file })),
      saveHistoryVersion,
      exportDocument,
      requestClose,
      pickFile: vi.fn(async () => null),
      readFile: vi.fn(async () => initialFile),
      setDirty: vi.fn(),
      setTitle: vi.fn(),
    }

    const incoming = new Set<(message: HostToEditorMessage) => void>()
    const bridge = {
      send: vi.fn(),
      subscribe: vi.fn((handler: (message: HostToEditorMessage) => void) => {
        incoming.add(handler)
        return () => incoming.delete(handler)
      }),
    } as unknown as EditorIframeBridge

    const controller = createDocsWebDesktopController(host, bridge)
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()
    for (const handler of incoming) {
      handler({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:init',
        requestId: 'init-lifecycle',
        payload: { kind: 'docx', mode: 'edit', file: initialFile },
      })
    }

    const opened = await pendingOpen
    expect(opened).not.toBeNull()

    const historyBytes = bytesOf('history snapshot')
    expect(await controller.desktopApi.saveHistoryVersion?.(initialFile.name, historyBytes)).toEqual({
      ok: true,
      error: undefined,
    })
    expect(saveHistoryVersion).toHaveBeenCalledWith(
      expect.objectContaining({
        file: expect.objectContaining({ id: 'doc-1', name: initialFile.name, version: 'v7' }),
        baseVersion: 'v7',
      }),
    )

    const exportBytes = bytesOf('export snapshot')
    expect(await controller.desktopApi.exportDocx?.('导出副本.docx', exportBytes)).toEqual({
      ok: true,
      error: undefined,
    })
    expect(exportDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        format: 'docx',
        file: expect.objectContaining({ id: 'doc-1', name: '导出副本.docx' }),
      }),
    )

    await controller.desktopApi.requestHostClose?.()
    expect(requestClose).toHaveBeenCalledTimes(1)

    // History/export are side effects only; the open/save identity remains the
    // same platform file until a real Save As succeeds.
    expect(await controller.desktopApi.openDocxPath(opened!.path)).toMatchObject({
      path: opened!.path,
      name: initialFile.name,
    })

    controller.destroy()
  })
})
