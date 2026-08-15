import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { OfficeHostApi } from '@genoffice/office-host-api'
import type { HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'

const engine = vi.hoisted(() => ({
  createBlankXlsxWorkbook: vi.fn(),
  deleteXlsxSession: vi.fn(async () => undefined),
  openXlsxWorkbookBytes: vi.fn(),
  readXlsxWorkbookFormulaCells: vi.fn(),
  readXlsxWorkbookMedia: vi.fn(),
  readXlsxWorkbookRange: vi.fn(),
  recalcXlsxWorkbook: vi.fn(),
  saveXlsxArchiveMutation: vi.fn(),
}))

vi.mock('../src/web/engine-client', () => engine)
vi.mock('../src/web/xlsx-save', () => ({
  saveWorkbookRequestViaEngine: vi.fn(),
}))

import { createSheetsWebDesktopController } from '../src/web/desktop-api'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function buffer(...values: number[]): ArrayBuffer {
  return new Uint8Array(values).buffer
}

function workbook(sessionId: string, name: string) {
  return {
    sessionId,
    name,
    sheetNames: ['Sheet1'],
  } as never
}

function installBrowserGlobals(): void {
  const target = new EventTarget()
  vi.stubGlobal('window', target)
  vi.stubGlobal('document', {
    documentElement: {
      lang: 'en-US',
      dataset: {} as Record<string, string>,
    },
  })
  vi.stubGlobal('navigator', { language: 'en-US' })
}

function createHost() {
  const host = {
    getLocale: vi.fn(async () => 'en-US'),
    saveDocument: vi.fn(),
    pickDocument: vi.fn(),
    confirmDocumentOpened: vi.fn(),
    releasePickedDocument: vi.fn(async () => undefined),
    pickAssets: vi.fn(),
    approveClose: vi.fn(async () => undefined),
    cancelClose: vi.fn(async () => undefined),
    pickFile: vi.fn(),
    readFile: vi.fn(),
    setDirty: vi.fn(),
    setTitle: vi.fn(),
  } as unknown as OfficeHostApi
  return host
}

function createBridge() {
  let handler: ((message: HostToEditorMessage) => void) | null = null
  const bridge = {
    subscribe: vi.fn((next: (message: HostToEditorMessage) => void) => {
      handler = next
      return () => {
        handler = null
      }
    }),
    send: vi.fn(),
  } as unknown as EditorIframeBridge

  return {
    bridge,
    emit(message: HostToEditorMessage) {
      if (!handler) throw new Error('Bridge handler is not installed')
      handler(message)
    },
  }
}

async function initializeOldWorkbook(host: OfficeHostApi, bridge: ReturnType<typeof createBridge>) {
  engine.openXlsxWorkbookBytes.mockResolvedValueOnce(workbook('old-session', 'old.xlsx'))
  const controller = createSheetsWebDesktopController(host, bridge.bridge)

  bridge.emit({
    protocol: 1,
    type: 'office:init',
    requestId: 'init-old',
    payload: {
      kind: 'xlsx',
      mode: 'edit',
      file: {
        id: 'old-node',
        nodeId: 'old-node',
        name: 'old.xlsx',
        mimeType: XLSX_MIME,
        version: 4,
        bytes: buffer(1),
        transport: 'buffer',
      },
    },
  })

  await vi.waitFor(() => expect(engine.openXlsxWorkbookBytes).toHaveBeenCalledTimes(1))
  // The renderer consumes the workbook queued by office:init before a later
  // File -> Open action asks the Host for another document.
  const initial = await controller.desktopApi.selectWorkbook()
  expect(initial?.sessionId).toBe('old-session')
  return controller
}

beforeEach(() => {
  vi.clearAllMocks()
  installBrowserGlobals()
})

describe('Sheets Web transactional document open', () => {
  it('keeps the old workbook active and releases the selection when candidate parsing fails', async () => {
    const host = createHost()
    const bridge = createBridge()
    const controller = await initializeOldWorkbook(host, bridge)

    vi.mocked(host.pickDocument!).mockResolvedValue({
      status: 'selected',
      selectionId: 'selection-bad',
      file: {
        id: 'bad-node',
        nodeId: 'bad-node',
        name: 'bad.xlsx',
        mimeType: XLSX_MIME,
        version: 1,
        bytes: buffer(2),
        transport: 'buffer',
      },
    })
    engine.openXlsxWorkbookBytes.mockRejectedValueOnce(new Error('invalid xlsx'))

    await expect(controller.desktopApi.selectWorkbook()).rejects.toThrow('invalid xlsx')

    expect(host.confirmDocumentOpened).not.toHaveBeenCalled()
    expect(host.releasePickedDocument).toHaveBeenCalledWith('selection-bad')
    expect(engine.deleteXlsxSession).not.toHaveBeenCalledWith('old-session')

    controller.destroy()
  })

  it('deletes only the candidate session when Host binding fails', async () => {
    const host = createHost()
    const bridge = createBridge()
    const controller = await initializeOldWorkbook(host, bridge)

    vi.mocked(host.pickDocument!).mockResolvedValue({
      status: 'selected',
      selectionId: 'selection-bind-fail',
      file: {
        id: 'candidate-node',
        nodeId: 'candidate-node',
        name: 'candidate.xlsx',
        mimeType: XLSX_MIME,
        version: 2,
        bytes: buffer(3),
        transport: 'buffer',
      },
    })
    engine.openXlsxWorkbookBytes.mockResolvedValueOnce(
      workbook('candidate-session', 'candidate.xlsx'),
    )
    vi.mocked(host.confirmDocumentOpened!).mockResolvedValue({
      ok: false,
      code: 'PERMISSION_REQUIRED',
      error: 'binding denied',
    })

    await expect(controller.desktopApi.selectWorkbook()).rejects.toThrow('binding denied')

    expect(host.releasePickedDocument).toHaveBeenCalledWith('selection-bind-fail')
    expect(engine.deleteXlsxSession).toHaveBeenCalledWith('candidate-session')
    expect(engine.deleteXlsxSession).not.toHaveBeenCalledWith('old-session')

    controller.destroy()
  })

  it('replaces the old session only after parsing and binding both succeed', async () => {
    const host = createHost()
    const bridge = createBridge()
    const controller = await initializeOldWorkbook(host, bridge)

    vi.mocked(host.pickDocument!).mockResolvedValue({
      status: 'selected',
      selectionId: 'selection-ok',
      file: {
        id: 'candidate-node',
        nodeId: 'candidate-node',
        name: 'candidate.xlsx',
        mimeType: XLSX_MIME,
        version: 2,
        bytes: buffer(4),
        transport: 'buffer',
      },
    })
    engine.openXlsxWorkbookBytes.mockResolvedValueOnce(
      workbook('candidate-session', 'candidate.xlsx'),
    )
    vi.mocked(host.confirmDocumentOpened!).mockResolvedValue({
      ok: true,
      file: {
        id: 'candidate-node',
        nodeId: 'candidate-node',
        name: 'candidate.xlsx',
        mimeType: XLSX_MIME,
        version: 3,
      },
    })

    const opened = await controller.desktopApi.selectWorkbook()

    expect(opened?.sessionId).toBe('candidate-session')
    expect(host.confirmDocumentOpened).toHaveBeenCalledWith('selection-ok')
    expect(host.releasePickedDocument).not.toHaveBeenCalled()
    expect(engine.deleteXlsxSession).toHaveBeenCalledWith('old-session')
    expect(engine.deleteXlsxSession).not.toHaveBeenCalledWith('candidate-session')
    expect(host.setTitle).toHaveBeenLastCalledWith('candidate.xlsx')

    controller.destroy()
  })
})
