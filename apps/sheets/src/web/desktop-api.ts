import type { AiChatResponse, AiSettings, AiStreamChunk, GenSparkAccountStatus } from '@genoffice/ai-provider'
import type {
  OfficeEditorMode,
  OfficeFile,
  OfficeHostApi,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import type {
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentReadResult,
  DesktopApi,
  MenuAction,
  ScreenCaptureResult,
  ScreenSourcesResult,
  WorkbookExportPdfResult,
  WorkbookFile,
  WorkbookFormulaCellsRequest,
  WorkbookFormulaCellsResult,
  WorkbookMediaRequest,
  WorkbookMediaResult,
  WorkbookPivotDefinition,
  WorkbookPivotRequest,
  WorkbookRangeRequest,
  WorkbookRangeResult,
  WorkbookRecalcRequest,
  WorkbookRecalcResult,
  WorkbookSaveRequest,
  WorkbookSaveResult,
} from '../shared/desktop-api'
import { deleteXlsxSession, openXlsxWorkbookBytes, readXlsxWorkbookRange } from './engine-client'

type SheetsLanguage = Awaited<ReturnType<DesktopApi['getLanguage']>>
type LanguageHandler = Parameters<DesktopApi['onLanguageChanged']>[0]
type MenuHandler = Parameters<DesktopApi['onMenuAction']>[0]

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

export interface SheetsWebDesktopController {
  desktopApi: DesktopApi
  notifyReady(): void
  destroy(): void
}

function unavailable(name: string): never {
  throw new Error(`${name} is not available in Sheets Web yet.`)
}

function noopUnsubscribe(): () => void {
  return () => undefined
}

function normalizeLanguage(locale: string): SheetsLanguage {
  const value = locale.toLowerCase()
  if (value.startsWith('zh')) return 'zh'
  if (value.startsWith('ja')) return 'ja'
  if (value.startsWith('ko')) return 'ko'
  if (value.startsWith('fr')) return 'fr'
  if (value.startsWith('de')) return 'de'
  if (value.startsWith('es')) return 'es'
  if (value.startsWith('th')) return 'th'
  if (value.startsWith('id')) return 'id'
  if (value.startsWith('ru')) return 'ru'
  if (value.startsWith('ar')) return 'ar'
  return 'en'
}

async function selectedToOfficeFile(
  host: OfficeHostApi,
  selected: SelectedOfficeFile,
): Promise<OfficeFile> {
  if (selected.transport === 'buffer' && selected.bytes) {
    return {
      id: selected.id,
      name: selected.name,
      mimeType: selected.mimeType || XLSX_MIME,
      size: selected.size ?? selected.bytes.byteLength,
      version: selected.version ?? null,
      bytes: selected.bytes,
    }
  }
  return host.readFile(selected.id)
}

export function createSheetsWebDesktopController(
  host: OfficeHostApi,
  bridge?: EditorIframeBridge,
): SheetsWebDesktopController {
  let currentLanguage = normalizeLanguage(document.documentElement.lang || navigator.language || 'en')
  let currentMode: OfficeEditorMode = 'edit'
  let currentTitle = 'Untitled.xlsx'
  let pendingWorkbook: WorkbookFile | null = null
  let pendingNewBlank = false
  let pendingOpenSignal = false
  let readyNotified = false
  let dirty = false
  let saving = false
  let aiSettings = { provider: '', providers: {} } as unknown as AiSettings

  const languageHandlers = new Set<LanguageHandler>()
  const menuHandlers = new Set<MenuHandler>()

  const setLanguage = (locale: string): void => {
    const next = normalizeLanguage(locale)
    if (next === currentLanguage) return
    currentLanguage = next
    for (const handler of languageHandlers) handler(next)
  }

  const emitOpen = (): void => {
    if (menuHandlers.size === 0) {
      pendingOpenSignal = true
      return
    }
    pendingOpenSignal = false
    for (const handler of menuHandlers) handler('open')
  }

  const setWorkbookFromOfficeFile = async (file: OfficeFile): Promise<void> => {
    if (pendingWorkbook) {
      await deleteXlsxSession(pendingWorkbook.sessionId).catch(() => undefined)
    }
    pendingWorkbook = await openXlsxWorkbookBytes(file.name, file.bytes.slice(0))
    currentTitle = file.name
    host.setTitle(file.name)
  }

  const pickWorkbook = async (): Promise<WorkbookFile | null> => {
    if (pendingWorkbook) {
      const workbook = pendingWorkbook
      pendingWorkbook = null
      return workbook
    }

    const selected = await host.pickFile({
      multiple: false,
      accept: [XLSX_MIME, '.xlsx'],
      mode: 'file',
    })
    if (!selected?.[0]) return null
    const file = await selectedToOfficeFile(host, selected[0])
    const workbook = await openXlsxWorkbookBytes(file.name, file.bytes.slice(0))
    currentTitle = file.name
    host.setTitle(file.name)
    return workbook
  }

  const handleBridgeMessage = async (message: HostToEditorMessage): Promise<void> => {
    if (message.type === 'office:init') {
      if (message.payload.kind !== 'xlsx') return
      currentMode = message.payload.mode
      document.documentElement.dataset.officeMode = currentMode
      if (message.payload.locale) setLanguage(message.payload.locale)
      await setWorkbookFromOfficeFile(message.payload.file)
      emitOpen()
      return
    }

    if (message.type === 'office:new') {
      if (message.payload.kind !== 'xlsx') return
      currentMode = message.payload.mode
      document.documentElement.dataset.officeMode = currentMode
      if (message.payload.locale) setLanguage(message.payload.locale)
      pendingNewBlank = true
      currentTitle = 'Untitled.xlsx'
      host.setTitle(currentTitle)
      return
    }

    if (message.type === 'office:set-locale') {
      setLanguage(message.payload.locale)
      return
    }

    if (message.type === 'office:set-mode') {
      currentMode = message.payload.mode
      document.documentElement.dataset.officeMode = currentMode
      return
    }

    if (message.type === 'office:query-state') {
      bridge?.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:state-result',
        requestId: message.requestId,
        payload: {
          ready: readyNotified,
          dirty,
          saving,
          mode: currentMode,
          title: currentTitle,
        },
      })
      return
    }

    if (message.type === 'office:save') {
      for (const handler of menuHandlers) handler('save')
    }
  }

  const unsubscribeBridge = bridge?.subscribe((message) => {
    void handleBridgeMessage(message).catch((error) => {
      bridge.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:error',
        requestId: 'requestId' in message ? message.requestId : undefined,
        payload: {
          code: 'SHEETS_WEB_HOST_ERROR',
          message: error instanceof Error ? error.message : String(error),
        },
      })
    })
  })

  const desktopApi: DesktopApi = {
    getLanguage: async () => currentLanguage,
    onLanguageChanged: (handler) => {
      languageHandlers.add(handler)
      return () => languageHandlers.delete(handler)
    },

    selectWorkbook: pickWorkbook,
    readWorkbookRange: async (request: WorkbookRangeRequest): Promise<WorkbookRangeResult> =>
      readXlsxWorkbookRange(request),
    readWorkbookFormulas: async (
      _request: WorkbookFormulaCellsRequest,
    ): Promise<WorkbookFormulaCellsResult> => unavailable('readWorkbookFormulas'),
    recalcWorkbook: async (_request: WorkbookRecalcRequest): Promise<WorkbookRecalcResult> =>
      unavailable('recalcWorkbook'),
    readWorkbookMedia: async (_request: WorkbookMediaRequest): Promise<WorkbookMediaResult> =>
      unavailable('readWorkbookMedia'),
    readPivotDefinition: async (_request: WorkbookPivotRequest): Promise<WorkbookPivotDefinition> =>
      unavailable('readPivotDefinition'),
    readLocalImage: async () => unavailable('readLocalImage'),
    captureScreenSources: async (): Promise<ScreenSourcesResult> => ({
      status: 'denied',
      sources: [],
    }),
    captureScreenSource: async (): Promise<ScreenCaptureResult | null> => null,
    saveWorkbookEdits: async (_request: WorkbookSaveRequest): Promise<WorkbookSaveResult> => {
      saving = true
      try {
        return { canceled: true }
      } finally {
        saving = false
      }
    },
    writeWorkbookRecovery: async () => ({ ok: true }),
    autoRenameWorkbook: async () => ({ renamed: false }),
    exportPdf: async (): Promise<WorkbookExportPdfResult> => ({ canceled: true }),
    closeWorkbook: deleteXlsxSession,
    openExternal: async (url: string) => {
      if (/^https?:\/\//.test(url)) window.open(url, '_blank', 'noopener,noreferrer')
    },

    onMenuAction: (handler) => {
      menuHandlers.add(handler)
      if (pendingOpenSignal) queueMicrotask(() => emitOpen())
      return () => menuHandlers.delete(handler)
    },
    onWorkbookRenamed: () => noopUnsubscribe(),
    notifyPendingEdits: (count: number) => {
      const next = count > 0
      if (dirty === next) return
      dirty = next
      host.setDirty(next)
    },
    onCloseSaveRequest: () => noopUnsubscribe(),
    reportCloseSaveResult: () => undefined,
    consumeNewBlankWorkbook: async () => {
      const value = pendingNewBlank
      pendingNewBlank = false
      return value
    },
    hasQueuedWorkbook: async () => pendingWorkbook !== null,

    // AI is platform-disabled in the Web Office product policy. Keep a valid
    // bridge surface so the existing renderer does not need Web-only forks.
    getAiSettings: async () => aiSettings,
    setAiSettings: async (settings) => {
      aiSettings = settings
    },
    aiChat: async (): Promise<AiChatResponse> => unavailable('aiChat'),
    aiStream: async () => unavailable('aiStream'),
    aiStreamCancel: async () => undefined,
    aiGskStatus: async (): Promise<GenSparkAccountStatus> =>
      ({ loggedIn: false }) as GenSparkAccountStatus,
    aiGskLogin: async () => undefined,
    webSearch: async () => ({ results: [], method: 'disabled' }),
    onAiStream: (_handler: (chunk: AiStreamChunk) => void) => noopUnsubscribe(),

    pickAttachments: async (): Promise<AttachmentAddResult | null> => null,
    addAttachmentPaths: async (): Promise<AttachmentAddResult> => ({ accepted: [], rejected: [] }),
    addPastedImage: async (): Promise<AttachmentAddResult> => ({ accepted: [], rejected: [] }),
    readAttachment: async (): Promise<AttachmentReadResult> => ({
      ok: false,
      error: 'Attachments are disabled in Sheets Web.',
    }),
    readAttachmentImage: async (): Promise<AttachmentImageResult> => ({
      ok: false,
      error: 'Attachments are disabled in Sheets Web.',
    }),
    getPathForFile: (file: File) => `browser-file://${encodeURIComponent(file.name)}`,
  }

  return {
    desktopApi,
    notifyReady: () => {
      if (readyNotified) return
      readyNotified = true
      bridge?.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:ready',
        payload: { kind: 'xlsx' },
      })
    },
    destroy: () => {
      unsubscribeBridge?.()
      languageHandlers.clear()
      menuHandlers.clear()
      if (pendingWorkbook) void deleteXlsxSession(pendingWorkbook.sessionId)
      pendingWorkbook = null
    },
  }
}

export function createSheetsWebDesktopApi(host: OfficeHostApi): DesktopApi {
  return createSheetsWebDesktopController(host).desktopApi
}
