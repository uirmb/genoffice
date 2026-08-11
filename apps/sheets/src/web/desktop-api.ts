import type { AiChatResponse, AiSettings, AiStreamChunk, GenSparkAccountStatus } from '@genoffice/ai-provider'
import type {
  OfficeEditorMode,
  OfficeFile,
  OfficeFileDescriptor,
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
import {
  createBlankXlsxWorkbook,
  deleteXlsxSession,
  openXlsxWorkbookBytes,
  readXlsxWorkbookFormulaCells,
  readXlsxWorkbookMedia,
  readXlsxWorkbookRange,
  recalcXlsxWorkbook,
  saveXlsxArchiveMutation,
} from './engine-client'
import { saveWorkbookRequestViaEngine } from './xlsx-save'

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

function hasWorkbookMutations(request: WorkbookSaveRequest): boolean {
  return (
    request.edits.length > 0 ||
    request.structuralOps.length > 0 ||
    request.chartEdits.length > 0 ||
    request.visualEdits.length > 0 ||
    request.visualAdditions.length > 0 ||
    request.tableAdditions.length > 0 ||
    request.pivotAdditions.length > 0 ||
    request.sheetOps.length > 0 ||
    request.filterStates.length > 0 ||
    request.hyperlinkEdits.length > 0 ||
    request.cfStates.length > 0 ||
    request.dvStates.length > 0 ||
    request.pageSetupStates.length > 0 ||
    request.noteStates.length > 0 ||
    request.formulaValues.length > 0 ||
    request.pivotCacheRefreshPaths.length > 0 ||
    request.pivotRefreshUpdates.length > 0 ||
    request.sheetProtections.length > 0 ||
    request.sparklineAdditions.length > 0 ||
    request.definedNamesState !== null
  )
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

function officeDescriptor(file: OfficeFile | null, workbook: WorkbookFile): OfficeFileDescriptor {
  if (file) {
    return {
      id: file.id,
      name: file.name,
      mimeType: file.mimeType || XLSX_MIME,
      ...(file.size === undefined ? {} : { size: file.size }),
      ...(file.version === undefined ? {} : { version: file.version }),
    }
  }
  return {
    id: `new:${crypto.randomUUID()}`,
    name: workbook.name,
    mimeType: XLSX_MIME,
    version: null,
  }
}

export function createSheetsWebDesktopController(
  host: OfficeHostApi,
  bridge?: EditorIframeBridge,
): SheetsWebDesktopController {
  let currentLanguage = normalizeLanguage(document.documentElement.lang || navigator.language || 'en')
  let currentMode: OfficeEditorMode = 'edit'
  let currentTitle = 'Untitled.xlsx'
  let currentOfficeFile: OfficeFile | null = null
  let activeWorkbook: WorkbookFile | null = null
  let pendingWorkbook: WorkbookFile | null = null
  let pendingOpenSignal = false
  let currentIsNewDocument = false
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

  const setActiveWorkbook = (workbook: WorkbookFile): void => {
    activeWorkbook = workbook
    pendingWorkbook = workbook
    currentTitle = workbook.name
    host.setTitle(workbook.name)
  }

  const setWorkbookFromOfficeFile = async (file: OfficeFile): Promise<void> => {
    if (activeWorkbook) {
      await deleteXlsxSession(activeWorkbook.sessionId).catch(() => undefined)
    }
    currentOfficeFile = file
    currentIsNewDocument = false
    setActiveWorkbook(await openXlsxWorkbookBytes(file.name, file.bytes.slice(0)))
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
    if (activeWorkbook) {
      await deleteXlsxSession(activeWorkbook.sessionId).catch(() => undefined)
    }
    currentOfficeFile = file
    currentIsNewDocument = false
    const workbook = await openXlsxWorkbookBytes(file.name, file.bytes.slice(0))
    activeWorkbook = workbook
    currentTitle = file.name
    host.setTitle(file.name)
    return workbook
  }

  const createNewWorkbook = async (): Promise<void> => {
    if (activeWorkbook) {
      await deleteXlsxSession(activeWorkbook.sessionId).catch(() => undefined)
    }
    const workbook = await createBlankXlsxWorkbook('Untitled.xlsx')
    currentOfficeFile = null
    currentIsNewDocument = true
    setActiveWorkbook(workbook)
    emitOpen()
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
      await createNewWorkbook()
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
      const requestId = 'requestId' in message ? message.requestId : undefined
      bridge.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:error',
        ...(requestId === undefined ? {} : { requestId }),
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
      request: WorkbookFormulaCellsRequest,
    ): Promise<WorkbookFormulaCellsResult> => readXlsxWorkbookFormulaCells(request),
    recalcWorkbook: async (request: WorkbookRecalcRequest): Promise<WorkbookRecalcResult> => {
      if (!activeWorkbook) throw new Error('No active workbook session.')
      return recalcXlsxWorkbook(request, activeWorkbook)
    },
    readWorkbookMedia: async (request: WorkbookMediaRequest): Promise<WorkbookMediaResult> => {
      if (!activeWorkbook) throw new Error('No active workbook session.')
      return readXlsxWorkbookMedia(request, activeWorkbook)
    },
    readPivotDefinition: async (_request: WorkbookPivotRequest): Promise<WorkbookPivotDefinition> =>
      unavailable('readPivotDefinition'),
    readLocalImage: async () => unavailable('readLocalImage'),
    captureScreenSources: async (): Promise<ScreenSourcesResult> => ({
      status: 'denied',
      sources: [],
    }),
    captureScreenSource: async (): Promise<ScreenCaptureResult | null> => null,
    saveWorkbookEdits: async (request: WorkbookSaveRequest): Promise<WorkbookSaveResult> => {
      if (currentMode !== 'edit') throw new Error('Workbook is read-only.')
      if (!activeWorkbook) throw new Error('No active workbook session.')

      saving = true
      const previousSessionId = activeWorkbook.sessionId
      try {
        const saved =
          request.mode === 'save-as' && !hasWorkbookMutations(request)
            ? {
                ...(await saveXlsxArchiveMutation(request.sessionId, activeWorkbook.name, {
                  replacements: new Map(),
                  removals: [],
                  additions: new Map(),
                })),
                touchedEntries: [] as readonly string[],
              }
            : await saveWorkbookRequestViaEngine(request, activeWorkbook, activeWorkbook.name)
        const descriptor = officeDescriptor(currentOfficeFile, activeWorkbook)
        const result = await host.saveDocument({
          file: descriptor,
          bytes: saved.bytes.slice(0),
          baseVersion: descriptor.version ?? null,
          mode: request.mode === 'save-as' ? 'saveAs' : 'save',
          newDocument: currentIsNewDocument,
        })

        if (!result.ok || !result.file) {
          await deleteXlsxSession(saved.file.sessionId).catch(() => undefined)
          return { canceled: true }
        }

        const nextWorkbook: WorkbookFile = {
          ...saved.file,
          name: result.file.name,
        }
        activeWorkbook = nextWorkbook
        currentOfficeFile = {
          ...result.file,
          bytes: saved.bytes.slice(0),
        }
        currentTitle = result.file.name
        currentIsNewDocument = false
        dirty = false
        host.setTitle(currentTitle)
        host.setDirty(false)
        if (previousSessionId !== nextWorkbook.sessionId) {
          await deleteXlsxSession(previousSessionId).catch(() => undefined)
        }

        return {
          canceled: false,
          file: nextWorkbook,
          touchedEntries: [...saved.touchedEntries],
        }
      } finally {
        saving = false
      }
    },
    writeWorkbookRecovery: async () => ({ ok: true }),
    autoRenameWorkbook: async () => ({ renamed: false }),
    exportPdf: async (): Promise<WorkbookExportPdfResult> => ({ canceled: true }),
    closeWorkbook: async (sessionId: string) => {
      if (activeWorkbook?.sessionId === sessionId) activeWorkbook = null
      await deleteXlsxSession(sessionId)
    },
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
    consumeNewBlankWorkbook: async () => false,
    hasQueuedWorkbook: async () => pendingWorkbook !== null,

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
      if (activeWorkbook) void deleteXlsxSession(activeWorkbook.sessionId)
      activeWorkbook = null
      pendingWorkbook = null
    },
  }
}

export function createSheetsWebDesktopApi(host: OfficeHostApi): DesktopApi {
  return createSheetsWebDesktopController(host).desktopApi
}
