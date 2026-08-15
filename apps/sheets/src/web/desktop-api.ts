import type {
  AiChatResponse,
  AiSettings,
  AiStreamChunk,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'
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
import {
  SHEETS_WEB_FILE_ACTION_EVENT,
  type SheetsWebFileAction,
  type SheetsWebSnapshotHost,
} from './file-actions'
import { saveWorkbookRequestViaEngine } from './xlsx-save'

type SheetsLanguage = Awaited<ReturnType<DesktopApi['getLanguage']>>
type LanguageHandler = Parameters<DesktopApi['onLanguageChanged']>[0]
type MenuHandler = Parameters<DesktopApi['onMenuAction']>[0]
type MenuAction = Parameters<MenuHandler>[0]
type RendererFileAction = MenuAction | 'save-history' | 'export-xlsx'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

interface MaterializedWorkbook {
  file: WorkbookFile
  bytes: ArrayBuffer
  touchedEntries: readonly string[]
}

export interface SheetsWebDesktopController {
  desktopApi: DesktopApi
  snapshotHost: SheetsWebSnapshotHost
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
      ...(selected.nodeId ? { nodeId: selected.nodeId } : {}),
      ...(selected.tenantId ? { tenantId: selected.tenantId } : {}),
      name: selected.name,
      mimeType: selected.mimeType || XLSX_MIME,
      size: selected.size ?? selected.bytes.byteLength,
      version: selected.version ?? null,
      bytes: selected.bytes,
      transport: 'buffer',
    }
  }
  return host.readFile(selected.id)
}

function officeDescriptor(file: OfficeFile | null, workbook: WorkbookFile): OfficeFileDescriptor {
  if (file) {
    return {
      id: file.id,
      ...(file.nodeId ? { nodeId: file.nodeId } : {}),
      ...(file.tenantId ? { tenantId: file.tenantId } : {}),
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

function emptySecondPhaseRequest(
  sessionId: string,
  request: WorkbookSaveRequest,
  tableAdditions: WorkbookSaveRequest['tableAdditions'],
  pivotAdditions: WorkbookSaveRequest['pivotAdditions'],
  definedNamesState: WorkbookSaveRequest['definedNamesState'],
): WorkbookSaveRequest {
  return {
    sessionId,
    mode: 'save',
    edits: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions,
    pivotAdditions,
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState,
  }
}

export function createSheetsWebDesktopController(
  host: OfficeHostApi,
  bridge?: EditorIframeBridge,
): SheetsWebDesktopController {
  let currentLanguage = normalizeLanguage(
    document.documentElement.lang || navigator.language || 'en',
  )
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
  let closeAfterSave = false
  let aiSettings = { provider: '', providers: {} } as unknown as AiSettings

  const languageHandlers = new Set<LanguageHandler>()
  const menuHandlers = new Set<MenuHandler>()

  const dispatchRendererFileAction = (action: RendererFileAction): void => {
    // save-history/export-xlsx are Web-only extensions. The shared renderer's
    // menu callback ultimately forwards unknown file modes to handleSave();
    // keep the Electron MenuAction contract untouched.
    for (const handler of menuHandlers) handler(action as MenuAction)
  }

  const handleWebFileShortcut = (event: KeyboardEvent): void => {
    if (event.repeat || event.altKey || !(event.metaKey || event.ctrlKey)) return
    const key = event.key.toLowerCase()
    let action: MenuAction | null = null
    if (key === 's') action = event.shiftKey ? 'save-as' : 'save'
    else if (key === 'o' && !event.shiftKey) action = 'open'
    if (!action) return

    // Electron owns these accelerators in the desktop build. Sheets Web has no
    // native application menu, so keep the browser from interpreting Ctrl/Cmd+S
    // as "save this HTML page" and route all file commands through the same
    // renderer menu-action path instead.
    event.preventDefault()
    if (action === 'open' && menuHandlers.size === 0) {
      pendingOpenSignal = true
      return
    }
    dispatchRendererFileAction(action)
  }
  window.addEventListener('keydown', handleWebFileShortcut)

  const requestHostClose = async (): Promise<void> => {
    if (host.approveClose) {
      await host.approveClose()
      return
    }
    await host.requestClose?.()
  }

  const handleWebFileAction = (event: Event): void => {
    const action = (event as CustomEvent<SheetsWebFileAction>).detail
    if (!action) return
    if (action === 'open' || action === 'save' || action === 'save-as') {
      dispatchRendererFileAction(action)
      return
    }
    if (action === 'save-history' || action === 'export-xlsx') {
      dispatchRendererFileAction(action)
      return
    }
    if (action === 'save-and-exit') {
      closeAfterSave = true
      dispatchRendererFileAction('save')
      return
    }
    if (action === 'discard-and-exit') {
      closeAfterSave = false
      void requestHostClose()
    }
  }
  window.addEventListener(SHEETS_WEB_FILE_ACTION_EVENT, handleWebFileAction)

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
    dispatchRendererFileAction('open')
  }

  const setActiveWorkbook = (workbook: WorkbookFile): void => {
    activeWorkbook = workbook
    pendingWorkbook = workbook
    currentTitle = workbook.name
    host.setTitle(workbook.name)
  }

  const setWorkbookFromOfficeFile = async (file: OfficeFile): Promise<void> => {
    const previous = activeWorkbook
    const workbook = await openXlsxWorkbookBytes(file.name, file.bytes.slice(0))
    currentOfficeFile = file
    currentIsNewDocument = false
    setActiveWorkbook(workbook)
    if (previous && previous.sessionId !== workbook.sessionId) {
      await deleteXlsxSession(previous.sessionId).catch(() => undefined)
    }
  }

  const pickWorkbook = async (): Promise<WorkbookFile | null> => {
    if (pendingWorkbook) {
      const workbook = pendingWorkbook
      pendingWorkbook = null
      return workbook
    }

    if (host.pickDocument && host.confirmDocumentOpened && host.releasePickedDocument) {
      const picked = await host.pickDocument({ accept: [XLSX_MIME, '.xlsx'] })
      if (picked.status === 'cancelled') return null

      let candidate: WorkbookFile | null = null
      let bound = false
      try {
        candidate = await openXlsxWorkbookBytes(picked.file.name, picked.file.bytes.slice(0))
        const bindResult = await host.confirmDocumentOpened(picked.selectionId)
        if (!bindResult.ok) {
          throw new Error(
            `${bindResult.code ?? 'FILE_BIND_FAILED'}: ${bindResult.error || 'The Host could not bind the selected workbook.'}`,
          )
        }
        bound = true

        const previous = activeWorkbook
        const descriptor = bindResult.file ?? officeDescriptor(picked.file, candidate)
        const nextWorkbook: WorkbookFile = {
          ...candidate,
          name: descriptor.name || candidate.name,
        }
        currentOfficeFile = {
          ...picked.file,
          ...descriptor,
          bytes: picked.file.bytes.slice(0),
          transport: 'buffer',
        }
        currentIsNewDocument = false
        activeWorkbook = nextWorkbook
        currentTitle = nextWorkbook.name
        host.setTitle(currentTitle)

        if (previous && previous.sessionId !== nextWorkbook.sessionId) {
          await deleteXlsxSession(previous.sessionId).catch(() => undefined)
        }
        return nextWorkbook
      } catch (error) {
        if (!bound) {
          await host.releasePickedDocument(picked.selectionId).catch(() => undefined)
        }
        if (candidate && candidate.sessionId !== activeWorkbook?.sessionId) {
          await deleteXlsxSession(candidate.sessionId).catch(() => undefined)
        }
        throw error
      }
    }

    // Protocol-v1 fallback. Parse the candidate before releasing the previous
    // workbook session so a corrupt/unsupported file cannot destroy the active
    // workbook. Embedded runtime translates this call to pick-document.
    const selected = await host.pickFile({
      multiple: false,
      accept: [XLSX_MIME, '.xlsx'],
      mode: 'file',
    })
    if (!selected?.[0]) return null
    const file = await selectedToOfficeFile(host, selected[0])
    const candidate = await openXlsxWorkbookBytes(file.name, file.bytes.slice(0))
    const previous = activeWorkbook
    currentOfficeFile = file
    currentIsNewDocument = false
    activeWorkbook = candidate
    currentTitle = file.name
    host.setTitle(file.name)
    if (previous && previous.sessionId !== candidate.sessionId) {
      await deleteXlsxSession(previous.sessionId).catch(() => undefined)
    }
    return candidate
  }

  const createNewWorkbook = async (): Promise<void> => {
    const previous = activeWorkbook
    const workbook = await createBlankXlsxWorkbook('Untitled.xlsx')
    currentOfficeFile = null
    currentIsNewDocument = true
    setActiveWorkbook(workbook)
    if (previous && previous.sessionId !== workbook.sessionId) {
      await deleteXlsxSession(previous.sessionId).catch(() => undefined)
    }
    emitOpen()
  }

  const materializeWorkbook = async (
    request: WorkbookSaveRequest,
  ): Promise<MaterializedWorkbook> => {
    if (currentMode !== 'edit') throw new Error('Workbook is read-only.')
    if (!activeWorkbook) throw new Error('No active workbook session.')

    if (!hasWorkbookMutations(request)) {
      return {
        ...(await saveXlsxArchiveMutation(request.sessionId, activeWorkbook.name, {
          replacements: new Map(),
          removals: [],
          additions: new Map(),
        })),
        touchedEntries: [],
      }
    }

    const hasShifts = request.structuralOps.length > 0 || request.sheetOps.length > 0
    const heldPivots = hasShifts ? request.pivotAdditions : []
    const heldTables = request.structuralOps.length > 0 ? request.tableAdditions : []
    const heldNames = hasShifts ? request.definedNamesState : null
    const split = heldPivots.length > 0 || heldTables.length > 0 || heldNames !== null

    if (!split) return saveWorkbookRequestViaEngine(request, activeWorkbook, activeWorkbook.name)

    const firstRequest: WorkbookSaveRequest = {
      ...request,
      mode: 'save',
      tableAdditions: heldTables.length > 0 ? [] : request.tableAdditions,
      pivotAdditions: heldPivots.length > 0 ? [] : request.pivotAdditions,
      definedNamesState: null,
    }
    const first = await saveWorkbookRequestViaEngine(
      firstRequest,
      activeWorkbook,
      activeWorkbook.name,
    )
    try {
      const second = await saveWorkbookRequestViaEngine(
        emptySecondPhaseRequest(first.file.sessionId, request, heldTables, heldPivots, heldNames),
        first.file,
        activeWorkbook.name,
      )
      if (first.file.sessionId !== second.file.sessionId) {
        await deleteXlsxSession(first.file.sessionId).catch(() => undefined)
      }
      return {
        ...second,
        touchedEntries: [...new Set([...first.touchedEntries, ...second.touchedEntries])],
      }
    } catch (error) {
      await deleteXlsxSession(first.file.sessionId).catch(() => undefined)
      throw error
    }
  }

  const snapshotHost: SheetsWebSnapshotHost = {
    saveHistoryVersion: async (request) => {
      if (!currentOfficeFile) {
        throw new Error('SAVE_FAILED: Save the workbook before creating a history version.')
      }
      if (!host.saveHistoryVersion) {
        throw new Error('SAVE_FAILED: The current Host does not support history versions.')
      }
      saving = true
      const materialized = await materializeWorkbook(request)
      try {
        const descriptor = officeDescriptor(currentOfficeFile, activeWorkbook ?? materialized.file)
        const result = await host.saveHistoryVersion({
          file: descriptor,
          bytes: materialized.bytes.slice(0),
          baseVersion: descriptor.version ?? null,
        })
        if (result.ok) {
          if (result.file) {
            currentOfficeFile = {
              ...currentOfficeFile,
              ...result.file,
              bytes: materialized.bytes.slice(0),
              transport: 'buffer',
            }
            currentTitle = result.file.name
            host.setTitle(currentTitle)
          }
          return { canceled: false }
        }
        if (result.code === 'CANCELLED') return { canceled: true }
        throw new Error(
          `${result.code ?? 'SAVE_FAILED'}: ${result.error || 'The Host could not save a history version.'}`,
        )
      } finally {
        saving = false
        await deleteXlsxSession(materialized.file.sessionId).catch(() => undefined)
      }
    },
    exportXlsx: async (request) => {
      if (!host.downloadDocument && !host.exportDocument) {
        throw new Error('SAVE_FAILED: The current Host does not support XLSX download.')
      }
      saving = true
      const materialized = await materializeWorkbook(request)
      try {
        const workbook = activeWorkbook ?? materialized.file
        const descriptor = officeDescriptor(currentOfficeFile, workbook)
        const input = {
          format: 'xlsx' as const,
          file: { ...descriptor, name: descriptor.name || 'Untitled.xlsx' },
          bytes: materialized.bytes.slice(0),
        }
        const result = host.downloadDocument
          ? await host.downloadDocument(input)
          : await host.exportDocument!(input)
        if (result.ok) return { canceled: false }
        if (result.code === 'CANCELLED') return { canceled: true }
        throw new Error(
          `${result.code ?? 'DOWNLOAD_FAILED'}: ${result.error || 'The Host could not download this workbook.'}`,
        )
      } finally {
        saving = false
        await deleteXlsxSession(materialized.file.sessionId).catch(() => undefined)
      }
    },
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
      dispatchRendererFileAction('save')
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

        if (!result.ok) {
          closeAfterSave = false
          await deleteXlsxSession(saved.file.sessionId).catch(() => undefined)
          if (result.code === 'CANCELLED') return { canceled: true }
          const code = result.code ?? 'SAVE_FAILED'
          throw new Error(`${code}: ${result.error || 'The host could not save this workbook.'}`)
        }
        if (!result.file) {
          closeAfterSave = false
          await deleteXlsxSession(saved.file.sessionId).catch(() => undefined)
          throw new Error('SAVE_FAILED: The host reported success without a saved file descriptor.')
        }

        const nextWorkbook: WorkbookFile = {
          ...saved.file,
          name: result.file.name,
        }
        activeWorkbook = nextWorkbook
        currentOfficeFile = {
          ...result.file,
          bytes: saved.bytes.slice(0),
          transport: 'buffer',
        }
        currentTitle = result.file.name
        currentIsNewDocument = false
        dirty = false
        host.setTitle(currentTitle)
        host.setDirty(false)
        if (previousSessionId !== nextWorkbook.sessionId) {
          await deleteXlsxSession(previousSessionId).catch(() => undefined)
        }

        const shouldClose = closeAfterSave
        closeAfterSave = false
        if (shouldClose) await requestHostClose()

        return {
          canceled: false,
          file: nextWorkbook,
          touchedEntries: [...saved.touchedEntries],
        }
      } catch (error) {
        closeAfterSave = false
        throw error
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
    snapshotHost,
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
      window.removeEventListener('keydown', handleWebFileShortcut)
      window.removeEventListener(SHEETS_WEB_FILE_ACTION_EVENT, handleWebFileAction)
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
