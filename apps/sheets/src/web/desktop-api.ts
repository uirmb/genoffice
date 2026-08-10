import type { AiChatResponse, AiSettings, AiStreamChunk, GenSparkAccountStatus } from '@genoffice/ai-provider'
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
import { deleteXlsxSession, openXlsxWorkbook, readXlsxWorkbookRange } from './engine-client'

function unavailable(name: string): never {
  throw new Error(`${name} is not available in Sheets Web yet.`)
}

function noopUnsubscribe(): () => void {
  return () => undefined
}

function pickLocalXlsx(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    input.style.display = 'none'
    document.body.append(input)

    let settled = false
    const finish = (file: File | null) => {
      if (settled) return
      settled = true
      input.remove()
      resolve(file)
    }

    input.addEventListener(
      'change',
      () => {
        const file = input.files?.[0] ?? null
        finish(file && /\.xlsx$/i.test(file.name) ? file : null)
      },
      { once: true },
    )
    window.addEventListener(
      'focus',
      () => {
        window.setTimeout(() => {
          if (!input.files?.length) finish(null)
        }, 250)
      },
      { once: true },
    )
    input.click()
  })
}

export function createSheetsWebDesktopApi(): DesktopApi {
  return {
    getLanguage: async () => 'zh',
    onLanguageChanged: () => noopUnsubscribe(),

    // Standalone Web uses a browser picker only as a development fallback.
    // Embedded UC Excel will replace the file source with office:pick-file while
    // keeping the same Rust workbook transport and renderer contract.
    selectWorkbook: async (): Promise<WorkbookFile | null> => {
      const file = await pickLocalXlsx()
      return file ? openXlsxWorkbook(file) : null
    },
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
    saveWorkbookEdits: async (_request: WorkbookSaveRequest): Promise<WorkbookSaveResult> => ({
      canceled: true,
    }),
    writeWorkbookRecovery: async () => ({ ok: true }),
    autoRenameWorkbook: async () => ({ renamed: false }),
    exportPdf: async (): Promise<WorkbookExportPdfResult> => ({ canceled: true }),
    closeWorkbook: deleteXlsxSession,
    openExternal: async (url: string) => {
      if (/^https?:\/\//.test(url)) window.open(url, '_blank', 'noopener,noreferrer')
    },

    onMenuAction: () => noopUnsubscribe(),
    onWorkbookRenamed: () => noopUnsubscribe(),
    notifyPendingEdits: () => undefined,
    onCloseSaveRequest: () => noopUnsubscribe(),
    reportCloseSaveResult: () => undefined,
    consumeNewBlankWorkbook: async () => false,
    hasQueuedWorkbook: async () => false,

    // AI is platform-disabled in the Web Office product policy. Keep a valid
    // bridge surface so the existing renderer does not need Web-only forks.
    getAiSettings: async () => ({ provider: '', providers: {} }) as unknown as AiSettings,
    setAiSettings: async () => undefined,
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
}
