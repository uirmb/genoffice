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

function unavailable(name: string): never {
  throw new Error(`${name} is not available in Sheets Web yet.`)
}

function noopUnsubscribe(): () => void {
  return () => undefined
}

export function createSheetsWebDesktopApi(): DesktopApi {
  return {
    getLanguage: async () => 'zh',
    onLanguageChanged: () => noopUnsubscribe(),

    // Workbook transport is intentionally isolated behind DesktopApi so the
    // existing React + Univer renderer can run unchanged. These methods are
    // replaced incrementally with xlsx-engine-service calls in the next slice.
    selectWorkbook: async (): Promise<WorkbookFile | null> => null,
    readWorkbookRange: async (_request: WorkbookRangeRequest): Promise<WorkbookRangeResult> =>
      unavailable('readWorkbookRange'),
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
    closeWorkbook: async () => undefined,
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
    getAiSettings: async () => ({ provider: '', providers: {} }) as AiSettings,
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
