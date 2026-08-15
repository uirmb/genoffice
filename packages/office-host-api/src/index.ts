export type OfficeDocumentKind = 'docx' | 'pptx' | 'xlsx'
export type OfficeEditorMode = 'view' | 'edit'
export type OfficeSaveMode = 'save' | 'saveAs'
export type OfficeExportFormat = 'docx' | 'pptx' | 'xlsx'
export type OfficeAutoSavePolicy = 'disabled' | 'host' | 'editor'

/**
 * File versions are opaque to GenOffice. UC may expose a numeric revision while
 * standalone/demo hosts may use a string token. Editors must only echo/compare
 * values supplied by the Host and must never derive identity from file names.
 */
export type OfficeFileVersion = string | number | null

export interface OfficeHostCapabilities {
  ai: boolean
  open: boolean
  save: boolean
  saveAs: boolean
  saveHistoryVersion: boolean
  exportDocx: boolean
  exportPptx: boolean
  exportXlsx: boolean
  close: boolean
  autoSave: OfficeAutoSavePolicy
  download: boolean
  print: boolean
  systemFilePicker: boolean
  pageCropMarks: boolean
  /** New stable capability: choose a document transactionally before binding it. */
  openDocument?: boolean | undefined
  /** New stable capability: choose read-only UC assets for insertion. */
  pickAssets?: boolean | undefined
}

export const DEFAULT_STANDALONE_OFFICE_CAPABILITIES: OfficeHostCapabilities = {
  ai: false,
  open: true,
  save: true,
  saveAs: true,
  saveHistoryVersion: false,
  exportDocx: true,
  exportPptx: true,
  exportXlsx: true,
  close: false,
  autoSave: 'disabled',
  download: true,
  print: true,
  systemFilePicker: true,
  pageCropMarks: true,
  openDocument: true,
  pickAssets: true,
}

export const DEFAULT_EMBEDDED_OFFICE_CAPABILITIES: OfficeHostCapabilities = {
  ai: false,
  open: true,
  save: true,
  saveAs: true,
  saveHistoryVersion: true,
  exportDocx: true,
  exportPptx: true,
  exportXlsx: true,
  close: true,
  autoSave: 'host',
  download: false,
  print: true,
  systemFilePicker: true,
  pageCropMarks: true,
  openDocument: true,
  pickAssets: true,
}

export interface OfficeFileDescriptor {
  /** Stable file identity. In UC this is normally the FsNode/nodeId value. */
  id: string
  /** Optional explicit UC identity while legacy callers still consume id. */
  nodeId?: string | undefined
  tenantId?: string | undefined
  /** Parent FsNode identity when the Host exposes it. */
  parentId?: string | null | undefined
  name: string
  mimeType: string
  /** Canonical descriptors always carry the byte size represented by this revision. */
  size: number
  /** Canonical descriptors always carry a Host-supplied version, or null when none exists yet. */
  version: OfficeFileVersion
  /** Host timestamp for the represented revision, normally an ISO-8601 string. */
  updatedAt?: string | undefined
  /** Stable office:* file transport is always an in-memory buffer. */
  transport: 'buffer'
}

export interface OfficeFile extends OfficeFileDescriptor {
  bytes: ArrayBuffer
}

export interface SaveDocumentInput {
  file: OfficeFileDescriptor
  bytes: ArrayBuffer
  baseVersion?: OfficeFileVersion | undefined
  mode?: OfficeSaveMode | undefined
  /** First persistence of a blank editor document; the Host should choose/create its destination. */
  newDocument?: boolean | undefined
}

export interface SaveDocumentResult {
  ok: boolean
  file?: OfficeFileDescriptor | undefined
  error?: string | undefined
  /** Preserve platform error codes instead of collapsing all failures to SAVE_FAILED. */
  code?: string | undefined
}

export interface SaveHistoryVersionInput {
  file: OfficeFileDescriptor
  bytes: ArrayBuffer
  baseVersion?: OfficeFileVersion | undefined
}

export interface SaveHistoryVersionResult {
  ok: boolean
  /** Successful history creation returns the actual latest descriptor/version. */
  file?: OfficeFileDescriptor | undefined
  error?: string | undefined
  code?: string | undefined
}

export interface DownloadDocumentInput {
  format: OfficeExportFormat
  file: OfficeFileDescriptor
  bytes: ArrayBuffer
}

export interface DownloadDocumentResult {
  ok: boolean
  error?: string | undefined
  code?: string | undefined
}

/** @deprecated Use DownloadDocumentInput. */
export type ExportDocumentInput = DownloadDocumentInput
/** @deprecated Use DownloadDocumentResult. */
export type ExportDocumentResult = DownloadDocumentResult

export interface PickDocumentOptions {
  accept?: string[] | undefined
}

export interface PickAssetsOptions {
  multiple?: boolean | undefined
  accept?: string[] | undefined
}

export interface SelectedOfficeDocument {
  status: 'selected'
  selectionId: string
  file: OfficeFile
}

export interface CancelledOfficeDocumentSelection {
  status: 'cancelled'
  selectionId: null
  file: null
}

export interface FailedOfficeDocumentSelection {
  status: 'failed'
  code: string
  error: string
}

export type PickDocumentResult =
  SelectedOfficeDocument | CancelledOfficeDocumentSelection | FailedOfficeDocumentSelection

export interface SelectedOfficeAssets {
  status: 'selected'
  files: OfficeFile[]
}

export interface CancelledOfficeAssetsSelection {
  status: 'cancelled'
  files: []
}

export interface FailedOfficeAssetsSelection {
  status: 'failed'
  files: []
  code: string
  error: string
}

export type PickAssetsResult =
  SelectedOfficeAssets | CancelledOfficeAssetsSelection | FailedOfficeAssetsSelection

export interface DocumentOpenedResult {
  ok: boolean
  file?: OfficeFileDescriptor | undefined
  error?: string | undefined
  code?: string | undefined
}

/** Legacy generic picker retained only for compatibility during the v1 migration. */
export interface PickFileOptions {
  multiple?: boolean | undefined
  accept?: string[] | undefined
  mode?: 'file' | 'folder' | undefined
}

/** Legacy token transport retained only for old callers. New Office flows use buffer bytes only. */
export type SelectedOfficeFile = Omit<OfficeFileDescriptor, 'transport'> & {
  transport: 'buffer' | 'token'
  bytes?: ArrayBuffer | undefined
  token?: string | undefined
  url?: string | undefined
}

export interface OfficeHostApi {
  getLocale(): Promise<string>

  saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult>
  saveHistoryVersion?(input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult>
  downloadDocument?(input: DownloadDocumentInput): Promise<DownloadDocumentResult>

  /** Transactional document open. Selecting a file does not bind it as current. */
  pickDocument?(options: PickDocumentOptions): Promise<PickDocumentResult>
  /** Called only after the editor has successfully loaded the selected bytes. */
  confirmDocumentOpened?(selectionId: string): Promise<DocumentOpenedResult>
  /** Releases a pending selection after editor load failure/cancellation. */
  releasePickedDocument?(selectionId: string): Promise<void>
  /** Read-only asset selection for insertion; never changes the current document. */
  pickAssets?(options: PickAssetsOptions): Promise<PickAssetsResult>

  /** Editor grants a Host close request or initiates File -> Exit. */
  approveClose?(requestId?: string): Promise<void>
  cancelClose?(requestId: string): Promise<void>

  /** @deprecated Compatibility alias for downloadDocument. */
  exportDocument?(input: ExportDocumentInput): Promise<ExportDocumentResult>
  /** @deprecated Compatibility alias for approveClose. */
  requestClose?(): Promise<void>
  /** @deprecated Use pickDocument/pickAssets. */
  pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null>
  /** @deprecated New stable flows always receive bytes in the picker result. */
  readFile(fileId: string): Promise<OfficeFile>

  setDirty(dirty: boolean): void
  setTitle(title: string): void
}
