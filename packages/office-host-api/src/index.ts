export type OfficeDocumentKind = 'docx' | 'pptx' | 'xlsx'
export type OfficeEditorMode = 'view' | 'edit'
export type OfficeSaveMode = 'save' | 'saveAs'
export type OfficeExportFormat = 'docx' | 'pptx' | 'xlsx'
export type OfficeAutoSavePolicy = 'disabled' | 'host' | 'editor'

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
}

export interface OfficeFileDescriptor {
  id: string
  name: string
  mimeType: string
  size?: number
  version?: string | null
}

export interface OfficeFile extends OfficeFileDescriptor {
  bytes: ArrayBuffer
}

export interface SaveDocumentInput {
  file: OfficeFileDescriptor
  bytes: ArrayBuffer
  baseVersion?: string | null
  mode?: OfficeSaveMode
  /** First persistence of a blank editor document; the Host should choose/create its destination. */
  newDocument?: boolean
}

export interface SaveDocumentResult {
  ok: boolean
  file?: OfficeFileDescriptor
  error?: string
  code?: 'VERSION_CONFLICT' | 'PERMISSION_DENIED' | 'NOT_FOUND' | 'SAVE_FAILED' | 'CANCELLED'
}

export interface SaveHistoryVersionInput {
  file: OfficeFileDescriptor
  bytes: ArrayBuffer
  baseVersion?: string | null
}

export interface SaveHistoryVersionResult {
  ok: boolean
  error?: string
  code?: 'PERMISSION_DENIED' | 'NOT_FOUND' | 'SAVE_FAILED' | 'CANCELLED'
}

export interface ExportDocumentInput {
  format: OfficeExportFormat
  file: OfficeFileDescriptor
  bytes: ArrayBuffer
}

export interface ExportDocumentResult {
  ok: boolean
  error?: string
  code?: 'PERMISSION_DENIED' | 'SAVE_FAILED' | 'CANCELLED'
}

export interface PickFileOptions {
  multiple?: boolean
  accept?: string[]
  mode?: 'file' | 'folder'
}

export interface SelectedOfficeFile extends OfficeFileDescriptor {
  transport: 'buffer' | 'token'
  bytes?: ArrayBuffer
  token?: string
  url?: string
}

export interface OfficeHostApi {
  getLocale(): Promise<string>
  saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult>
  saveHistoryVersion?(input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult>
  exportDocument?(input: ExportDocumentInput): Promise<ExportDocumentResult>
  requestClose?(): Promise<void>
  pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null>
  readFile(fileId: string): Promise<OfficeFile>
  setDirty(dirty: boolean): void
  setTitle(title: string): void
}
