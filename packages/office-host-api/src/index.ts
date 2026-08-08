export type OfficeDocumentKind = 'docx' | 'pptx' | 'xlsx'
export type OfficeEditorMode = 'view' | 'edit'
export type OfficeSaveMode = 'save' | 'saveAs'
export type OfficeAutoSavePolicy = 'disabled' | 'host' | 'editor'

export interface OfficeHostCapabilities {
  ai: boolean
  open: boolean
  save: boolean
  saveAs: boolean
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
}

export interface SaveDocumentResult {
  ok: boolean
  file?: OfficeFileDescriptor
  error?: string
  code?:
    | 'VERSION_CONFLICT'
    | 'PERMISSION_DENIED'
    | 'NOT_FOUND'
    | 'SAVE_FAILED'
    | 'CANCELLED'
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
  pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null>
  readFile(fileId: string): Promise<OfficeFile>
  setDirty(dirty: boolean): void
  setTitle(title: string): void
}
