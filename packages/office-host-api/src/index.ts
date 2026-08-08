export type OfficeDocumentKind = 'docx' | 'pptx' | 'xlsx'
export type OfficeEditorMode = 'view' | 'edit'

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
}

export interface SaveDocumentResult {
  ok: boolean
  file?: OfficeFileDescriptor
  error?: string
  code?: 'VERSION_CONFLICT' | 'PERMISSION_DENIED' | 'NOT_FOUND' | 'SAVE_FAILED'
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
