import type {
  DocumentOpenedResult,
  DownloadDocumentResult,
  ExportDocumentResult,
  OfficeDocumentKind,
  OfficeEditorMode,
  OfficeExportFormat,
  OfficeFile,
  OfficeFileDescriptor,
  OfficeFileVersion,
  OfficeHostCapabilities,
  OfficeSaveMode,
  PickAssetsOptions,
  PickAssetsResult,
  PickDocumentOptions,
  PickDocumentResult,
  PickFileOptions,
  SaveDocumentResult,
  SaveHistoryVersionResult,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'

/**
 * Protocol v1 is intentionally retained while the stable file lifecycle is
 * introduced additively. Legacy pick-file/read-file/export-document and
 * close-request messages remain compatibility aliases until all callers move.
 */
export const OFFICE_PROTOCOL_VERSION = 1 as const

export interface OfficeInitPayload {
  kind: OfficeDocumentKind
  mode: OfficeEditorMode
  locale?: string | undefined
  /** Initial content is already materialized; the editor must not read it again. */
  file: OfficeFile
  capabilities?: Partial<OfficeHostCapabilities> | undefined
}

export interface OfficeNewPayload {
  kind: OfficeDocumentKind
  mode: OfficeEditorMode
  locale?: string | undefined
  capabilities?: Partial<OfficeHostCapabilities> | undefined
}

export interface OfficeEditorState {
  ready: boolean
  dirty: boolean
  saving: boolean
  mode: OfficeEditorMode
  title?: string | undefined
}

export interface OfficeProtocolErrorPayload {
  code: string
  message: string
}

export type HostToEditorMessage =
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:init'
      requestId: string
      payload: OfficeInitPayload
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:new'
      requestId: string
      payload: OfficeNewPayload
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:set-locale'
      requestId?: string | undefined
      payload: { locale: string }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:set-mode'
      requestId?: string | undefined
      payload: { mode: OfficeEditorMode }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save'
      requestId: string
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:request-close'
      requestId: string
      payload: { reason: 'window-close' }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:query-state'
      requestId: string
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:pick-document-result'
      requestId: string
      payload: PickDocumentResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:document-opened-result'
      requestId: string
      payload: DocumentOpenedResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:pick-assets-result'
      requestId: string
      payload: PickAssetsResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-document-result'
      requestId: string
      payload: SaveDocumentResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-history-version-result'
      requestId: string
      payload: SaveHistoryVersionResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:download-document-result'
      requestId: string
      payload: DownloadDocumentResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:error'
      requestId?: string | undefined
      payload: OfficeProtocolErrorPayload
    }
  // ---- v1 compatibility aliases ----
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:export-document-result'
      requestId: string
      payload: ExportDocumentResult
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:pick-file-result'
      requestId: string
      payload: { files: SelectedOfficeFile[] | null }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:read-file-result'
      requestId: string
      payload: { file: OfficeFile }
    }

export type EditorToHostMessage =
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:ready'
      payload: { kind: OfficeDocumentKind }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:dirty-change'
      payload: { dirty: boolean }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:title-change'
      payload: { title: string }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:state-result'
      requestId: string
      payload: OfficeEditorState
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-result'
      requestId: string
      payload: { ok: boolean; error?: string | undefined }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:pick-document'
      requestId: string
      payload: PickDocumentOptions
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:document-opened'
      requestId: string
      payload: { selectionId: string }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:document-open-failed'
      requestId: string
      payload: { selectionId: string; code?: string | undefined; message?: string | undefined }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:pick-assets'
      requestId: string
      payload: PickAssetsOptions
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-document'
      requestId: string
      payload: {
        file: OfficeFileDescriptor
        bytes: ArrayBuffer
        baseVersion?: OfficeFileVersion | undefined
        mode?: OfficeSaveMode | undefined
        newDocument?: boolean | undefined
      }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-history-version'
      requestId: string
      payload: {
        file: OfficeFileDescriptor
        bytes: ArrayBuffer
        baseVersion?: OfficeFileVersion | undefined
      }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:download-document'
      requestId: string
      payload: {
        format: OfficeExportFormat
        file: OfficeFileDescriptor
        bytes: ArrayBuffer
      }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:close-approved'
      requestId: string
      payload: { reason: 'file-menu' | 'window-close' }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:close-cancelled'
      requestId: string
      payload: { reason: 'user-cancelled' }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:error'
      requestId?: string | undefined
      payload: OfficeProtocolErrorPayload
    }
  // ---- v1 compatibility aliases ----
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:export-document'
      requestId: string
      payload: {
        format: OfficeExportFormat
        file: OfficeFileDescriptor
        bytes: ArrayBuffer
      }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:close-request'
      requestId?: string | undefined
      payload: { reason: 'file-menu' | 'window-close' }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:pick-file'
      requestId: string
      payload: PickFileOptions
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:read-file'
      requestId: string
      payload: { fileId: string }
    }

export type OfficeProtocolMessage = HostToEditorMessage | EditorToHostMessage

export function isOfficeProtocolMessage(value: unknown): value is OfficeProtocolMessage {
  if (!value || typeof value !== 'object') return false
  const message = value as { protocol?: unknown; type?: unknown }
  return message.protocol === OFFICE_PROTOCOL_VERSION && typeof message.type === 'string'
}
