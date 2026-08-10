import type {
  OfficeDocumentKind,
  OfficeEditorMode,
  OfficeExportFormat,
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostCapabilities,
  OfficeSaveMode,
  ExportDocumentResult,
  PickFileOptions,
  SaveDocumentResult,
  SaveHistoryVersionResult,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'

export const OFFICE_PROTOCOL_VERSION = 1 as const

export interface OfficeInitPayload {
  kind: OfficeDocumentKind
  mode: OfficeEditorMode
  locale?: string
  file: OfficeFile
  capabilities?: Partial<OfficeHostCapabilities>
}

export interface OfficeNewPayload {
  kind: OfficeDocumentKind
  mode: OfficeEditorMode
  locale?: string
  capabilities?: Partial<OfficeHostCapabilities>
}

export interface OfficeEditorState {
  ready: boolean
  dirty: boolean
  saving: boolean
  mode: OfficeEditorMode
  title?: string
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
      requestId?: string
      payload: { locale: string }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:set-mode'
      requestId?: string
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
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:error'
      requestId?: string
      payload: OfficeProtocolErrorPayload
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
      payload: { ok: boolean; error?: string }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-document'
      requestId: string
      payload: {
        file: OfficeFileDescriptor
        bytes: ArrayBuffer
        baseVersion?: string | null
        mode?: OfficeSaveMode
        newDocument?: boolean
      }
    }
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:save-history-version'
      requestId: string
      payload: {
        file: OfficeFileDescriptor
        bytes: ArrayBuffer
        baseVersion?: string | null
      }
    }
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
      requestId?: string
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
  | {
      protocol: typeof OFFICE_PROTOCOL_VERSION
      type: 'office:error'
      requestId?: string
      payload: OfficeProtocolErrorPayload
    }

export type OfficeProtocolMessage = HostToEditorMessage | EditorToHostMessage

export function isOfficeProtocolMessage(value: unknown): value is OfficeProtocolMessage {
  if (!value || typeof value !== 'object') return false
  const message = value as { protocol?: unknown; type?: unknown }
  return message.protocol === OFFICE_PROTOCOL_VERSION && typeof message.type === 'string'
}
