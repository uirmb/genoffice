import type { OfficeHostApi } from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type { EditorIframeBridge } from '@genoffice/web-runtime'

import {
  SHEETS_WEB_FILE_ACTION_EVENT,
  SHEETS_WEB_HOST_CLOSE_REQUEST_EVENT,
  type SheetsWebFileAction,
} from './file-actions'

export interface SheetsWebCloseLifecycle {
  host: OfficeHostApi
  destroy(): void
}

function createCloseAwareHost(
  baseHost: OfficeHostApi,
  bridge: EditorIframeBridge,
  getPendingRequestId: () => string | null,
  clearPendingRequest: () => void,
): OfficeHostApi {
  const approveClose = async (explicitRequestId?: string): Promise<void> => {
    const requestId = explicitRequestId ?? getPendingRequestId()
    if (!requestId) {
      if (baseHost.approveClose) await baseHost.approveClose()
      else await baseHost.requestClose?.()
      return
    }

    clearPendingRequest()
    if (baseHost.approveClose) {
      await baseHost.approveClose(requestId)
      return
    }
    bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:close-approved',
      requestId,
      payload: { reason: 'window-close' },
    })
  }

  const cancelClose = async (requestId: string): Promise<void> => {
    clearPendingRequest()
    if (baseHost.cancelClose) {
      await baseHost.cancelClose(requestId)
      return
    }
    bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:close-cancelled',
      requestId,
      payload: { reason: 'user-cancelled' },
    })
  }

  return {
    getLocale: () => baseHost.getLocale(),
    saveDocument: (input) => baseHost.saveDocument(input),
    ...(baseHost.saveHistoryVersion
      ? { saveHistoryVersion: (input) => baseHost.saveHistoryVersion!(input) }
      : {}),
    ...(baseHost.downloadDocument
      ? { downloadDocument: (input) => baseHost.downloadDocument!(input) }
      : {}),
    ...(baseHost.exportDocument
      ? { exportDocument: (input) => baseHost.exportDocument!(input) }
      : {}),
    ...(baseHost.pickDocument
      ? { pickDocument: (options) => baseHost.pickDocument!(options) }
      : {}),
    ...(baseHost.confirmDocumentOpened
      ? { confirmDocumentOpened: (selectionId) => baseHost.confirmDocumentOpened!(selectionId) }
      : {}),
    ...(baseHost.releasePickedDocument
      ? { releasePickedDocument: (selectionId) => baseHost.releasePickedDocument!(selectionId) }
      : {}),
    ...(baseHost.pickAssets ? { pickAssets: (options) => baseHost.pickAssets!(options) } : {}),
    approveClose,
    cancelClose,
    // v1 compatibility for renderer code that still calls requestClose().
    requestClose: () => approveClose(),
    pickFile: (options) => baseHost.pickFile(options),
    readFile: (fileId) => baseHost.readFile(fileId),
    setDirty: (dirty) => baseHost.setDirty(dirty),
    setTitle: (title) => baseHost.setTitle(title),
  }
}

export function createSheetsWebCloseLifecycle(
  baseHost: OfficeHostApi,
  bridge: EditorIframeBridge,
): SheetsWebCloseLifecycle {
  let pendingHostCloseRequestId: string | null = null

  const clearPendingRequest = (): void => {
    pendingHostCloseRequestId = null
  }

  const host = createCloseAwareHost(
    baseHost,
    bridge,
    () => pendingHostCloseRequestId,
    clearPendingRequest,
  )

  const unsubscribeBridge = bridge.subscribe((message: HostToEditorMessage) => {
    if (message.type !== 'office:request-close') return

    // Only one Host window-close transaction may be active at a time. Repeated
    // clicks on the UC window close button are intentionally ignored until the
    // current Excel exit flow either grants or cancels the transaction.
    if (pendingHostCloseRequestId !== null) return

    pendingHostCloseRequestId = message.requestId
    window.dispatchEvent(new Event(SHEETS_WEB_HOST_CLOSE_REQUEST_EVENT))
  })

  const handleFileAction = (event: Event): void => {
    const action = (event as CustomEvent<SheetsWebFileAction>).detail
    if (action !== 'cancel-exit' || pendingHostCloseRequestId === null) return

    const requestId = pendingHostCloseRequestId
    void host.cancelClose?.(requestId)
  }
  window.addEventListener(SHEETS_WEB_FILE_ACTION_EVENT, handleFileAction)

  return {
    host,
    destroy: () => {
      unsubscribeBridge()
      window.removeEventListener(SHEETS_WEB_FILE_ACTION_EVENT, handleFileAction)
      clearPendingRequest()
    },
  }
}
