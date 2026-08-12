import type { WorkbookSaveRequest } from '../shared/desktop-api'

export type SheetsWebFileAction =
  | 'open'
  | 'save'
  | 'save-as'
  | 'save-history'
  | 'export-xlsx'
  | 'save-and-exit'
  | 'discard-and-exit'

export const SHEETS_WEB_FILE_ACTION_EVENT = 'genoffice:sheets-web-file-action'

export interface SheetsWebSnapshotActionResult {
  canceled: boolean
}

export interface SheetsWebSnapshotHost {
  saveHistoryVersion(request: WorkbookSaveRequest): Promise<SheetsWebSnapshotActionResult>
  exportXlsx(request: WorkbookSaveRequest): Promise<SheetsWebSnapshotActionResult>
}

const SNAPSHOT_HOST_KEY = '__genofficeSheetsWebSnapshotHost'

type SnapshotWindow = Window & {
  [SNAPSHOT_HOST_KEY]?: SheetsWebSnapshotHost
}

export function dispatchSheetsWebFileAction(action: SheetsWebFileAction): void {
  window.dispatchEvent(new CustomEvent<SheetsWebFileAction>(SHEETS_WEB_FILE_ACTION_EVENT, { detail: action }))
}

export function installSheetsWebSnapshotHost(host: SheetsWebSnapshotHost): () => void {
  const target = window as SnapshotWindow
  Object.defineProperty(target, SNAPSHOT_HOST_KEY, {
    configurable: true,
    value: host,
  })
  return () => {
    if (target[SNAPSHOT_HOST_KEY] === host) delete target[SNAPSHOT_HOST_KEY]
  }
}

export function getSheetsWebSnapshotHost(): SheetsWebSnapshotHost | null {
  return (window as SnapshotWindow)[SNAPSHOT_HOST_KEY] ?? null
}
