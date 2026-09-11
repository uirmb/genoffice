/** Toast event bus, split from the ToastHost component so this module has no
 * component exports: React Fast Refresh replaces mixed-export modules wholesale,
 * which strands callers (file-actions) on a stale emitter during dev HMR. */
import type {
  OfficeNotificationCode,
  OfficeNotificationLevel,
  OfficeNotificationOperation,
} from '@genoffice/office-protocol'
import { notifyOffice } from '@genoffice/web-runtime'

export interface ToastData {
  text: string
  kind: 'success' | 'error'
}

export interface ToastNotificationOptions {
  code: OfficeNotificationCode
  operation?: OfficeNotificationOperation | undefined
  dedupeKey?: string | undefined
  durationMs?: number | undefined
  requestId?: string | undefined
  level?: OfficeNotificationLevel | undefined
}

let emit: ((toast: ToastData) => void) | null = null

/** Registered by ToastHost on mount; null while unmounted. */
export function setToastEmitter(fn: ((toast: ToastData) => void) | null): void {
  emit = fn
}

/**
 * The editor decides what happened; the negotiated runtime decides where it is
 * displayed. Calls without protocol metadata remain editor-local so unrelated
 * legacy feedback is not silently promoted to Host notifications.
 */
export function showToast(
  text: string,
  kind: 'success' | 'error' = 'success',
  notification?: ToastNotificationOptions,
): void {
  if (notification) {
    const { level, ...metadata } = notification
    if (notifyOffice({ ...metadata, level: level ?? kind, message: text })) return
  }
  emit?.({ text, kind })
}
