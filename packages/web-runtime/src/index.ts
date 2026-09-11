import { OfficeIframeBridge, createOfficeRequestId } from '@genoffice/iframe-bridge'
import type {
  DocumentOpenedResult,
  DownloadDocumentInput,
  DownloadDocumentResult,
  ExportDocumentInput,
  ExportDocumentResult,
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostApi,
  PickAssetsOptions,
  PickAssetsResult,
  PickDocumentOptions,
  PickDocumentResult,
  PickFileOptions,
  SaveDocumentInput,
  SaveDocumentResult,
  SaveHistoryVersionInput,
  SaveHistoryVersionResult,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  type EditorToHostMessage,
  type HostToEditorMessage,
  type OfficeNotificationLevel,
  type OfficeNotificationPayload,
  type OfficeNotificationSettings,
} from '@genoffice/office-protocol'

export type WebRuntimeMode = 'standalone' | 'embedded'
export type EditorIframeBridge = OfficeIframeBridge<HostToEditorMessage, EditorToHostMessage>

export function detectWebRuntimeMode(currentWindow: Window = window): WebRuntimeMode {
  return currentWindow.parent === currentWindow ? 'standalone' : 'embedded'
}

export type OfficeNotificationInput = Omit<OfficeNotificationPayload, 'id'> & {
  id?: string | undefined
  requestId?: string | undefined
}

const ALL_NOTIFICATION_LEVELS: readonly OfficeNotificationLevel[] = [
  'success',
  'info',
  'warning',
  'error',
]

/**
 * Notification ownership is negotiated only by office:init / office:new.
 * Embedded mode therefore defaults to editor-owned notifications until a
 * trusted Host explicitly opts into office:notification.
 */
export class OfficeNotificationAdapter {
  private owner: 'host' | 'editor' = 'editor'
  private transport: 'office:notification' | null = null
  private levels = new Set<OfficeNotificationLevel>(ALL_NOTIFICATION_LEVELS)

  constructor(private readonly bridge?: EditorIframeBridge) {}

  configure(settings?: OfficeNotificationSettings): void {
    if (!settings || settings.owner !== 'host' || settings.transport !== 'office:notification') {
      this.owner = 'editor'
      this.transport = null
      this.levels = new Set(ALL_NOTIFICATION_LEVELS)
      return
    }

    this.owner = 'host'
    this.transport = settings.transport
    this.levels = new Set(settings.levels ?? ALL_NOTIFICATION_LEVELS)
  }

  notify(notification: OfficeNotificationInput): boolean {
    if (
      !this.bridge ||
      this.owner !== 'host' ||
      this.transport !== 'office:notification' ||
      !this.levels.has(notification.level)
    ) {
      return false
    }

    const { requestId, id, ...payload } = notification
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:notification',
      ...(requestId ? { requestId } : {}),
      payload: {
        id: id || createOfficeRequestId('notification'),
        ...payload,
      },
    })
    return true
  }
}

let activeOfficeNotificationAdapter: OfficeNotificationAdapter | null = null

function activateOfficeNotificationAdapter(adapter: OfficeNotificationAdapter): () => void {
  activeOfficeNotificationAdapter = adapter
  return () => {
    if (activeOfficeNotificationAdapter === adapter) activeOfficeNotificationAdapter = null
  }
}

/**
 * Try to delegate a notification to the negotiated Host owner.
 * Returns false when the editor owns display, so existing editor Toasts can
 * render locally without knowing whether they are standalone or embedded.
 */
export function notifyOffice(notification: OfficeNotificationInput): boolean {
  return activeOfficeNotificationAdapter?.notify(notification) ?? false
}

const localNotificationTimers = new WeakMap<HTMLElement, number>()

function showLocalWebOfficeNotification(notification: OfficeNotificationInput): void {
  const rootId = 'genoffice-web-notification-root'
  let root = document.getElementById(rootId)
  if (!root) {
    root = document.createElement('div')
    root.id = rootId
    root.setAttribute('aria-live', 'polite')
    Object.assign(root.style, {
      position: 'fixed',
      top: '14px',
      left: '50%',
      transform: 'translateX(-50%)',
      zIndex: '2147483646',
      display: 'flex',
      flexDirection: 'column',
      gap: '8px',
      alignItems: 'center',
      pointerEvents: 'none',
      maxWidth: 'min(680px, calc(100vw - 32px))',
    })
    document.body.append(root)
  }

  const dedupeKey = notification.dedupeKey || notification.code
  const existing = [...root.children].find(
    (node) => (node as HTMLElement).dataset.dedupeKey === dedupeKey,
  ) as HTMLElement | undefined
  const toast = existing ?? document.createElement('div')
  if (!existing) root.append(toast)

  const previousTimer = localNotificationTimers.get(toast)
  if (previousTimer !== undefined) window.clearTimeout(previousTimer)

  const palette = {
    success: { background: '#ecfdf3', border: '#abefc6', color: '#067647' },
    info: { background: '#eff8ff', border: '#b2ddff', color: '#175cd3' },
    warning: { background: '#fffaeb', border: '#fedf89', color: '#b54708' },
    error: { background: '#fef3f2', border: '#fecdca', color: '#b42318' },
  }[notification.level]

  toast.dataset.dedupeKey = dedupeKey
  toast.textContent = notification.message
  toast.setAttribute('role', notification.level === 'error' ? 'alert' : 'status')
  Object.assign(toast.style, {
    pointerEvents: 'auto',
    padding: '10px 14px',
    borderRadius: '8px',
    border: `1px solid ${palette.border}`,
    background: palette.background,
    color: palette.color,
    boxShadow: '0 8px 24px rgba(16, 24, 40, 0.12)',
    font: '13px/1.45 system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    overflowWrap: 'anywhere',
  })

  const duration =
    notification.durationMs ??
    (notification.level === 'error' || notification.level === 'warning' ? 4000 : 2500)
  const timer = window.setTimeout(
    () => {
      localNotificationTimers.delete(toast)
      toast.remove()
      if (root && root.childElementCount === 0) root.remove()
    },
    Math.max(1000, duration),
  )
  localNotificationTimers.set(toast, timer)
}

/**
 * Web-only convenience for editors that do not already have a Toast component.
 * In UC it delegates to the Host when negotiated; standalone Web falls back to
 * a lightweight editor-owned Toast. Electron callers without an active Web
 * runtime are intentionally left untouched.
 */
export function showWebOfficeNotification(notification: OfficeNotificationInput): boolean {
  if (!activeOfficeNotificationAdapter) return false
  if (activeOfficeNotificationAdapter.notify(notification)) return true
  showLocalWebOfficeNotification(notification)
  return true
}

function randomSuffix(): string {
  return typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function localFileId(file: File): string {
  return `local:${file.name}:${file.size}:${file.lastModified}:${randomSuffix()}`
}

function standaloneSavedFileId(name: string): string {
  return `download:${name}:${randomSuffix()}`
}

function downloadBuffer(bytes: ArrayBuffer, name: string, mimeType: string): void {
  const blob = new Blob([bytes], { type: mimeType || 'application/octet-stream' })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = name
  anchor.style.display = 'none'
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}

async function pickBrowserFiles(options: {
  multiple?: boolean | undefined
  accept?: string[] | undefined
}): Promise<File[] | null> {
  const input = document.createElement('input')
  input.type = 'file'
  input.multiple = options.multiple === true
  if (options.accept?.length) input.accept = options.accept.join(',')

  return new Promise<File[] | null>((resolve) => {
    let settled = false
    const finish = (value: File[] | null) => {
      if (settled) return
      settled = true
      input.remove()
      resolve(value)
    }
    input.addEventListener('change', () => finish(input.files ? [...input.files] : null), {
      once: true,
    })
    input.addEventListener('cancel', () => finish(null), { once: true })
    input.style.display = 'none'
    document.body.append(input)
    input.click()
  })
}

async function browserFileToOfficeFile(file: File): Promise<OfficeFile> {
  const bytes = await file.arrayBuffer()
  return {
    id: localFileId(file),
    name: file.name,
    mimeType: file.type || 'application/octet-stream',
    size: file.size,
    version: String(file.lastModified),
    bytes,
    transport: 'buffer',
  }
}

function descriptorOf(file: OfficeFile): OfficeFileDescriptor {
  return {
    id: file.id,
    ...(file.nodeId ? { nodeId: file.nodeId } : {}),
    ...(file.tenantId ? { tenantId: file.tenantId } : {}),
    ...(file.parentId !== undefined ? { parentId: file.parentId } : {}),
    name: file.name,
    mimeType: file.mimeType,
    size: file.size,
    version: file.version,
    ...(file.updatedAt ? { updatedAt: file.updatedAt } : {}),
    transport: 'buffer',
  }
}

function pickerFailure(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

function legacyPickUsesAssets(options: PickFileOptions): boolean {
  if (options.multiple === true) return true
  return Boolean(
    options.accept?.some(
      (value) =>
        value === 'image/*' ||
        value.startsWith('image/') ||
        /^\.(png|jpe?g|gif|bmp|webp|svg)$/i.test(value),
    ),
  )
}

export class StandaloneOfficeHost implements OfficeHostApi {
  private readonly files = new Map<string, OfficeFile>()
  private readonly pendingDocuments = new Map<string, OfficeFile>()
  private readonly notifications = new OfficeNotificationAdapter()
  private readonly deactivateNotifications = activateOfficeNotificationAdapter(this.notifications)
  private dirty = false

  async getLocale(): Promise<string> {
    return document.documentElement.lang || navigator.language || 'en'
  }

  async saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult> {
    let file = input.file
    if (input.mode === 'saveAs' || input.newDocument === true) {
      const requested = window.prompt('Save as', input.file.name)
      if (requested === null) {
        return { ok: false, code: 'CANCELLED', error: 'Save As cancelled.' }
      }
      const name = requested.trim()
      if (!name) {
        return { ok: false, code: 'CANCELLED', error: 'A file name is required.' }
      }
      file = { ...input.file, id: standaloneSavedFileId(name), name, version: null }
    }

    downloadBuffer(input.bytes, file.name, file.mimeType)
    return {
      ok: true,
      file: {
        ...file,
        size: input.bytes.byteLength,
        version:
          input.mode === 'saveAs' || input.newDocument === true
            ? null
            : (input.baseVersion ?? file.version ?? null),
      },
    }
  }

  async saveHistoryVersion(_input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult> {
    return {
      ok: false,
      code: 'SAVE_FAILED',
      error: 'History versions require an embedded platform host.',
    }
  }

  async downloadDocument(input: DownloadDocumentInput): Promise<DownloadDocumentResult> {
    downloadBuffer(input.bytes, input.file.name, input.file.mimeType)
    return { ok: true }
  }

  async pickDocument(options: PickDocumentOptions): Promise<PickDocumentResult> {
    try {
      const files = await pickBrowserFiles({ multiple: false, accept: options.accept })
      if (!files?.[0]) return { status: 'cancelled', selectionId: null, file: null }

      const file = await browserFileToOfficeFile(files[0])
      const selectionId = `selection:${randomSuffix()}`
      this.pendingDocuments.set(selectionId, file)
      return {
        status: 'selected',
        selectionId,
        file: { ...file, bytes: file.bytes.slice(0), transport: 'buffer' },
      }
    } catch (error) {
      return {
        status: 'failed',
        code: 'PICK_DOCUMENT_FAILED',
        error: error instanceof Error ? error.message : String(error),
      }
    }
  }

  async confirmDocumentOpened(selectionId: string): Promise<DocumentOpenedResult> {
    const file = this.pendingDocuments.get(selectionId)
    if (!file) {
      return {
        ok: false,
        code: 'NOT_FOUND',
        error: `Standalone document selection is not available: ${selectionId}`,
      }
    }
    this.pendingDocuments.delete(selectionId)
    this.files.set(file.id, file)
    return { ok: true, file: descriptorOf(file) }
  }

  async releasePickedDocument(selectionId: string): Promise<void> {
    this.pendingDocuments.delete(selectionId)
  }

  async pickAssets(options: PickAssetsOptions): Promise<PickAssetsResult> {
    try {
      const files = await pickBrowserFiles(options)
      if (!files?.length) return { status: 'cancelled', files: [] }

      const selected: OfficeFile[] = []
      for (const file of files) {
        const officeFile = await browserFileToOfficeFile(file)
        selected.push({ ...officeFile, bytes: officeFile.bytes.slice(0), transport: 'buffer' })
      }
      return { status: 'selected', files: selected }
    } catch (error) {
      return {
        status: 'failed',
        files: [],
        code: 'PICK_ASSETS_FAILED',
        error: error instanceof Error ? error.message : String(error),
      }
    }
  }

  async approveClose(_requestId?: string): Promise<void> {
    window.close()
  }

  async cancelClose(_requestId: string): Promise<void> {}

  /** @deprecated Compatibility alias. */
  async exportDocument(input: ExportDocumentInput): Promise<ExportDocumentResult> {
    return this.downloadDocument(input)
  }

  /** @deprecated Compatibility alias. */
  async requestClose(): Promise<void> {
    await this.approveClose()
  }

  /** @deprecated Compatibility generic picker. */
  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {
    if (options.mode !== 'folder' && legacyPickUsesAssets(options)) {
      const result = await this.pickAssets({ multiple: options.multiple, accept: options.accept })
      if (result.status === 'cancelled') return null
      if (result.status === 'failed') throw pickerFailure(result.code, result.error)
      return result.files.map((file) => ({
        ...descriptorOf(file),
        transport: 'buffer',
        bytes: file.bytes.slice(0),
      }))
    }

    const files = await pickBrowserFiles(options)
    if (!files?.length) return null

    const selected: SelectedOfficeFile[] = []
    for (const file of files) {
      const officeFile = await browserFileToOfficeFile(file)
      this.files.set(officeFile.id, officeFile)
      selected.push({
        id: officeFile.id,
        name: officeFile.name,
        mimeType: officeFile.mimeType,
        size: officeFile.size,
        version: officeFile.version,
        transport: 'buffer',
        bytes: officeFile.bytes.slice(0),
      })
    }
    return selected
  }

  /** @deprecated Stable pickers return bytes directly. */
  async readFile(fileId: string): Promise<OfficeFile> {
    const file = this.files.get(fileId)
    if (!file) throw new Error(`Standalone file is not available: ${fileId}`)
    return { ...file, bytes: file.bytes.slice(0), transport: 'buffer' }
  }

  setDirty(dirty: boolean): void {
    if (this.dirty === dirty) return
    this.dirty = dirty
    if (dirty) window.addEventListener('beforeunload', this.beforeUnload)
    else window.removeEventListener('beforeunload', this.beforeUnload)
  }

  setTitle(title: string): void {
    document.title = title
  }

  destroy(): void {
    window.removeEventListener('beforeunload', this.beforeUnload)
    this.deactivateNotifications()
    this.files.clear()
    this.pendingDocuments.clear()
  }

  private readonly beforeUnload = (event: BeforeUnloadEvent): void => {
    event.preventDefault()
  }
}

export interface EmbeddedOfficeRuntimeOptions {
  hostOrigin: string
  locale?: string
  currentWindow?: Window
  requestTimeoutMs?: number
}

export class EmbeddedOfficeHost implements OfficeHostApi {
  private legacyPendingDocumentSelectionId: string | null = null
  private legacyDocumentBind: Promise<DocumentOpenedResult> | null = null

  constructor(
    private readonly bridge: EditorIframeBridge,
    private readonly locale?: string,
  ) {}

  async getLocale(): Promise<string> {
    return this.locale || document.documentElement.lang || navigator.language || 'en'
  }

  private async consumeLegacyDocumentBind(): Promise<DocumentOpenedResult | null> {
    const pending = this.legacyDocumentBind
    if (!pending) return null
    const result = await pending
    if (this.legacyDocumentBind === pending) this.legacyDocumentBind = null
    return result
  }

  async saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult> {
    const bindResult = await this.consumeLegacyDocumentBind()
    if (bindResult && !bindResult.ok) {
      return {
        ok: false,
        code: bindResult.code || 'FILE_BIND_FAILED',
        error: bindResult.error || 'The selected document could not be bound as current.',
      }
    }

    const requestId = createOfficeRequestId('save')
    const bytes = input.bytes.slice(0)
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:save-document-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document',
        requestId,
        payload: {
          file: input.file,
          bytes,
          baseVersion: input.baseVersion,
          mode: input.mode,
          newDocument: input.newDocument,
        },
      },
      'office:save-document-result',
      [bytes],
    )
    return response.payload
  }

  async saveHistoryVersion(input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult> {
    const bindResult = await this.consumeLegacyDocumentBind()
    if (bindResult && !bindResult.ok) {
      return {
        ok: false,
        code: bindResult.code || 'FILE_BIND_FAILED',
        error: bindResult.error || 'The selected document could not be bound as current.',
      }
    }

    const requestId = createOfficeRequestId('history')
    const bytes = input.bytes.slice(0)
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:save-history-version-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version',
        requestId,
        payload: { file: input.file, bytes, baseVersion: input.baseVersion },
      },
      'office:save-history-version-result',
      [bytes],
    )
    return response.payload
  }

  async downloadDocument(input: DownloadDocumentInput): Promise<DownloadDocumentResult> {
    const requestId = createOfficeRequestId('download')
    const bytes = input.bytes.slice(0)
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:download-document-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:download-document',
        requestId,
        payload: { format: input.format, file: input.file, bytes },
      },
      'office:download-document-result',
      [bytes],
    )
    return response.payload
  }

  async pickDocument(options: PickDocumentOptions): Promise<PickDocumentResult> {
    const requestId = createOfficeRequestId('pick-document')
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:pick-document-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-document',
        requestId,
        payload: options,
      },
      'office:pick-document-result',
    )
    return response.payload
  }

  async confirmDocumentOpened(selectionId: string): Promise<DocumentOpenedResult> {
    const requestId = createOfficeRequestId('document-opened')
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:document-opened-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:document-opened',
        requestId,
        payload: { selectionId },
      },
      'office:document-opened-result',
    )
    return response.payload
  }

  async releasePickedDocument(selectionId: string): Promise<void> {
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:document-open-failed',
      requestId: createOfficeRequestId('document-open-failed'),
      payload: { selectionId },
    })
  }

  async pickAssets(options: PickAssetsOptions): Promise<PickAssetsResult> {
    const requestId = createOfficeRequestId('pick-assets')
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:pick-assets-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-assets',
        requestId,
        payload: options,
      },
      'office:pick-assets-result',
    )
    return response.payload
  }

  async approveClose(requestId?: string): Promise<void> {
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:close-approved',
      requestId: requestId ?? createOfficeRequestId('close'),
      payload: { reason: requestId ? 'window-close' : 'file-menu' },
    })
  }

  async cancelClose(requestId: string): Promise<void> {
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:close-cancelled',
      requestId,
      payload: { reason: 'user-cancelled' },
    })
  }

  /** @deprecated Runtime compatibility alias; wire protocol uses download-document. */
  async exportDocument(input: ExportDocumentInput): Promise<ExportDocumentResult> {
    return this.downloadDocument(input)
  }

  /** @deprecated Runtime compatibility alias; wire protocol uses close-approved. */
  async requestClose(): Promise<void> {
    await this.approveClose()
  }

  /**
   * @deprecated Editor compatibility shim. Embedded wire traffic is translated
   * to pick-assets or transactional pick-document; only folder-mode callers
   * still use the old generic wire message.
   */
  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {
    if (options.mode !== 'folder' && legacyPickUsesAssets(options)) {
      const result = await this.pickAssets({ multiple: options.multiple, accept: options.accept })
      if (result.status === 'cancelled') return null
      if (result.status === 'failed') throw pickerFailure(result.code, result.error)
      return result.files.map((file) => ({
        ...descriptorOf(file),
        transport: 'buffer',
        bytes: file.bytes.slice(0),
      }))
    }

    if (options.mode !== 'folder') {
      if (this.legacyPendingDocumentSelectionId) {
        await this.releasePickedDocument(this.legacyPendingDocumentSelectionId)
        this.legacyPendingDocumentSelectionId = null
      }
      const result = await this.pickDocument({ accept: options.accept })
      if (result.status === 'cancelled') return null
      if (result.status === 'failed') throw pickerFailure(result.code, result.error)
      this.legacyPendingDocumentSelectionId = result.selectionId
      return [
        {
          ...descriptorOf(result.file),
          transport: 'buffer',
          bytes: result.file.bytes.slice(0),
        },
      ]
    }

    const requestId = createOfficeRequestId('pick-legacy')
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:pick-file-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-file',
        requestId,
        payload: options,
      },
      'office:pick-file-result',
    )
    return response.payload.files
  }

  /** @deprecated Stable pickers return bytes directly. */
  async readFile(fileId: string): Promise<OfficeFile> {
    const requestId = createOfficeRequestId('read')
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:read-file-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:read-file',
        requestId,
        payload: { fileId },
      },
      'office:read-file-result',
    )
    return response.payload.file
  }

  setDirty(dirty: boolean): void {
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:dirty-change',
      payload: { dirty },
    })
  }

  setTitle(title: string): void {
    if (this.legacyPendingDocumentSelectionId) {
      const selectionId = this.legacyPendingDocumentSelectionId
      this.legacyPendingDocumentSelectionId = null
      this.legacyDocumentBind = this.confirmDocumentOpened(selectionId)
    }
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:title-change',
      payload: { title },
    })
  }

  destroy(): void {
    if (this.legacyPendingDocumentSelectionId) {
      void this.releasePickedDocument(this.legacyPendingDocumentSelectionId)
      this.legacyPendingDocumentSelectionId = null
    }
    this.legacyDocumentBind = null
  }
}

export interface EmbeddedOfficeRuntime {
  bridge: EditorIframeBridge
  host: EmbeddedOfficeHost
  notifications: OfficeNotificationAdapter
  destroy(): void
}

export function createEmbeddedOfficeRuntime(
  options: EmbeddedOfficeRuntimeOptions,
): EmbeddedOfficeRuntime {
  const currentWindow = options.currentWindow ?? window
  if (currentWindow.parent === currentWindow) {
    throw new Error('Embedded Office runtime requires an iframe parent window')
  }
  if (!options.hostOrigin || options.hostOrigin === '*') {
    throw new Error('Embedded Office runtime requires an explicit hostOrigin')
  }

  const bridge = new OfficeIframeBridge<HostToEditorMessage, EditorToHostMessage>({
    sourceWindow: currentWindow,
    targetWindow: currentWindow.parent,
    targetOrigin: options.hostOrigin,
    requestTimeoutMs: options.requestTimeoutMs,
  })
  const notifications = new OfficeNotificationAdapter(bridge)
  const unsubscribeNotifications = bridge.subscribe((message) => {
    if (message.type === 'office:init' || message.type === 'office:new') {
      notifications.configure(message.payload.notifications)
    }
  })
  const deactivateNotifications = activateOfficeNotificationAdapter(notifications)
  bridge.start()
  const host = new EmbeddedOfficeHost(bridge, options.locale)

  return {
    bridge,
    host,
    notifications,
    destroy: () => {
      unsubscribeNotifications()
      deactivateNotifications()
      host.destroy()
      bridge.destroy()
    },
  }
}
