import type {
  OfficeEditorMode,
  OfficeFile,
  OfficeFileDescriptor,
  OfficeFileVersion,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  isOfficeProtocolMessage,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const DEFAULT_SHEETS_URL = 'http://127.0.0.1:5275'
const DEFAULT_PLUGIN_ID = 'thirdparty.plugin.excel-online'
const INVALID_FILENAME_CHARS = new Set(['\\', '/', ':', '*', '?', '"', '<', '>', '|'])

interface UcRpcResponse {
  type: 'uc-plugin-rpc-response'
  id: string
  pluginId: string
  result?: unknown
  data?: unknown
  payload?: unknown
  error?: unknown
}

interface PendingRpc {
  resolve(value: unknown): void
  reject(error: Error): void
  timer: number
}

class UcRpcError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'UcRpcError'
  }
}

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
}

function asRecord(value: unknown): Record<string, any> {
  return value && typeof value === 'object' ? (value as Record<string, any>) : {}
}

function stringValue(value: unknown): string | null {
  if (value === undefined || value === null) return null
  const text = String(value).trim()
  return text || null
}

function versionValue(value: unknown): OfficeFileVersion | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  return stringValue(value) ?? undefined
}

function normalizeXlsxName(value: string): string {
  const cleaned = [...value]
    .map((char) => (char.charCodeAt(0) <= 0x1f || INVALID_FILENAME_CHARS.has(char) ? ' ' : char))
    .join('')
    .replace(/\s+/g, ' ')
    .trim()
  if (!cleaned) throw new Error('文件名不能为空。')
  return /\.xlsx$/i.test(cleaned) ? cleaned : `${cleaned}.xlsx`
}

function referrerOrigin(): string | null {
  if (!document.referrer) return null
  try {
    return new URL(document.referrer).origin
  } catch {
    return null
  }
}

const frame = requireElement<HTMLIFrameElement>('office-frame')
const errorPanel = requireElement<HTMLDivElement>('host-error')
const query = new URLSearchParams(window.location.search)
const sheetsUrl = query.get('sheetsUrl') || DEFAULT_SHEETS_URL
const sheetsOrigin = new URL(sheetsUrl).origin
const pluginId = query.get('pluginId') || DEFAULT_PLUGIN_ID
const ucHostOrigin = (() => {
  const origin = query.get('ucHostOrigin') || referrerOrigin()
  if (!origin) {
    throw new Error('UC Host origin is required. Use ?ucHostOrigin=https://webos.example.com.')
  }
  return new URL(origin).origin
})()

let ucSeq = 0
let officeSeq = 0
let editorReady = false
let currentMode: OfficeEditorMode = 'edit'
let currentLocale = query.get('locale') || 'zh-CN'
let currentFile: OfficeFile | null = null
const pendingRpc = new Map<string, PendingRpc>()
const pendingDocuments = new Map<string, OfficeFile>()
const legacyFiles = new Map<string, OfficeFile>()

function showError(error: unknown): void {
  const message = error instanceof Error ? error.message : String(error)
  errorPanel.textContent = message
  errorPanel.style.display = 'block'
  console.error('[UC GenOffice Excel Host]', error)
}

function officeRequestId(prefix: string): string {
  officeSeq += 1
  return `${prefix}-${Date.now()}-${officeSeq}`
}

function ucRequestId(): string {
  ucSeq += 1
  return `uc-${Date.now()}-${ucSeq}`
}

function ucResult(message: UcRpcResponse): unknown {
  if (message.result !== undefined) return message.result
  if (message.data !== undefined) return message.data
  if (message.payload !== undefined) return message.payload
  return message
}

function rpcError(error: unknown): UcRpcError {
  const value = asRecord(error)
  const code = stringValue(value.code) || stringValue(value.errorCode) || 'UC_RPC_FAILED'
  const message =
    stringValue(value.message) || stringValue(value.error) || JSON.stringify(error) || 'UC RPC failed.'
  return new UcRpcError(code, message)
}

function ucCall(method: string, params?: unknown, timeoutMs = 30_000): Promise<unknown> {
  const id = ucRequestId()
  return new Promise((resolve, reject) => {
    const timer = window.setTimeout(() => {
      pendingRpc.delete(id)
      reject(new UcRpcError('UC_RPC_TIMEOUT', `UC RPC timed out: ${method}`))
    }, timeoutMs)
    pendingRpc.set(id, { resolve, reject, timer })
    window.parent.postMessage(
      {
        type: 'uc-plugin-rpc-request',
        id,
        pluginId,
        method,
        params,
      },
      ucHostOrigin,
    )
  })
}

function sendOffice(message: HostToEditorMessage, transfer: Transferable[] = []): void {
  const target = frame.contentWindow
  if (!target) throw new Error('GenOffice Sheets iframe is not available.')
  target.postMessage(message, sheetsOrigin, transfer)
}

function forwardOfficeControl(message: EditorToHostMessage): void {
  window.parent.postMessage(message, ucHostOrigin)
}

function officeCapabilities() {
  return {
    ai: false,
    open: true,
    openDocument: true,
    pickAssets: true,
    save: currentMode === 'edit',
    saveAs: currentMode === 'edit',
    saveHistoryVersion: currentMode === 'edit',
    exportDocx: false,
    exportPptx: false,
    exportXlsx: true,
    close: true,
    autoSave: 'host' as const,
    download: true,
    print: true,
    systemFilePicker: true,
    pageCropMarks: false,
  }
}

function descriptorOf(file: OfficeFile): OfficeFileDescriptor {
  return {
    id: file.id,
    ...(file.nodeId ? { nodeId: file.nodeId } : {}),
    ...(file.tenantId ? { tenantId: file.tenantId } : {}),
    name: file.name,
    mimeType: file.mimeType,
    ...(file.size === undefined ? {} : { size: file.size }),
    ...(file.version === undefined ? {} : { version: file.version }),
  }
}

function unwrapLaunchParams(value: unknown): Record<string, any> {
  const root = asRecord(value)
  return asRecord(root.launchParams || root.params || root.context || root)
}

function cancelled(value: unknown): boolean {
  const root = asRecord(value)
  return root.cancelled === true || root.status === 'cancelled'
}

function fileRecord(value: unknown): Record<string, any> {
  const root = asRecord(value)
  return asRecord(root.file || root.node || root.savedFile || root.result || root)
}

function fileIdentity(value: unknown): string | null {
  const root = asRecord(value)
  const candidate = fileRecord(value)
  return (
    stringValue(candidate.nodeId) ||
    stringValue(candidate.id) ||
    stringValue(root.nodeId) ||
    stringValue(root.id)
  )
}

function fileDescriptorFromUc(
  value: unknown,
  fallback?: OfficeFileDescriptor | null,
  fallbackName = 'workbook.xlsx',
): OfficeFileDescriptor {
  const root = asRecord(value)
  const candidate = fileRecord(value)
  const id = fileIdentity(value) || fallback?.id || ''
  if (!id) throw new Error('UC 返回文件信息缺少 nodeId/id。')
  const name = normalizeXlsxName(
    stringValue(candidate.name) ||
      stringValue(candidate.filename) ||
      stringValue(root.name) ||
      stringValue(root.filename) ||
      fallback?.name ||
      fallbackName,
  )
  const version =
    versionValue(candidate.version) ??
    versionValue(candidate.fileVersion) ??
    versionValue(root.version) ??
    versionValue(root.fileVersion) ??
    fallback?.version
  return {
    id,
    nodeId: stringValue(candidate.nodeId) || stringValue(root.nodeId) || fallback?.nodeId || id,
    ...(stringValue(candidate.tenantId) || stringValue(root.tenantId) || fallback?.tenantId
      ? {
          tenantId:
            stringValue(candidate.tenantId) || stringValue(root.tenantId) || fallback?.tenantId,
        }
      : {}),
    name,
    mimeType:
      stringValue(candidate.mimeType) ||
      stringValue(candidate.contentType) ||
      stringValue(root.mimeType) ||
      stringValue(root.contentType) ||
      fallback?.mimeType ||
      XLSX_MIME,
    ...(typeof candidate.size === 'number'
      ? { size: candidate.size }
      : typeof root.size === 'number'
        ? { size: root.size }
        : fallback?.size === undefined
          ? {}
          : { size: fallback.size }),
    ...(version === undefined ? {} : { version }),
  }
}

async function officeFileFromUc(value: unknown, fallbackName = 'workbook.xlsx'): Promise<OfficeFile> {
  const root = asRecord(value)
  const candidate = fileRecord(value)
  const blob = candidate.blob instanceof Blob ? candidate.blob : root.blob instanceof Blob ? root.blob : null
  if (!blob) throw new Error('UC 未返回可读取的文件 Blob。')
  const descriptor = fileDescriptorFromUc(value, null, fallbackName)
  return {
    ...descriptor,
    size: descriptor.size ?? blob.size,
    bytes: await blob.arrayBuffer(),
    transport: 'buffer',
  }
}

function sendOfficeInitOrNew(): void {
  if (!editorReady) return
  if (!currentFile) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:new',
      requestId: officeRequestId('new'),
      payload: {
        kind: 'xlsx',
        mode: currentMode,
        locale: currentLocale,
        capabilities: officeCapabilities(),
      },
    })
    return
  }

  const bytes = currentFile.bytes.slice(0)
  sendOffice(
    {
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: officeRequestId('init'),
      payload: {
        kind: 'xlsx',
        mode: currentMode,
        locale: currentLocale,
        capabilities: officeCapabilities(),
        file: { ...currentFile, bytes, transport: 'buffer' },
      },
    },
    [bytes],
  )
}

async function initializeFromUc(): Promise<void> {
  await ucCall('uc.ready', undefined, 30_000)

  try {
    const launch = unwrapLaunchParams(await ucCall('uc.host.getLaunchParams', undefined, 30_000))
    if (launch.mode === 'view') currentMode = 'view'
    const launchLocale = stringValue(launch.locale)
    if (launchLocale) currentLocale = launchLocale
  } catch (error) {
    console.warn('[UC GenOffice Excel Host] launch params unavailable', error)
  }

  const result = await ucCall('uc.fs.readCurrentFile', undefined, 120_000)
  if (result === null || result === undefined || cancelled(result)) {
    currentFile = null
  } else {
    currentFile = await officeFileFromUc(result)
    legacyFiles.set(currentFile.id, currentFile)
  }
  sendOfficeInitOrNew()
}

function errorCode(error: unknown, fallback = 'FILE_OPERATION_FAILED'): string {
  return error instanceof UcRpcError ? error.code : fallback
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function saveFileTypes() {
  return [
    {
      id: 'xlsx',
      label: 'Excel 工作簿',
      extension: '.xlsx',
      mimeType: XLSX_MIME,
    },
  ]
}

async function handlePickDocument(
  message: Extract<EditorToHostMessage, { type: 'office:pick-document' }>,
): Promise<void> {
  try {
    const result = await ucCall(
      'uc.fs.pickFile',
      { accept: message.payload.accept || [XLSX_MIME, '.xlsx'], multiple: false },
      300_000,
    )
    if (cancelled(result)) {
      sendOffice({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-document-result',
        requestId: message.requestId,
        payload: { status: 'cancelled', selectionId: null, file: null },
      })
      return
    }

    const root = asRecord(result)
    const selectionId = stringValue(root.selectionId)
    if (!selectionId) throw new Error('uc.fs.pickFile 未返回 selectionId。')
    const file = await officeFileFromUc(result)
    pendingDocuments.set(selectionId, file)
    const bytes = file.bytes.slice(0)
    sendOffice(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-document-result',
        requestId: message.requestId,
        payload: {
          status: 'selected',
          selectionId,
          file: { ...file, bytes, transport: 'buffer' },
        },
      },
      [bytes],
    )
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:error',
      requestId: message.requestId,
      payload: { code: errorCode(error, 'FILE_PICK_FAILED'), message: errorMessage(error) },
    })
  }
}

async function handleDocumentOpened(
  message: Extract<EditorToHostMessage, { type: 'office:document-opened' }>,
): Promise<void> {
  const candidate = pendingDocuments.get(message.payload.selectionId)
  if (!candidate) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:document-opened-result',
      requestId: message.requestId,
      payload: { ok: false, code: 'NOT_FOUND', error: '待提交文件选择已经失效。' },
    })
    return
  }

  try {
    const result = await ucCall(
      'uc.fs.bindCurrentFile',
      { selectionId: message.payload.selectionId },
      30_000,
    )
    const descriptor = fileDescriptorFromUc(result, descriptorOf(candidate), candidate.name)
    currentFile = {
      ...candidate,
      ...descriptor,
      bytes: candidate.bytes,
      transport: 'buffer',
    }
    pendingDocuments.delete(message.payload.selectionId)
    legacyFiles.set(currentFile.id, currentFile)
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:document-opened-result',
      requestId: message.requestId,
      payload: { ok: true, file: descriptor },
    })
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:document-opened-result',
      requestId: message.requestId,
      payload: {
        ok: false,
        code: errorCode(error, 'FILE_BIND_FAILED'),
        error: errorMessage(error),
      },
    })
  }
}

async function handleDocumentOpenFailed(
  message: Extract<EditorToHostMessage, { type: 'office:document-open-failed' }>,
): Promise<void> {
  pendingDocuments.delete(message.payload.selectionId)
  await ucCall(
    'uc.fs.releasePickedFile',
    { selectionId: message.payload.selectionId },
    30_000,
  ).catch((error) => {
    console.warn('[UC GenOffice Excel Host] releasePickedFile failed', error)
  })
}

async function pickAssetsFromUc(accept?: string[], multiple = false): Promise<OfficeFile[]> {
  const result = await ucCall(
    'uc.fs.pickAssets',
    { accept: accept?.length ? accept : ['image/*'], multiple },
    300_000,
  )
  if (cancelled(result)) return []
  const root = asRecord(result)
  const values = Array.isArray(root.files)
    ? root.files
    : Array.isArray(root.items)
      ? root.items
      : Array.isArray(result)
        ? result
        : []
  const files: OfficeFile[] = []
  for (const value of values) {
    const file = await officeFileFromUc(value, 'asset')
    files.push(file)
    legacyFiles.set(file.id, file)
  }
  return files
}

async function handlePickAssets(
  message: Extract<EditorToHostMessage, { type: 'office:pick-assets' }>,
): Promise<void> {
  try {
    const files = await pickAssetsFromUc(message.payload.accept, message.payload.multiple === true)
    if (!files.length) {
      sendOffice({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-assets-result',
        requestId: message.requestId,
        payload: { status: 'cancelled', files: [] },
      })
      return
    }
    const responseFiles = files.map((file) => ({
      ...file,
      bytes: file.bytes.slice(0),
      transport: 'buffer' as const,
    }))
    sendOffice(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-assets-result',
        requestId: message.requestId,
        payload: { status: 'selected', files: responseFiles },
      },
      responseFiles.map((file) => file.bytes),
    )
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:error',
      requestId: message.requestId,
      payload: { code: errorCode(error, 'ASSET_PICK_FAILED'), message: errorMessage(error) },
    })
  }
}

async function saveOfficeDocument(
  message: Extract<EditorToHostMessage, { type: 'office:save-document' }>,
): Promise<void> {
  if (currentMode !== 'edit') {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-document-result',
      requestId: message.requestId,
      payload: { ok: false, code: 'PERMISSION_REQUIRED', error: '当前工作簿为只读模式。' },
    })
    return
  }

  try {
    const bytes = message.payload.bytes.slice(0)
    const filename = normalizeXlsxName(message.payload.file.name || currentFile?.name || 'workbook.xlsx')
    const saveAs = message.payload.mode === 'saveAs' || message.payload.newDocument === true || !currentFile
    const blob = new Blob([bytes], { type: XLSX_MIME })

    const result = saveAs
      ? await ucCall(
          'uc.fs.saveFileAs',
          {
            blob,
            suggestedName: filename,
            fileTypes: saveFileTypes(),
            title: '另存为 Excel 工作簿',
            confirmText: '保存',
          },
          300_000,
        )
      : await ucCall(
          'uc.fs.saveCurrentFile',
          {
            blob,
            filename,
            baseVersion: message.payload.baseVersion,
          },
          120_000,
        )

    if (cancelled(result)) {
      sendOffice({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: { ok: false, code: 'CANCELLED', error: '已取消保存。' },
      })
      return
    }

    const descriptor = fileDescriptorFromUc(result, saveAs ? null : currentFile, filename)
    if (saveAs && currentFile && descriptor.id === currentFile.id) {
      throw new Error('uc.fs.saveFileAs 必须创建并返回新的 nodeId。')
    }
    currentFile = {
      ...descriptor,
      size: bytes.byteLength,
      bytes,
      transport: 'buffer',
    }
    legacyFiles.set(currentFile.id, currentFile)
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-document-result',
      requestId: message.requestId,
      payload: { ok: true, file: { ...descriptor, size: bytes.byteLength } },
    })
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-document-result',
      requestId: message.requestId,
      payload: {
        ok: false,
        code: errorCode(error, 'FILE_SAVE_FAILED'),
        error: errorMessage(error),
      },
    })
  }
}

async function saveHistoryVersion(
  message: Extract<EditorToHostMessage, { type: 'office:save-history-version' }>,
): Promise<void> {
  if (!currentFile) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-history-version-result',
      requestId: message.requestId,
      payload: { ok: false, code: 'NOT_FOUND', error: '当前工作簿尚未保存。' },
    })
    return
  }

  try {
    const bytes = message.payload.bytes.slice(0)
    const result = await ucCall(
      'uc.fs.createFileVersion',
      {
        blob: new Blob([bytes], { type: XLSX_MIME }),
        filename: currentFile.name,
        baseVersion: message.payload.baseVersion,
      },
      120_000,
    )
    const descriptor = fileDescriptorFromUc(result, currentFile, currentFile.name)
    currentFile = { ...currentFile, ...descriptor, bytes, transport: 'buffer' }
    legacyFiles.set(currentFile.id, currentFile)
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-history-version-result',
      requestId: message.requestId,
      payload: { ok: true, file: descriptor },
    })
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-history-version-result',
      requestId: message.requestId,
      payload: {
        ok: false,
        code: errorCode(error, 'FILE_SAVE_FAILED'),
        error: errorMessage(error),
      },
    })
  }
}

async function downloadOfficeDocument(
  message: Extract<EditorToHostMessage, { type: 'office:download-document' }>,
): Promise<void> {
  try {
    const bytes = message.payload.bytes.slice(0)
    await ucCall(
      'uc.download.saveFile',
      {
        blob: new Blob([bytes], { type: message.payload.file.mimeType || XLSX_MIME }),
        filename: message.payload.file.name,
      },
      120_000,
    )
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:download-document-result',
      requestId: message.requestId,
      payload: { ok: true },
    })
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:download-document-result',
      requestId: message.requestId,
      payload: {
        ok: false,
        code: errorCode(error, 'DOWNLOAD_FAILED'),
        error: errorMessage(error),
      },
    })
  }
}

function isAssetLegacyPick(message: Extract<EditorToHostMessage, { type: 'office:pick-file' }>): boolean {
  return (
    message.payload.multiple === true ||
    Boolean(message.payload.accept?.some((value) => value.startsWith('image/') || value === 'image/*'))
  )
}

async function handleLegacyPickFile(
  message: Extract<EditorToHostMessage, { type: 'office:pick-file' }>,
): Promise<void> {
  try {
    let files: OfficeFile[] = []
    if (isAssetLegacyPick(message)) {
      files = await pickAssetsFromUc(message.payload.accept, message.payload.multiple === true)
    } else {
      const result = await ucCall(
        'uc.fs.pickFile',
        { accept: message.payload.accept || [XLSX_MIME, '.xlsx'], multiple: false },
        300_000,
      )
      if (!cancelled(result)) {
        const root = asRecord(result)
        const selectionId = stringValue(root.selectionId)
        const file = await officeFileFromUc(result)
        if (selectionId) {
          // Compatibility only: the old protocol has no editor-load confirmation,
          // so bind immediately. New Office flows never use this path.
          await ucCall('uc.fs.bindCurrentFile', { selectionId }, 30_000)
        }
        currentFile = file
        legacyFiles.set(file.id, file)
        files = [file]
      }
    }

    const selected: SelectedOfficeFile[] = files.map((file) => ({
      ...descriptorOf(file),
      transport: 'buffer',
      bytes: file.bytes.slice(0),
    }))
    sendOffice(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-file-result',
        requestId: message.requestId,
        payload: { files: selected.length ? selected : null },
      },
      selected.flatMap((file) => (file.bytes ? [file.bytes] : [])),
    )
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:error',
      requestId: message.requestId,
      payload: { code: errorCode(error, 'FILE_PICK_FAILED'), message: errorMessage(error) },
    })
  }
}

async function handleLegacyExport(
  message: Extract<EditorToHostMessage, { type: 'office:export-document' }>,
): Promise<void> {
  try {
    const bytes = message.payload.bytes.slice(0)
    await ucCall(
      'uc.download.saveFile',
      {
        blob: new Blob([bytes], { type: message.payload.file.mimeType || XLSX_MIME }),
        filename: message.payload.file.name,
      },
      120_000,
    )
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:export-document-result',
      requestId: message.requestId,
      payload: { ok: true },
    })
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:export-document-result',
      requestId: message.requestId,
      payload: {
        ok: false,
        code: errorCode(error, 'DOWNLOAD_FAILED'),
        error: errorMessage(error),
      },
    })
  }
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'xlsx') return
      editorReady = true
      sendOfficeInitOrNew()
      return
    case 'office:pick-document':
      await handlePickDocument(message)
      return
    case 'office:document-opened':
      await handleDocumentOpened(message)
      return
    case 'office:document-open-failed':
      await handleDocumentOpenFailed(message)
      return
    case 'office:pick-assets':
      await handlePickAssets(message)
      return
    case 'office:save-document':
      await saveOfficeDocument(message)
      return
    case 'office:save-history-version':
      await saveHistoryVersion(message)
      return
    case 'office:download-document':
      await downloadOfficeDocument(message)
      return
    case 'office:close-approved':
    case 'office:close-cancelled':
      // Window ownership belongs to UC Host. Forward the stable control message
      // instead of trying to manipulate the parent DOM from the Bridge.
      forwardOfficeControl(message)
      return

    // ---- v1 compatibility aliases ----
    case 'office:pick-file':
      await handleLegacyPickFile(message)
      return
    case 'office:read-file': {
      const file = legacyFiles.get(message.payload.fileId)
      if (!file) {
        sendOffice({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:error',
          requestId: message.requestId,
          payload: { code: 'NOT_FOUND', message: '请求的文件不在当前 Office 会话中。' },
        })
        return
      }
      const bytes = file.bytes.slice(0)
      sendOffice(
        {
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:read-file-result',
          requestId: message.requestId,
          payload: { file: { ...file, bytes, transport: 'buffer' } },
        },
        [bytes],
      )
      return
    }
    case 'office:export-document':
      await handleLegacyExport(message)
      return
    case 'office:close-request':
      forwardOfficeControl({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:close-approved',
        requestId: message.requestId || officeRequestId('legacy-close'),
        payload: { reason: message.payload.reason },
      })
      return

    case 'office:dirty-change':
    case 'office:title-change':
    case 'office:state-result':
    case 'office:save-result':
      return
    case 'office:error':
      showError(new Error(`${message.payload.code}: ${message.payload.message}`))
      return
  }
}

window.addEventListener('message', (event) => {
  if (event.source === window.parent && event.origin === ucHostOrigin) {
    const message = event.data as Partial<UcRpcResponse>
    if (
      message?.type === 'uc-plugin-rpc-response' &&
      message.pluginId === pluginId &&
      typeof message.id === 'string'
    ) {
      const pending = pendingRpc.get(message.id)
      if (!pending) return
      pendingRpc.delete(message.id)
      window.clearTimeout(pending.timer)
      if (message.error !== undefined && message.error !== null) pending.reject(rpcError(message.error))
      else pending.resolve(ucResult(message as UcRpcResponse))
    }
    return
  }

  if (event.source === frame.contentWindow && event.origin === sheetsOrigin) {
    if (!isOfficeProtocolMessage(event.data)) return
    void handleEditorMessage(event.data as EditorToHostMessage).catch(showError)
  }
})

frame.src = `${sheetsUrl.replace(/\/$/, '')}/?hostOrigin=${encodeURIComponent(window.location.origin)}`
void initializeFromUc().catch(showError)
