import type {
  OfficeEditorMode,
  OfficeFile,
  OfficeFileDescriptor,
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

interface UcRpcResponse {
  type: 'uc-plugin-rpc-response'
  id: string
  pluginId: string
  result?: unknown
  data?: unknown
  payload?: unknown
  error?: unknown
}

interface UcFileAccess {
  nodeId?: string
  id?: string
  resultNodeId?: string
  filename?: string
  version?: string | number | null
  fileVersion?: string | number | null
  writeMode?: string
  [key: string]: unknown
}

interface PendingRpc {
  resolve(value: unknown): void
  reject(error: Error): void
  timer: number
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

function normalizeXlsxName(value: string): string {
  const cleaned = value
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
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
let currentAccess: UcFileAccess | null = null
const pendingRpc = new Map<string, PendingRpc>()
const localAssets = new Map<string, OfficeFile>()

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

function rpcError(error: unknown): Error {
  if (error instanceof Error) return error
  const value = asRecord(error)
  const message = stringValue(value.message) || stringValue(value.error) || JSON.stringify(error)
  return new Error(message || 'UC RPC failed.')
}

function ucCall(method: string, params?: unknown, timeoutMs = 30_000): Promise<unknown> {
  const id = ucRequestId()
  return new Promise((resolve, reject) => {
    const timer = window.setTimeout(() => {
      pendingRpc.delete(id)
      reject(new Error(`UC RPC timed out: ${method}`))
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

function officeCapabilities() {
  return {
    ai: false,
    open: false,
    save: currentMode === 'edit',
    saveAs: currentMode === 'edit',
    saveHistoryVersion: false,
    exportDocx: false,
    exportPptx: false,
    exportXlsx: true,
    close: true,
    autoSave: 'host' as const,
    download: false,
    print: true,
    systemFilePicker: true,
    pageCropMarks: false,
  }
}

function sendOfficeInit(): void {
  if (!editorReady || !currentFile) return
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
        file: { ...currentFile, bytes },
      },
    },
    [bytes],
  )
}

function unwrapLaunchParams(value: unknown): Record<string, any> {
  const root = asRecord(value)
  return asRecord(root.launchParams || root.params || root.context || root)
}

async function requestSelectedFileAccess(
  writeMode: 'selected' | 'result',
  filename: string,
): Promise<UcFileAccess> {
  const result = await ucCall(
    'uc.fs.requestSelectedFileAccess',
    {
      writeMode,
      filename,
      state: `excel-${writeMode}-${Date.now()}`,
    },
    30_000,
  )
  if (!result || typeof result !== 'object') {
    throw new Error('宿主未返回有效的文件访问授权。')
  }
  return result as UcFileAccess
}

async function readSelectedFile(
  access: UcFileAccess,
  fallbackName: string,
  fallbackId: string,
): Promise<OfficeFile> {
  const result = asRecord(await ucCall('uc.fs.readSelectedFile', undefined, 120_000))
  const blob = result.blob
  if (!(blob instanceof Blob)) throw new Error('宿主未返回可读取的 Excel 文件 blob。')

  const name = normalizeXlsxName(stringValue(result.filename) || stringValue(access.filename) || fallbackName)
  const id =
    stringValue(access.nodeId) || stringValue(access.id) || fallbackId || `uc-xlsx-${Date.now()}`
  const version =
    stringValue(access.version) || stringValue(access.fileVersion) || null

  return {
    id,
    name,
    mimeType: stringValue(result.contentType) || XLSX_MIME,
    size: blob.size,
    version,
    bytes: await blob.arrayBuffer(),
  }
}

async function initializeFromUc(): Promise<void> {
  await ucCall('uc.ready', undefined, 30_000)
  const launch = unwrapLaunchParams(await ucCall('uc.host.getLaunchParams', undefined, 30_000))
  const launchFile = asRecord(launch.file || launch.selectedFile || launch.node || {})
  const filename = normalizeXlsxName(
    stringValue(launch.fileName) ||
      stringValue(launchFile.name) ||
      stringValue(launchFile.filename) ||
      query.get('fileName') ||
      'workbook.xlsx',
  )
  const nodeId =
    stringValue(launch.nodeId) || stringValue(launchFile.nodeId) || stringValue(launchFile.id)
  if (!nodeId) throw new Error('UC 启动参数缺少 nodeId，无法安全打开工作簿。')

  if (launch.mode === 'view' || launchFile.mode === 'view') currentMode = 'view'
  const launchLocale = stringValue(launch.locale) || stringValue(launchFile.locale)
  if (launchLocale) currentLocale = launchLocale

  currentAccess = await requestSelectedFileAccess('selected', filename)
  currentFile = await readSelectedFile(currentAccess, filename, nodeId)
  sendOfficeInit()
}

async function pickSaveDestination(suggestedName: string): Promise<{ filename: string } | null> {
  const response = asRecord(
    await ucCall(
      'uc.fs.pickSaveDestination',
      {
        title: '另存为 Excel 工作簿',
        confirmText: '保存',
        suggestedName,
        fileTypes: [
          {
            id: 'xlsx',
            label: 'Excel 工作簿',
            extension: '.xlsx',
            mimeType: XLSX_MIME,
          },
        ],
        activeFileTypeId: 'xlsx',
        allowFileTypeChange: false,
      },
      300_000,
    ),
  )
  if (response.cancelled) return null
  return { filename: normalizeXlsxName(stringValue(response.filename) || suggestedName) }
}

function saveResponseDescriptor(
  response: unknown,
  access: UcFileAccess,
  filename: string,
  fallback: OfficeFile | null,
): OfficeFileDescriptor {
  const root = asRecord(response)
  const candidate = asRecord(root.file || root.node || root.result || root.savedFile || root)
  const id =
    stringValue(candidate.nodeId) ||
    stringValue(candidate.id) ||
    stringValue(root.nodeId) ||
    stringValue(root.id) ||
    stringValue(access.resultNodeId) ||
    stringValue(access.nodeId) ||
    fallback?.id ||
    ''
  const version =
    stringValue(candidate.version) ||
    stringValue(candidate.fileVersion) ||
    stringValue(root.version) ||
    stringValue(root.fileVersion) ||
    stringValue(access.version) ||
    fallback?.version ||
    null

  if (!id) throw new Error('保存成功，但宿主没有返回 nodeId/id。')
  return {
    id,
    name: filename,
    mimeType: XLSX_MIME,
    version,
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
      payload: { ok: false, code: 'SAVE_FAILED', error: '当前工作簿为只读模式。' },
    })
    return
  }

  try {
    const saveAs = message.payload.mode === 'saveAs' || !currentFile
    let filename = normalizeXlsxName(currentFile?.name || message.payload.file.name || 'workbook.xlsx')
    let writeMode: 'selected' | 'result' = 'selected'

    if (saveAs) {
      const destination = await pickSaveDestination(filename)
      if (!destination) {
        sendOffice({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:save-document-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'CANCELLED', error: '已取消另存为。' },
        })
        return
      }
      filename = destination.filename
      writeMode = 'result'
    }

    const access = await requestSelectedFileAccess(writeMode, filename)
    const bytes = message.payload.bytes.slice(0)
    const response = await ucCall(
      'uc.fs.saveResultFile',
      {
        blob: new Blob([bytes], { type: XLSX_MIME }),
        filename,
      },
      120_000,
    )
    const descriptor = saveResponseDescriptor(response, access, filename, currentFile)

    if (saveAs && currentFile && descriptor.id === currentFile.id) {
      throw new Error(
        'writeMode=result 已完成，但后端没有返回新文件的 nodeId/id；否则后续 Ctrl+S 无法安全写回新文件。',
      )
    }

    currentAccess = access
    currentFile = {
      ...descriptor,
      size: bytes.byteLength,
      bytes,
    }
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-document-result',
      requestId: message.requestId,
      payload: {
        ok: true,
        file: {
          ...descriptor,
          size: bytes.byteLength,
        },
      },
    })
  } catch (error) {
    sendOffice({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-document-result',
      requestId: message.requestId,
      payload: {
        ok: false,
        code: 'SAVE_FAILED',
        error: error instanceof Error ? error.message : String(error),
      },
    })
  }
}

async function openLocalAssetPicker(
  message: Extract<EditorToHostMessage, { type: 'office:pick-file' }>,
): Promise<SelectedOfficeFile[] | null> {
  // The current UC plugin contract does not yet expose a confirmed interactive
  // open-file picker RPC. Keep the fallback isolated here; when UC adds one,
  // replace only this function and keep the office:* protocol unchanged.
  const input = document.createElement('input')
  input.type = 'file'
  input.multiple = message.payload.multiple === true
  if (message.payload.accept?.length) input.accept = message.payload.accept.join(',')
  input.style.display = 'none'
  document.body.append(input)

  const files = await new Promise<File[] | null>((resolve) => {
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
    input.click()
  })
  if (!files?.length) return null

  const selected: SelectedOfficeFile[] = []
  for (const file of files) {
    const id = `local-asset:${crypto.randomUUID()}`
    const bytes = await file.arrayBuffer()
    const officeFile: OfficeFile = {
      id,
      name: file.name,
      mimeType: file.type || 'application/octet-stream',
      size: bytes.byteLength,
      version: String(file.lastModified),
      bytes,
    }
    localAssets.set(id, officeFile)
    selected.push({
      id,
      name: officeFile.name,
      mimeType: officeFile.mimeType,
      size: officeFile.size,
      version: officeFile.version,
      transport: 'token',
      token: id,
    })
  }
  return selected
}

function fileForRead(fileId: string): OfficeFile | null {
  if (currentFile?.id === fileId) return currentFile
  return localAssets.get(fileId) || null
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'xlsx') return
      editorReady = true
      sendOfficeInit()
      return
    case 'office:pick-file': {
      const files = await openLocalAssetPicker(message)
      sendOffice({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-file-result',
        requestId: message.requestId,
        payload: { files },
      })
      return
    }
    case 'office:read-file': {
      const file = fileForRead(message.payload.fileId)
      if (!file) {
        sendOffice({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:error',
          requestId: message.requestId,
          payload: { code: 'READ_FAILED', message: '请求的文件不在当前 UC Office 会话中。' },
        })
        return
      }
      const bytes = file.bytes.slice(0)
      sendOffice(
        {
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:read-file-result',
          requestId: message.requestId,
          payload: { file: { ...file, bytes } },
        },
        [bytes],
      )
      return
    }
    case 'office:save-document':
      await saveOfficeDocument(message)
      return
    case 'office:save-history-version':
      sendOffice({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version-result',
        requestId: message.requestId,
        payload: { ok: false, code: 'SAVE_FAILED', error: 'UC Excel Host 暂未启用历史版本保存。' },
      })
      return
    case 'office:export-document':
      sendOffice({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:export-document-result',
        requestId: message.requestId,
        payload: { ok: false, code: 'SAVE_FAILED', error: 'UC Excel Host 暂未启用独立导出。' },
      })
      return
    case 'office:dirty-change':
    case 'office:title-change':
    case 'office:state-result':
    case 'office:save-result':
    case 'office:close-request':
    case 'office:close-cancelled':
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
