import type { OfficeEditorMode, OfficeFile, OfficeFileDescriptor } from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  isOfficeProtocolMessage,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
}

function defaultSheetsUrl(): string {
  const configured = import.meta.env.VITE_SHEETS_URL
  if (typeof configured === 'string' && configured.trim()) return configured.trim()
  const hostname = window.location.hostname || 'localhost'
  return `${window.location.protocol}//${hostname}:5275`
}

const frame = requireElement<HTMLIFrameElement>('office-frame')
const picker = requireElement<HTMLInputElement>('xlsx-picker')
const newButton = requireElement<HTMLButtonElement>('new-button')
const saveButton = requireElement<HTMLButtonElement>('save-button')
const downloadButton = requireElement<HTMLButtonElement>('download-button')
const modeButton = requireElement<HTMLButtonElement>('mode-button')
const localeButton = requireElement<HTMLButtonElement>('locale-button')
const fileName = requireElement<HTMLSpanElement>('file-name')
const dirtyState = requireElement<HTMLSpanElement>('dirty-state')
const hostState = requireElement<HTMLSpanElement>('host-state')

const query = new URLSearchParams(window.location.search)
const sheetsUrl = query.get('sheetsUrl') || defaultSheetsUrl()
const sheetsOrigin = new URL(sheetsUrl).origin

let editorReady = false
let currentFile: OfficeFile | null = null
let dirty = false
let mode: OfficeEditorMode = 'edit'
let locale = 'zh-CN'
let requestCounter = 0
let versionCounter = 0
let historyVersionCounter = 0
let hostStatus = 'loading'
const files = new Map<string, OfficeFile>()
const historyVersions: OfficeFile[] = []
const pendingDocuments = new Map<string, OfficeFile>()

function requestId(prefix: string): string {
  requestCounter += 1
  return `${prefix}-${requestCounter}`
}

function send(message: HostToEditorMessage, transfer: Transferable[] = []): void {
  frame.contentWindow?.postMessage(message, sheetsOrigin, transfer)
}

function setHostState(value: string): void {
  hostStatus = value
  hostState.textContent = value
}

function render(): void {
  fileName.textContent = currentFile?.name ?? '新建工作簿'
  dirtyState.textContent = dirty ? 'dirty' : 'clean'
  hostState.textContent = hostStatus
  saveButton.disabled = !editorReady || mode === 'view'
  downloadButton.disabled = !currentFile
  modeButton.disabled = !editorReady
  localeButton.disabled = !editorReady
  modeButton.textContent = mode === 'edit' ? '切换为预览' : '切换为编辑'
  localeButton.textContent = locale === 'zh-CN' ? '切换 English' : '切换中文'
}

function officeCapabilities() {
  return {
    ai: false,
    open: true,
    openDocument: true,
    pickAssets: true,
    save: true,
    saveAs: true,
    saveHistoryVersion: true,
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
    name: file.name,
    mimeType: file.mimeType,
    size: file.size,
    version: file.version,
  }
}

function sendNew(): void {
  if (!editorReady) return
  currentFile = null
  dirty = false
  versionCounter = 0
  historyVersionCounter = 0
  setHostState('opening')
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:new',
    requestId: requestId('new'),
    payload: {
      kind: 'xlsx',
      mode,
      locale,
      capabilities: officeCapabilities(),
    },
  })
  render()
}

function sendInit(): void {
  if (!editorReady || !currentFile) return
  setHostState('opening')
  const bytes = currentFile.bytes.slice(0)
  send(
    {
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: requestId('init'),
      payload: {
        kind: 'xlsx',
        mode,
        locale,
        capabilities: officeCapabilities(),
        file: { ...currentFile, bytes, transport: 'buffer' },
      },
    },
    [bytes],
  )
}

function downloadBytes(bytes: ArrayBuffer, name: string, mimeType = XLSX_MIME): void {
  const blob = new Blob([bytes], { type: mimeType })
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

function downloadCurrentFile(): void {
  if (!currentFile) return
  downloadBytes(currentFile.bytes, currentFile.name, currentFile.mimeType || XLSX_MIME)
}

async function toOfficeFile(file: File): Promise<OfficeFile> {
  const bytes = await file.arrayBuffer()
  return {
    id: `xlsx:${crypto.randomUUID()}`,
    name: file.name,
    mimeType: file.type || XLSX_MIME,
    size: bytes.byteLength,
    version: 'v1',
    bytes,
    transport: 'buffer',
  }
}

function normalizeXlsxName(value: string): string {
  const name = value.trim()
  return /\.xlsx$/i.test(name) ? name : `${name}.xlsx`
}

async function pickBrowserFiles(options: {
  accept?: string[] | undefined
  multiple?: boolean | undefined
}): Promise<File[] | null> {
  const input = document.createElement('input')
  input.type = 'file'
  input.multiple = options.multiple === true
  if (options.accept?.length) input.accept = options.accept.join(',')
  input.style.display = 'none'
  document.body.append(input)

  return new Promise((resolve) => {
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
}

async function pickDocument(
  message: Extract<EditorToHostMessage, { type: 'office:pick-document' }>,
): Promise<void> {
  const selected = await pickBrowserFiles({ accept: message.payload.accept, multiple: false })
  if (!selected?.[0]) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-document-result',
      requestId: message.requestId,
      payload: { status: 'cancelled', selectionId: null, file: null },
    })
    return
  }

  const file = await toOfficeFile(selected[0])
  const selectionId = `selection:${crypto.randomUUID()}`
  pendingDocuments.set(selectionId, file)
  const bytes = file.bytes.slice(0)
  send(
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
}

function bindDocument(
  message: Extract<EditorToHostMessage, { type: 'office:document-opened' }>,
): void {
  const file = pendingDocuments.get(message.payload.selectionId)
  if (!file) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:document-opened-result',
      requestId: message.requestId,
      payload: { ok: false, code: 'NOT_FOUND', error: 'Document selection expired.' },
    })
    return
  }

  pendingDocuments.delete(message.payload.selectionId)
  currentFile = file
  files.set(file.id, file)
  dirty = false
  versionCounter = 1
  historyVersionCounter = 0
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:document-opened-result',
    requestId: message.requestId,
    payload: { ok: true, file: descriptorOf(file) },
  })
  setHostState('opened')
  render()
}

async function pickAssets(
  message: Extract<EditorToHostMessage, { type: 'office:pick-assets' }>,
): Promise<void> {
  const selected = await pickBrowserFiles(message.payload)
  if (!selected?.length) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-assets-result',
      requestId: message.requestId,
      payload: { status: 'cancelled', files: [] },
    })
    return
  }

  const officeFiles: OfficeFile[] = []
  for (const browserFile of selected) officeFiles.push(await toOfficeFile(browserFile))
  const responseFiles = officeFiles.map((file) => ({
    ...file,
    bytes: file.bytes.slice(0),
    transport: 'buffer' as const,
  }))
  send(
    {
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-assets-result',
      requestId: message.requestId,
      payload: { status: 'selected', files: responseFiles },
    },
    responseFiles.map((file) => file.bytes),
  )
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'xlsx') return
      editorReady = true
      setHostState('ready')
      render()
      sendNew()
      return
    case 'office:dirty-change':
      dirty = message.payload.dirty
      render()
      return
    case 'office:title-change':
      if (currentFile) currentFile = { ...currentFile, name: message.payload.title }
      render()
      return
    case 'office:pick-document':
      await pickDocument(message)
      return
    case 'office:document-opened':
      bindDocument(message)
      return
    case 'office:document-open-failed':
      pendingDocuments.delete(message.payload.selectionId)
      return
    case 'office:pick-assets':
      await pickAssets(message)
      return
    case 'office:save-document': {
      const saveAs = message.payload.mode === 'saveAs' || message.payload.newDocument === true || !currentFile
      let name = currentFile?.name ?? message.payload.file.name ?? 'Untitled.xlsx'
      if (saveAs) {
        const requested = window.prompt('另存为文件名', name)
        if (requested === null || !requested.trim()) {
          send({
            protocol: OFFICE_PROTOCOL_VERSION,
            type: 'office:save-document-result',
            requestId: message.requestId,
            payload: { ok: false, code: 'CANCELLED', error: '已取消另存为。' },
          })
          setHostState('ready')
          render()
          return
        }
        name = normalizeXlsxName(requested)
      }

      versionCounter += 1
      const bytes = message.payload.bytes.slice(0)
      const id = saveAs ? `xlsx:${crypto.randomUUID()}` : currentFile?.id || `xlsx:${crypto.randomUUID()}`
      currentFile = {
        id,
        name,
        mimeType: XLSX_MIME,
        size: bytes.byteLength,
        version: `v${versionCounter}`,
        bytes,
        transport: 'buffer',
      }
      files.set(id, currentFile)
      dirty = false
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: { ok: true, file: descriptorOf(currentFile) },
      })
      setHostState('saved')
      render()
      return
    }
    case 'office:save-history-version': {
      if (!currentFile) {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:save-history-version-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'NOT_FOUND', error: 'Save the workbook before creating history.' },
        })
        return
      }
      historyVersionCounter += 1
      versionCounter += 1
      const bytes = message.payload.bytes.slice(0)
      historyVersions.push({
        ...currentFile,
        id: `history:${currentFile.id}:${historyVersionCounter}`,
        size: bytes.byteLength,
        version: `history-${historyVersionCounter}`,
        bytes,
        transport: 'buffer',
      })
      currentFile = {
        ...currentFile,
        size: bytes.byteLength,
        version: `v${versionCounter}`,
        bytes,
        transport: 'buffer',
      }
      files.set(currentFile.id, currentFile)
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version-result',
        requestId: message.requestId,
        payload: { ok: true, file: descriptorOf(currentFile) },
      })
      setHostState(`history saved (${historyVersionCounter})`)
      render()
      return
    }
    case 'office:download-document': {
      if (message.payload.format !== 'xlsx') {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:download-document-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'DOWNLOAD_FAILED', error: 'Demo Host only downloads XLSX.' },
        })
        return
      }
      const name = normalizeXlsxName(message.payload.file.name || 'Untitled.xlsx')
      downloadBytes(message.payload.bytes.slice(0), name, XLSX_MIME)
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:download-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      setHostState('downloaded')
      render()
      return
    }
    case 'office:close-approved':
      setHostState(`close approved (${message.payload.reason})`)
      render()
      return
    case 'office:close-cancelled':
      setHostState('close cancelled')
      render()
      return

    // ---- protocol-v1 compatibility aliases ----
    case 'office:pick-file': {
      const selected = await pickBrowserFiles(message.payload)
      const officeFiles: OfficeFile[] = []
      for (const browserFile of selected || []) {
        const file = await toOfficeFile(browserFile)
        files.set(file.id, file)
        officeFiles.push(file)
      }
      const responseFiles = officeFiles.map((file) => ({
        ...descriptorOf(file),
        transport: 'buffer' as const,
        bytes: file.bytes.slice(0),
      }))
      send(
        {
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:pick-file-result',
          requestId: message.requestId,
          payload: { files: responseFiles.length ? responseFiles : null },
        },
        responseFiles.flatMap((file) => (file.bytes ? [file.bytes] : [])),
      )
      return
    }
    case 'office:read-file': {
      const stored = files.get(message.payload.fileId)
      if (!stored) {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:error',
          requestId: message.requestId,
          payload: { code: 'NOT_FOUND', message: 'Demo XLSX file was not found.' },
        })
        return
      }
      const bytes = stored.bytes.slice(0)
      send(
        {
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:read-file-result',
          requestId: message.requestId,
          payload: { file: { ...stored, bytes, transport: 'buffer' } },
        },
        [bytes],
      )
      return
    }
    case 'office:export-document': {
      const name = normalizeXlsxName(message.payload.file.name || 'Untitled.xlsx')
      downloadBytes(message.payload.bytes.slice(0), name, XLSX_MIME)
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:export-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      setHostState('downloaded')
      render()
      return
    }
    case 'office:close-request':
      setHostState(`close approved (${message.payload.reason})`)
      render()
      return
    case 'office:save-result':
      setHostState(message.payload.ok ? 'saved' : `save failed: ${message.payload.error ?? ''}`)
      render()
      return
    case 'office:state-result':
      dirty = message.payload.dirty
      mode = message.payload.mode
      setHostState(message.payload.saving ? 'saving' : hostStatus)
      render()
      return
    case 'office:error':
      setHostState(`error: ${message.payload.message}`)
      console.error('[GenOffice Excel Web]', message.payload)
      return
  }
}

window.addEventListener('message', (event) => {
  if (event.source !== frame.contentWindow || event.origin !== sheetsOrigin) return
  if (!isOfficeProtocolMessage(event.data)) return
  void handleEditorMessage(event.data as EditorToHostMessage)
})

newButton.addEventListener('click', sendNew)

picker.addEventListener('change', async () => {
  const browserFile = picker.files?.[0]
  if (!browserFile) return
  currentFile = await toOfficeFile(browserFile)
  files.set(currentFile.id, currentFile)
  dirty = false
  versionCounter = 1
  historyVersionCounter = 0
  render()
  sendInit()
})

saveButton.addEventListener('click', () => {
  if (!editorReady || mode === 'view') return
  setHostState('saving')
  render()
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:save',
    requestId: requestId('save'),
  })
})

downloadButton.addEventListener('click', downloadCurrentFile)

modeButton.addEventListener('click', () => {
  mode = mode === 'edit' ? 'view' : 'edit'
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:set-mode',
    requestId: requestId('mode'),
    payload: { mode },
  })
  render()
})

localeButton.addEventListener('click', () => {
  locale = locale === 'zh-CN' ? 'en-US' : 'zh-CN'
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:set-locale',
    requestId: requestId('locale'),
    payload: { locale },
  })
  render()
})

frame.src = `${sheetsUrl.replace(/\/$/, '')}/?hostOrigin=${encodeURIComponent(window.location.origin)}`
render()
