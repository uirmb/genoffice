import type {
  OfficeEditorMode,
  OfficeFile,
  PickFileOptions,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  isOfficeProtocolMessage,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const DEFAULT_DOCS_URL = 'http://localhost:5273'

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
}

const frame = requireElement<HTMLIFrameElement>('office-frame')
const docxPicker = requireElement<HTMLInputElement>('docx-picker')
const assetPicker = requireElement<HTMLInputElement>('asset-picker')
const saveButton = requireElement<HTMLButtonElement>('save-button')
const downloadButton = requireElement<HTMLButtonElement>('download-button')
const modeButton = requireElement<HTMLButtonElement>('mode-button')
const fileName = requireElement<HTMLSpanElement>('file-name')
const dirtyState = requireElement<HTMLSpanElement>('dirty-state')
const hostState = requireElement<HTMLSpanElement>('host-state')
const pickerDialog = requireElement<HTMLDivElement>('picker-dialog')
const pickerDescription = requireElement<HTMLParagraphElement>('picker-description')
const pickerCancel = requireElement<HTMLButtonElement>('picker-cancel')

const query = new URLSearchParams(window.location.search)
const docsUrl = query.get('docsUrl') || DEFAULT_DOCS_URL
const docsOrigin = new URL(docsUrl).origin

let editorReady = false
let currentFile: OfficeFile | null = null
let dirty = false
let mode: OfficeEditorMode = 'edit'
let requestCounter = 0
let versionCounter = 1
let pendingPicker: { requestId: string; options: PickFileOptions } | null = null
const files = new Map<string, OfficeFile>()

function requestId(prefix: string): string {
  requestCounter += 1
  return `${prefix}-${requestCounter}`
}

function send(message: HostToEditorMessage): void {
  frame.contentWindow?.postMessage(message, docsOrigin)
}

function setHostState(value: string): void {
  hostState.textContent = value
}

function render(): void {
  fileName.textContent = currentFile?.name ?? '未选择文档'
  dirtyState.textContent = dirty ? 'dirty' : 'clean'
  saveButton.disabled = !editorReady || !currentFile || mode === 'view'
  downloadButton.disabled = !currentFile
  modeButton.disabled = !editorReady || !currentFile
  modeButton.textContent = mode === 'edit' ? '切换为只读' : '切换为编辑'
}

function sendInit(): void {
  if (!editorReady || !currentFile) return
  setHostState('opening')
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:init',
    requestId: requestId('init'),
    payload: {
      kind: 'docx',
      mode,
      locale: 'zh-CN',
      file: { ...currentFile, bytes: currentFile.bytes.slice(0) },
    },
  })
}

function downloadCurrentFile(): void {
  if (!currentFile) return
  const blob = new Blob([currentFile.bytes], { type: currentFile.mimeType || DOCX_MIME })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = currentFile.name
  anchor.click()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}

async function browserFileToOfficeFile(file: File, prefix: string): Promise<OfficeFile> {
  const bytes = await file.arrayBuffer()
  return {
    id: `${prefix}:${crypto.randomUUID()}`,
    name: file.name,
    mimeType: file.type || 'application/octet-stream',
    size: bytes.byteLength,
    version: 'v1',
    bytes,
  }
}

function closePicker(): void {
  pendingPicker = null
  assetPicker.value = ''
  pickerDialog.hidden = true
}

async function completePicker(fileList: FileList | null): Promise<void> {
  if (!pendingPicker) return
  const { requestId: pickerRequestId, options } = pendingPicker
  const selected: SelectedOfficeFile[] = []

  if (fileList) {
    const limit = options.multiple ? fileList.length : Math.min(fileList.length, 1)
    for (let index = 0; index < limit; index += 1) {
      const browserFile = fileList.item(index)
      if (!browserFile) continue
      const officeFile = await browserFileToOfficeFile(browserFile, 'asset')
      files.set(officeFile.id, officeFile)
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
  }

  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:pick-file-result',
    requestId: pickerRequestId,
    payload: { files: selected.length > 0 ? selected : null },
  })
  closePicker()
}

function openPicker(requestIdValue: string, options: PickFileOptions): void {
  pendingPicker = { requestId: requestIdValue, options }
  assetPicker.accept = options.accept?.join(',') ?? ''
  assetPicker.multiple = Boolean(options.multiple)
  pickerDescription.textContent = `编辑器请求选择${options.multiple ? '文件' : '一个文件'}。这里用本地文件模拟 Web OS 系统文件选择器。`
  pickerDialog.hidden = false
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      editorReady = true
      setHostState('ready')
      render()
      sendInit()
      break
    case 'office:dirty-change':
      dirty = message.payload.dirty
      render()
      break
    case 'office:title-change':
      if (currentFile) currentFile = { ...currentFile, name: message.payload.title }
      render()
      break
    case 'office:save-document': {
      const bytes = message.payload.bytes.slice(0)
      versionCounter += 1
      const saved: OfficeFile = {
        ...message.payload.file,
        size: bytes.byteLength,
        version: `v${versionCounter}`,
        bytes,
      }
      currentFile = saved
      files.set(saved.id, saved)
      dirty = false
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: {
          ok: true,
          file: {
            id: saved.id,
            name: saved.name,
            mimeType: saved.mimeType,
            size: saved.size,
            version: saved.version,
          },
        },
      })
      setHostState('saved')
      render()
      break
    }
    case 'office:save-result':
      setHostState(message.payload.ok ? 'saved' : `save failed: ${message.payload.error ?? ''}`)
      break
    case 'office:pick-file':
      openPicker(message.requestId, message.payload)
      break
    case 'office:read-file': {
      const stored = files.get(message.payload.fileId)
      if (!stored) {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:error',
          requestId: message.requestId,
          payload: { code: 'NOT_FOUND', message: 'Demo host file was not found.' },
        })
        break
      }
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:read-file-result',
        requestId: message.requestId,
        payload: { file: { ...stored, bytes: stored.bytes.slice(0) } },
      })
      break
    }
    case 'office:state-result':
      dirty = message.payload.dirty
      setHostState(message.payload.saving ? 'saving' : 'ready')
      render()
      break
    case 'office:error':
      setHostState(`error: ${message.payload.message}`)
      console.error('[GenOffice Web]', message.payload)
      break
    default:
      break
  }
}

window.addEventListener('message', (event) => {
  if (event.source !== frame.contentWindow || event.origin !== docsOrigin) return
  if (!isOfficeProtocolMessage(event.data)) return
  void handleEditorMessage(event.data as EditorToHostMessage)
})

docxPicker.addEventListener('change', async () => {
  const browserFile = docxPicker.files?.[0]
  if (!browserFile) return
  const opened = await browserFileToOfficeFile(browserFile, 'doc')
  opened.mimeType = browserFile.type || DOCX_MIME
  versionCounter = 1
  opened.version = 'v1'
  currentFile = opened
  files.set(opened.id, opened)
  dirty = false
  render()
  sendInit()
})

assetPicker.addEventListener('change', () => {
  void completePicker(assetPicker.files)
})

pickerCancel.addEventListener('click', () => {
  void completePicker(null)
})

saveButton.addEventListener('click', () => {
  if (!currentFile || mode === 'view') return
  setHostState('saving')
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

frame.src = `${docsUrl.replace(/\/$/, '')}/?hostOrigin=${encodeURIComponent(window.location.origin)}`
render()
