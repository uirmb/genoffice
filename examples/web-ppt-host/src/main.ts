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

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const DEFAULT_SLIDES_URL = 'http://localhost:5274'

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
}

const frame = requireElement<HTMLIFrameElement>('office-frame')
const pptxPicker = requireElement<HTMLInputElement>('pptx-picker')
const assetPicker = requireElement<HTMLInputElement>('asset-picker')
const newButton = requireElement<HTMLButtonElement>('new-button')
const saveButton = requireElement<HTMLButtonElement>('save-button')
const downloadButton = requireElement<HTMLButtonElement>('download-button')
const modeButton = requireElement<HTMLButtonElement>('mode-button')
const localeButton = requireElement<HTMLButtonElement>('locale-button')
const fileName = requireElement<HTMLSpanElement>('file-name')
const dirtyState = requireElement<HTMLSpanElement>('dirty-state')
const hostState = requireElement<HTMLSpanElement>('host-state')
const pickerDialog = requireElement<HTMLDivElement>('picker-dialog')
const pickerDescription = requireElement<HTMLParagraphElement>('picker-description')
const pickerCancel = requireElement<HTMLButtonElement>('picker-cancel')

const query = new URLSearchParams(window.location.search)
const slidesUrl = query.get('slidesUrl') || DEFAULT_SLIDES_URL
const slidesOrigin = new URL(slidesUrl).origin

let editorReady = false
let currentFile: OfficeFile | null = null
let dirty = false
let mode: OfficeEditorMode = 'edit'
let locale = 'zh-CN'
let requestCounter = 0
let versionCounter = 1
let pendingPicker: { requestId: string; options: PickFileOptions } | null = null
const files = new Map<string, OfficeFile>()
const history = new Map<string, ArrayBuffer[]>()

function requestId(prefix: string): string {
  requestCounter += 1
  return `${prefix}-${requestCounter}`
}

function send(message: HostToEditorMessage): void {
  frame.contentWindow?.postMessage(message, slidesOrigin)
}

function setHostState(value: string): void {
  hostState.textContent = value
}

function render(): void {
  fileName.textContent = currentFile?.name ?? '新建演示文稿'
  dirtyState.textContent = dirty ? 'dirty' : 'clean'
  saveButton.disabled = !editorReady || mode === 'view'
  downloadButton.disabled = !currentFile
  modeButton.disabled = !editorReady
  localeButton.disabled = !editorReady
  modeButton.textContent = mode === 'edit' ? '切换为预览' : '切换为编辑'
  localeButton.textContent = locale.startsWith('zh') ? '切换 English' : '切换中文'
}

function capabilities() {
  return {
    ai: false,
    open: true,
    save: true,
    saveAs: true,
    saveHistoryVersion: true,
    exportDocx: false,
    exportPptx: true,
    close: true,
    autoSave: 'host' as const,
    download: false,
    print: false,
    systemFilePicker: true,
    pageCropMarks: false,
  }
}

function sendNew(): void {
  if (!editorReady) return
  currentFile = null
  dirty = false
  versionCounter = 1
  setHostState('new')
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:new',
    requestId: requestId('new'),
    payload: {
      kind: 'pptx',
      mode,
      locale,
      capabilities: capabilities(),
    },
  })
  render()
}

function sendInit(): void {
  if (!editorReady || !currentFile) return
  setHostState('opening')
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:init',
    requestId: requestId('init'),
    payload: {
      kind: 'pptx',
      mode,
      locale,
      capabilities: capabilities(),
      file: { ...currentFile, bytes: currentFile.bytes.slice(0) },
    },
  })
}

function downloadBytes(name: string, mimeType: string, bytes: ArrayBuffer): void {
  const blob = new Blob([bytes], { type: mimeType })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = name
  anchor.click()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}

function downloadCurrentFile(): void {
  if (!currentFile) return
  downloadBytes(currentFile.name, currentFile.mimeType || PPTX_MIME, currentFile.bytes)
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
  pickerDescription.textContent = `编辑器请求选择${options.multiple ? '文件' : '一个文件'}。这里用本地文件模拟 Web OS 文件管理器。`
  pickerDialog.hidden = false
}

function saveCancelled(requestIdValue: string): void {
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:save-document-result',
    requestId: requestIdValue,
    payload: { ok: false, code: 'CANCELLED', error: '已取消保存。' },
  })
  setHostState('ready')
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'pptx') break
      editorReady = true
      setHostState('ready')
      render()
      if (currentFile) sendInit()
      else sendNew()
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
      const needsDestination =
        Boolean(message.payload.newDocument) || message.payload.mode === 'saveAs'
      let targetName = message.payload.file.name || currentFile?.name || 'Untitled.pptx'
      if (needsDestination) {
        const requested = window.prompt(
          message.payload.newDocument ? '保存新演示文稿' : '另存为文件名',
          targetName,
        )
        if (requested === null || !requested.trim()) {
          saveCancelled(message.requestId)
          break
        }
        targetName = /\.pptx$/i.test(requested.trim())
          ? requested.trim()
          : `${requested.trim()}.pptx`
        versionCounter = 1
      } else {
        versionCounter += 1
      }

      const saved: OfficeFile = {
        ...message.payload.file,
        id: needsDestination ? `ppt:${crypto.randomUUID()}` : message.payload.file.id,
        name: targetName,
        mimeType: PPTX_MIME,
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

    case 'office:save-history-version': {
      const key = message.payload.file.id
      const list = history.get(key) ?? []
      list.push(message.payload.bytes.slice(0))
      history.set(key, list)
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      setHostState(`history:${list.length}`)
      break
    }

    case 'office:export-document':
      if (message.payload.format !== 'pptx') break
      downloadBytes(
        message.payload.file.name || 'presentation.pptx',
        PPTX_MIME,
        message.payload.bytes,
      )
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:export-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      setHostState('exported')
      break

    case 'office:close-request':
      setHostState('close requested')
      break

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
      console.error('[GenOffice PPT Web]', message.payload)
      break

    default:
      break
  }
}

window.addEventListener('message', (event) => {
  if (event.source !== frame.contentWindow || event.origin !== slidesOrigin) return
  if (!isOfficeProtocolMessage(event.data)) return
  void handleEditorMessage(event.data as EditorToHostMessage)
})

pptxPicker.addEventListener('change', async () => {
  const browserFile = pptxPicker.files?.[0]
  if (!browserFile) return
  const opened = await browserFileToOfficeFile(browserFile, 'ppt')
  opened.mimeType = browserFile.type || PPTX_MIME
  opened.version = 'v1'
  versionCounter = 1
  currentFile = opened
  files.set(opened.id, opened)
  dirty = false
  render()
  sendInit()
})

assetPicker.addEventListener('change', () => void completePicker(assetPicker.files))
pickerCancel.addEventListener('click', () => void completePicker(null))

newButton.addEventListener('click', sendNew)

saveButton.addEventListener('click', () => {
  if (mode === 'view') return
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

localeButton.addEventListener('click', () => {
  locale = locale.startsWith('zh') ? 'en-US' : 'zh-CN'
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:set-locale',
    requestId: requestId('locale'),
    payload: { locale },
  })
  render()
})

frame.src = `${slidesUrl.replace(/\/$/, '')}/?hostOrigin=${encodeURIComponent(window.location.origin)}`
render()
