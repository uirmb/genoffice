import type {
  OfficeEditorMode,
  OfficeFile,
  OfficeFileDescriptor,
  PickAssetsOptions,
  PickDocumentOptions,
  PickFileOptions,
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
const files = new Map<string, OfficeFile>()
const history = new Map<string, ArrayBuffer[]>()
const pendingDocuments = new Map<string, OfficeFile>()

type PendingPicker =
  | { kind: 'document'; requestId: string; options: PickDocumentOptions }
  | { kind: 'assets'; requestId: string; options: PickAssetsOptions }
  | { kind: 'legacy'; requestId: string; options: PickFileOptions }

let pendingPicker: PendingPicker | null = null

function requestId(prefix: string): string {
  requestCounter += 1
  return `${prefix}-${requestCounter}`
}

function send(message: HostToEditorMessage, transfer: Transferable[] = []): void {
  frame.contentWindow?.postMessage(message, slidesOrigin, transfer)
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
    openDocument: true,
    pickAssets: true,
    save: true,
    saveAs: true,
    saveHistoryVersion: true,
    exportDocx: false,
    exportPptx: true,
    close: true,
    autoSave: 'host' as const,
    download: true,
    print: false,
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
  const bytes = currentFile.bytes.slice(0)
  send(
    {
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: requestId('init'),
      payload: {
        kind: 'pptx',
        mode,
        locale,
        capabilities: capabilities(),
        file: { ...currentFile, bytes, transport: 'buffer' },
      },
    },
    [bytes],
  )
}

function downloadBytes(name: string, mimeType: string, bytes: ArrayBuffer): void {
  const blob = new Blob([bytes], { type: mimeType || PPTX_MIME })
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
    transport: 'buffer',
  }
}

function closePicker(): void {
  pendingPicker = null
  assetPicker.value = ''
  pickerDialog.hidden = true
}

function openPicker(pending: PendingPicker): void {
  pendingPicker = pending
  const options = pending.options
  assetPicker.accept = options.accept?.join(',') ?? ''
  assetPicker.multiple = pending.kind === 'assets' ? Boolean(options.multiple) : false
  pickerDescription.textContent =
    pending.kind === 'document'
      ? '编辑器请求打开另一个 PPTX。这里用本地文件模拟系统文件选择器。'
      : pending.kind === 'assets'
        ? `编辑器请求选择${options.multiple ? '多个素材' : '一个素材'}。这里用本地文件模拟系统素材选择器。`
        : `兼容调用请求选择${options.multiple ? '多个文件' : '一个文件'}。`
  pickerDialog.hidden = false
}

async function completePicker(fileList: FileList | null): Promise<void> {
  const pending = pendingPicker
  if (!pending) return

  if (pending.kind === 'document') {
    const browserFile = fileList?.item(0)
    if (!browserFile) {
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-document-result',
        requestId: pending.requestId,
        payload: { status: 'cancelled', selectionId: null, file: null },
      })
      closePicker()
      return
    }
    const file = await browserFileToOfficeFile(browserFile, 'ppt')
    file.mimeType = browserFile.type || PPTX_MIME
    const selectionId = `selection:${crypto.randomUUID()}`
    pendingDocuments.set(selectionId, file)
    const bytes = file.bytes.slice(0)
    send(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-document-result',
        requestId: pending.requestId,
        payload: {
          status: 'selected',
          selectionId,
          file: { ...file, bytes, transport: 'buffer' },
        },
      },
      [bytes],
    )
    closePicker()
    return
  }

  const officeFiles: OfficeFile[] = []
  if (fileList) {
    const limit = pending.options.multiple ? fileList.length : Math.min(fileList.length, 1)
    for (let index = 0; index < limit; index += 1) {
      const browserFile = fileList.item(index)
      if (!browserFile) continue
      officeFiles.push(await browserFileToOfficeFile(browserFile, 'asset'))
    }
  }

  if (pending.kind === 'assets') {
    const responseFiles = officeFiles.map((file) => ({
      ...file,
      bytes: file.bytes.slice(0),
      transport: 'buffer' as const,
    }))
    send(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-assets-result',
        requestId: pending.requestId,
        payload: {
          status: responseFiles.length ? 'selected' : 'cancelled',
          files: responseFiles,
        },
      },
      responseFiles.map((file) => file.bytes),
    )
    closePicker()
    return
  }

  for (const file of officeFiles) files.set(file.id, file)
  const legacyFiles = officeFiles.map((file) => ({
    ...descriptorOf(file),
    transport: 'buffer' as const,
    bytes: file.bytes.slice(0),
  }))
  send(
    {
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-file-result',
      requestId: pending.requestId,
      payload: { files: legacyFiles.length ? legacyFiles : null },
    },
    legacyFiles.flatMap((file) => (file.bytes ? [file.bytes] : [])),
  )
  closePicker()
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
      if (message.payload.kind !== 'pptx') return
      editorReady = true
      setHostState('ready')
      render()
      if (currentFile) sendInit()
      else sendNew()
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
      openPicker({ kind: 'document', requestId: message.requestId, options: message.payload })
      return
    case 'office:document-opened': {
      const selected = pendingDocuments.get(message.payload.selectionId)
      if (!selected) {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:document-opened-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'NOT_FOUND', error: 'Document selection expired.' },
        })
        return
      }
      pendingDocuments.delete(message.payload.selectionId)
      currentFile = selected
      files.set(selected.id, selected)
      dirty = false
      versionCounter = 1
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:document-opened-result',
        requestId: message.requestId,
        payload: { ok: true, file: descriptorOf(selected) },
      })
      setHostState('opened')
      render()
      return
    }
    case 'office:document-open-failed':
      pendingDocuments.delete(message.payload.selectionId)
      return
    case 'office:pick-assets':
      openPicker({ kind: 'assets', requestId: message.requestId, options: message.payload })
      return
    case 'office:save-document': {
      const bytes = message.payload.bytes.slice(0)
      const needsDestination =
        Boolean(message.payload.newDocument) || message.payload.mode === 'saveAs' || !currentFile
      let targetName = message.payload.file.name || currentFile?.name || 'Untitled.pptx'
      if (needsDestination) {
        const requested = window.prompt(
          message.payload.newDocument ? '保存新演示文稿' : '另存为文件名',
          targetName,
        )
        if (requested === null || !requested.trim()) {
          saveCancelled(message.requestId)
          return
        }
        targetName = /\.pptx$/i.test(requested.trim())
          ? requested.trim()
          : `${requested.trim()}.pptx`
        versionCounter = 0
      }

      versionCounter += 1
      const saved: OfficeFile = {
        ...message.payload.file,
        id: needsDestination ? `ppt:${crypto.randomUUID()}` : currentFile?.id || message.payload.file.id,
        name: targetName,
        mimeType: PPTX_MIME,
        size: bytes.byteLength,
        version: `v${versionCounter}`,
        bytes,
        transport: 'buffer',
      }
      currentFile = saved
      files.set(saved.id, saved)
      dirty = false
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: { ok: true, file: descriptorOf(saved) },
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
          payload: { ok: false, code: 'NOT_FOUND', error: 'Save the presentation before creating history.' },
        })
        return
      }
      const key = currentFile.id
      const list = history.get(key) ?? []
      const bytes = message.payload.bytes.slice(0)
      list.push(bytes)
      history.set(key, list)
      versionCounter += 1
      currentFile = {
        ...currentFile,
        size: bytes.byteLength,
        version: `v${versionCounter}`,
        bytes,
        transport: 'buffer',
      }
      files.set(key, currentFile)
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version-result',
        requestId: message.requestId,
        payload: { ok: true, file: descriptorOf(currentFile) },
      })
      setHostState(`history:${list.length}`)
      render()
      return
    }
    case 'office:download-document':
      if (message.payload.format !== 'pptx') {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:download-document-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'DOWNLOAD_FAILED', error: 'Demo Host only downloads PPTX.' },
        })
        return
      }
      downloadBytes(
        message.payload.file.name || 'presentation.pptx',
        message.payload.file.mimeType || PPTX_MIME,
        message.payload.bytes,
      )
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:download-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      setHostState('downloaded')
      return
    case 'office:close-approved':
      setHostState(`close approved (${message.payload.reason})`)
      return
    case 'office:close-cancelled':
      setHostState('close cancelled')
      return

    // ---- protocol-v1 compatibility aliases ----
    case 'office:pick-file':
      openPicker({ kind: 'legacy', requestId: message.requestId, options: message.payload })
      return
    case 'office:read-file': {
      const stored = files.get(message.payload.fileId)
      if (!stored) {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:error',
          requestId: message.requestId,
          payload: { code: 'NOT_FOUND', message: 'Demo host file was not found.' },
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
    case 'office:export-document':
      downloadBytes(
        message.payload.file.name || 'presentation.pptx',
        message.payload.file.mimeType || PPTX_MIME,
        message.payload.bytes,
      )
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:export-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      setHostState('downloaded')
      return
    case 'office:close-request':
      setHostState(`close approved (${message.payload.reason})`)
      return
    case 'office:save-result':
      setHostState(message.payload.ok ? 'saved' : `save failed: ${message.payload.error ?? ''}`)
      return
    case 'office:state-result':
      dirty = message.payload.dirty
      setHostState(message.payload.saving ? 'saving' : 'ready')
      render()
      return
    case 'office:error':
      setHostState(`error: ${message.payload.message}`)
      console.error('[GenOffice PPT Web]', message.payload)
      return
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
