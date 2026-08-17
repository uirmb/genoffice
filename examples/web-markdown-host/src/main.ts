import type { OfficeFile, OfficeFileDescriptor } from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  isOfficeProtocolMessage,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'
import './style.css'

const MARKDOWN_MIME = 'text/markdown'
const DEFAULT_MARKDOWN_URL = 'http://localhost:5277'

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
}

const frame = requireElement<HTMLIFrameElement>('office-frame')
const markdownPicker = requireElement<HTMLInputElement>('markdown-picker')
const assetPicker = requireElement<HTMLInputElement>('asset-picker')
const fileName = requireElement<HTMLSpanElement>('file-name')
const hostState = requireElement<HTMLElement>('host-state')
const savedText = requireElement<HTMLElement>('saved-text')

const query = new URLSearchParams(window.location.search)
const markdownUrl = query.get('markdownUrl') || DEFAULT_MARKDOWN_URL
const markdownOrigin = new URL(markdownUrl).origin

let editorReady = false
let currentFile: OfficeFile | null = null
let pendingAssetRequestId: string | null = null
let saveVersion = 0
let requestSequence = 0

function requestId(prefix: string): string {
  requestSequence += 1
  return `${prefix}-${requestSequence}`
}

function send(message: HostToEditorMessage, transfer: Transferable[] = []): void {
  frame.contentWindow?.postMessage(message, markdownOrigin, transfer)
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

function capabilities() {
  return {
    ai: false,
    open: false,
    openDocument: false,
    pickAssets: true,
    save: true,
    saveAs: true,
    saveHistoryVersion: false,
    exportDocx: false,
    exportPptx: false,
    exportXlsx: false,
    close: true,
    autoSave: 'disabled' as const,
    download: false,
    print: false,
    systemFilePicker: true,
    pageCropMarks: false,
  }
}

async function browserFileToOfficeFile(file: File, prefix: string): Promise<OfficeFile> {
  const bytes = await file.arrayBuffer()
  return {
    id: `${prefix}:${crypto.randomUUID()}`,
    name: file.name,
    mimeType:
      file.type ||
      (file.name.match(/\.(md|markdown)$/i) ? MARKDOWN_MIME : 'application/octet-stream'),
    size: bytes.byteLength,
    version: String(file.lastModified || Date.now()),
    transport: 'buffer',
    bytes,
  }
}

function render(): void {
  fileName.textContent = currentFile?.name ?? '未选择 Markdown'
}

function setHostState(value: string): void {
  hostState.textContent = value
}

function sendInit(): void {
  if (!editorReady || !currentFile) return
  const bytes = currentFile.bytes.slice(0)
  setHostState('opening')
  send(
    {
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:init',
      requestId: requestId('init'),
      payload: {
        kind: 'markdown',
        mode: 'edit',
        locale: 'zh-CN',
        capabilities: capabilities(),
        file: { ...currentFile, bytes, transport: 'buffer' },
      },
    },
    [bytes],
  )
}

async function completeAssetPicker(file: File | null): Promise<void> {
  const requestIdValue = pendingAssetRequestId
  pendingAssetRequestId = null
  assetPicker.value = ''
  if (!requestIdValue) return

  if (!file) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-assets-result',
      requestId: requestIdValue,
      payload: { status: 'cancelled', files: [] },
    })
    return
  }

  try {
    const selected = await browserFileToOfficeFile(file, 'asset')
    const bytes = selected.bytes.slice(0)
    send(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-assets-result',
        requestId: requestIdValue,
        payload: {
          status: 'selected',
          files: [{ ...selected, bytes, transport: 'buffer' }],
        },
      },
      [bytes],
    )
  } catch (error) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-assets-result',
      requestId: requestIdValue,
      payload: {
        status: 'failed',
        files: [],
        code: 'PICK_ASSETS_FAILED',
        error: error instanceof Error ? error.message : String(error),
      },
    })
  }
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'markdown') return
      editorReady = true
      setHostState('ready')
      sendInit()
      return
    case 'office:dirty-change':
      setHostState(message.payload.dirty ? 'dirty' : 'clean')
      return
    case 'office:title-change':
      if (currentFile) currentFile = { ...currentFile, name: message.payload.title }
      render()
      return
    case 'office:save-document': {
      saveVersion += 1
      const bytes = message.payload.bytes.slice(0)
      const savedDescriptor: OfficeFileDescriptor = {
        ...message.payload.file,
        size: bytes.byteLength,
        version: `demo-${saveVersion}`,
        transport: 'buffer',
      }
      currentFile = { ...savedDescriptor, bytes, transport: 'buffer' }
      savedText.textContent =
        new TextDecoder().decode(bytes).replace(/\s+/g, ' ').trim() || '(empty)'
      render()
      setHostState('saved')
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: { ok: true, file: savedDescriptor },
      })
      return
    }
    case 'office:pick-assets':
      pendingAssetRequestId = message.requestId
      assetPicker.accept = message.payload.accept?.join(',') || 'image/png,image/jpeg,image/gif'
      assetPicker.click()
      return
    case 'office:close-approved':
      setHostState(`close approved (${message.payload.reason})`)
      return
    case 'office:close-cancelled':
      setHostState('close cancelled')
      return
    case 'office:state-result':
      setHostState(message.payload.dirty ? 'dirty' : 'clean')
      return
    case 'office:save-result':
      setHostState(message.payload.ok ? 'saved' : 'save failed')
      return
    case 'office:error':
      setHostState(`error: ${message.payload.message}`)
      console.error('[Markdown Web Host]', message.payload)
      return
    default:
      return
  }
}

window.addEventListener('message', (event) => {
  if (event.source !== frame.contentWindow || event.origin !== markdownOrigin) return
  if (!isOfficeProtocolMessage(event.data)) return
  void handleEditorMessage(event.data as EditorToHostMessage)
})

markdownPicker.addEventListener('change', async () => {
  const file = markdownPicker.files?.[0]
  if (!file) return
  currentFile = await browserFileToOfficeFile(file, 'markdown')
  render()
  sendInit()
})

assetPicker.addEventListener('change', () => {
  void completeAssetPicker(assetPicker.files?.[0] ?? null)
})
assetPicker.addEventListener('cancel', () => void completeAssetPicker(null))

frame.src = `${markdownUrl.replace(/\/$/, '')}/?hostOrigin=${encodeURIComponent(window.location.origin)}`
render()
