import type {
  OfficeFile,
  OfficeFileDescriptor,
  PickDocumentOptions,
} from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  isOfficeProtocolMessage,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'

const PDF_MIME = 'application/pdf'
const DEFAULT_PDF_URL = 'http://localhost:5276'

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
}

const frame = requireElement<HTMLIFrameElement>('office-frame')
const pdfPicker = requireElement<HTMLInputElement>('pdf-picker')
const requestedPicker = requireElement<HTMLInputElement>('requested-picker')
const fileName = requireElement<HTMLSpanElement>('file-name')
const hostState = requireElement<HTMLElement>('host-state')

const query = new URLSearchParams(window.location.search)
const pdfUrl = query.get('pdfUrl') || DEFAULT_PDF_URL
const pdfOrigin = new URL(pdfUrl).origin

let editorReady = false
let currentFile: OfficeFile | null = null
let pendingPicker: { requestId: string; options: PickDocumentOptions } | null = null
let requestSequence = 0
const pendingDocuments = new Map<string, OfficeFile>()

function requestId(prefix: string): string {
  requestSequence += 1
  return `${prefix}-${requestSequence}`
}

function send(message: HostToEditorMessage, transfer: Transferable[] = []): void {
  frame.contentWindow?.postMessage(message, pdfOrigin, transfer)
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

function render(): void {
  fileName.textContent = currentFile?.name ?? '未选择 PDF'
}

function setHostState(value: string): void {
  hostState.textContent = value
}

function capabilities() {
  return {
    ai: false,
    open: true,
    openDocument: true,
    pickAssets: false,
    save: false,
    saveAs: false,
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
    mimeType: file.type || PDF_MIME,
    size: bytes.byteLength,
    version: String(file.lastModified || Date.now()),
    transport: 'buffer',
    bytes,
  }
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
        kind: 'pdf',
        mode: 'view',
        locale: 'zh-CN',
        capabilities: capabilities(),
        file: { ...currentFile, bytes, transport: 'buffer' },
      },
    },
    [bytes],
  )
}

function completePickerWithCancellation(): void {
  const pending = pendingPicker
  pendingPicker = null
  requestedPicker.value = ''
  if (!pending) return
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:pick-document-result',
    requestId: pending.requestId,
    payload: { status: 'cancelled', selectionId: null, file: null },
  })
}

async function completeRequestedPicker(file: File | null): Promise<void> {
  const pending = pendingPicker
  pendingPicker = null
  requestedPicker.value = ''
  if (!pending) return
  if (!file) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-document-result',
      requestId: pending.requestId,
      payload: { status: 'cancelled', selectionId: null, file: null },
    })
    return
  }

  try {
    const selected = await browserFileToOfficeFile(file, 'picked-pdf')
    const selectionId = `selection:${crypto.randomUUID()}`
    pendingDocuments.set(selectionId, selected)
    const bytes = selected.bytes.slice(0)
    send(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:pick-document-result',
        requestId: pending.requestId,
        payload: {
          status: 'selected',
          selectionId,
          file: { ...selected, bytes, transport: 'buffer' },
        },
      },
      [bytes],
    )
  } catch (error) {
    send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:pick-document-result',
      requestId: pending.requestId,
      payload: {
        status: 'failed',
        code: 'PICK_DOCUMENT_FAILED',
        error: error instanceof Error ? error.message : String(error),
      },
    })
  }
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'pdf') return
      editorReady = true
      setHostState('ready')
      sendInit()
      return
    case 'office:dirty-change':
      if (message.payload.dirty) setHostState('unexpected dirty state')
      return
    case 'office:title-change':
      if (currentFile) currentFile = { ...currentFile, name: message.payload.title }
      render()
      setHostState('opened')
      return
    case 'office:pick-document':
      pendingPicker = { requestId: message.requestId, options: message.payload }
      requestedPicker.accept = message.payload.accept?.join(',') || `${PDF_MIME},.pdf`
      requestedPicker.click()
      return
    case 'office:document-opened': {
      const selected = pendingDocuments.get(message.payload.selectionId)
      if (!selected) {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:document-opened-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'NOT_FOUND', error: 'PDF selection expired.' },
        })
        return
      }
      pendingDocuments.delete(message.payload.selectionId)
      currentFile = selected
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:document-opened-result',
        requestId: message.requestId,
        payload: { ok: true, file: descriptorOf(selected) },
      })
      render()
      setHostState('opened')
      return
    }
    case 'office:document-open-failed':
      pendingDocuments.delete(message.payload.selectionId)
      setHostState('open failed')
      return
    case 'office:close-approved':
      setHostState(`close approved (${message.payload.reason})`)
      return
    case 'office:close-cancelled':
      setHostState('close cancelled')
      return
    case 'office:state-result':
      setHostState(message.payload.dirty ? 'unexpected dirty state' : 'ready')
      return
    case 'office:save-result':
      setHostState(message.payload.ok ? 'unexpected save' : 'read-only')
      return
    case 'office:error':
      setHostState(`error: ${message.payload.message}`)
      console.error('[PDF Web Host]', message.payload)
      return
    default:
      return
  }
}

window.addEventListener('message', (event) => {
  if (event.source !== frame.contentWindow || event.origin !== pdfOrigin) return
  if (!isOfficeProtocolMessage(event.data)) return
  void handleEditorMessage(event.data as EditorToHostMessage)
})

pdfPicker.addEventListener('change', async () => {
  const file = pdfPicker.files?.[0]
  if (!file) return
  currentFile = await browserFileToOfficeFile(file, 'pdf')
  render()
  sendInit()
})

requestedPicker.addEventListener('change', () => {
  void completeRequestedPicker(requestedPicker.files?.[0] ?? null)
})
requestedPicker.addEventListener('cancel', completePickerWithCancellation)

frame.src = `${pdfUrl.replace(/\/$/, '')}/?hostOrigin=${encodeURIComponent(window.location.origin)}`
render()
