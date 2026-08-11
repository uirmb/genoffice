import type { OfficeEditorMode, OfficeFile } from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  isOfficeProtocolMessage,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const DEFAULT_SHEETS_URL = 'http://127.0.0.1:5275'

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id)
  if (!element) throw new Error(`Missing #${id}`)
  return element as T
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
const sheetsUrl = query.get('sheetsUrl') || DEFAULT_SHEETS_URL
const sheetsOrigin = new URL(sheetsUrl).origin

let editorReady = false
let currentFile: OfficeFile | null = null
let dirty = false
let mode: OfficeEditorMode = 'edit'
let locale = 'zh-CN'
let requestCounter = 0
let versionCounter = 0
let hostStatus = 'loading'
const files = new Map<string, OfficeFile>()

function requestId(prefix: string): string {
  requestCounter += 1
  return `${prefix}-${requestCounter}`
}

function send(message: HostToEditorMessage): void {
  frame.contentWindow?.postMessage(message, sheetsOrigin)
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

function sendNew(): void {
  if (!editorReady) return
  currentFile = null
  dirty = false
  versionCounter = 0
  setHostState('opening')
  send({
    protocol: OFFICE_PROTOCOL_VERSION,
    type: 'office:new',
    requestId: requestId('new'),
    payload: {
      kind: 'xlsx',
      mode,
      locale,
      capabilities: {
        ai: false,
        open: true,
        save: true,
        saveAs: true,
        saveHistoryVersion: true,
        exportDocx: false,
        exportPptx: false,
        exportXlsx: true,
        close: true,
        autoSave: 'host',
        download: false,
        print: true,
        systemFilePicker: true,
        pageCropMarks: false,
      },
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
      kind: 'xlsx',
      mode,
      locale,
      capabilities: {
        ai: false,
        open: true,
        save: true,
        saveAs: true,
        saveHistoryVersion: true,
        exportDocx: false,
        exportPptx: false,
        exportXlsx: true,
        close: true,
        autoSave: 'host',
        download: false,
        print: true,
        systemFilePicker: true,
        pageCropMarks: false,
      },
      file: { ...currentFile, bytes: currentFile.bytes.slice(0) },
    },
  })
}

function downloadCurrentFile(): void {
  if (!currentFile) return
  const blob = new Blob([currentFile.bytes], { type: currentFile.mimeType || XLSX_MIME })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = currentFile.name
  anchor.click()
  setTimeout(() => URL.revokeObjectURL(url), 0)
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
  }
}

function normalizeXlsxName(value: string): string {
  const name = value.trim()
  return /\.xlsx$/i.test(name) ? name : `${name}.xlsx`
}

async function handleEditorMessage(message: EditorToHostMessage): Promise<void> {
  switch (message.type) {
    case 'office:ready':
      if (message.payload.kind !== 'xlsx') return
      editorReady = true
      setHostState('ready')
      render()
      sendNew()
      break
    case 'office:dirty-change':
      dirty = message.payload.dirty
      render()
      break
    case 'office:title-change':
      if (currentFile) currentFile = { ...currentFile, name: message.payload.title }
      render()
      break
    case 'office:pick-file': {
      const input = document.createElement('input')
      input.type = 'file'
      input.multiple = message.payload.multiple === true
      if (message.payload.accept?.length) input.accept = message.payload.accept.join(',')
      input.addEventListener(
        'change',
        async () => {
          const selected = input.files?.[0]
          if (!selected) {
            send({
              protocol: OFFICE_PROTOCOL_VERSION,
              type: 'office:pick-file-result',
              requestId: message.requestId,
              payload: { files: null },
            })
            return
          }
          const officeFile = await toOfficeFile(selected)
          files.set(officeFile.id, officeFile)
          send({
            protocol: OFFICE_PROTOCOL_VERSION,
            type: 'office:pick-file-result',
            requestId: message.requestId,
            payload: {
              files: [
                {
                  id: officeFile.id,
                  name: officeFile.name,
                  mimeType: officeFile.mimeType,
                  size: officeFile.size,
                  version: officeFile.version,
                  transport: 'token',
                  token: `demo:${officeFile.id}`,
                },
              ],
            },
          })
        },
        { once: true },
      )
      input.click()
      break
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
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:read-file-result',
        requestId: message.requestId,
        payload: { file: { ...stored, bytes: stored.bytes.slice(0) } },
      })
      break
    }
    case 'office:save-document': {
      const saveAs = message.payload.mode === 'saveAs'
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
          break
        }
        name = normalizeXlsxName(requested)
      }

      versionCounter += 1
      const bytes = message.payload.bytes.slice(0)
      const id = saveAs || !currentFile ? `xlsx:${crypto.randomUUID()}` : currentFile.id
      currentFile = {
        id,
        name,
        mimeType: XLSX_MIME,
        size: bytes.byteLength,
        version: `v${versionCounter}`,
        bytes,
      }
      files.set(id, currentFile)
      dirty = false
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: {
          ok: true,
          file: {
            id,
            name,
            mimeType: XLSX_MIME,
            size: bytes.byteLength,
            version: currentFile.version,
          },
        },
      })
      setHostState('saved')
      render()
      break
    }
    case 'office:save-result':
      setHostState(message.payload.ok ? 'saved' : `save failed: ${message.payload.error ?? ''}`)
      render()
      break
    case 'office:state-result':
      dirty = message.payload.dirty
      mode = message.payload.mode
      setHostState(message.payload.saving ? 'saving' : hostStatus)
      render()
      break
    case 'office:error':
      setHostState(`error: ${message.payload.message}`)
      console.error('[GenOffice Excel Web]', message.payload)
      break
    default:
      break
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
