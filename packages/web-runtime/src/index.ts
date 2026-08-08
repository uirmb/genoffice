import { OfficeIframeBridge, createOfficeRequestId } from '@genoffice/iframe-bridge'
import type {
  OfficeFile,
  OfficeHostApi,
  PickFileOptions,
  SaveDocumentInput,
  SaveDocumentResult,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  type EditorToHostMessage,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'

export type WebRuntimeMode = 'standalone' | 'embedded'
export type EditorIframeBridge = OfficeIframeBridge<HostToEditorMessage, EditorToHostMessage>

export function detectWebRuntimeMode(currentWindow: Window = window): WebRuntimeMode {
  return currentWindow.parent === currentWindow ? 'standalone' : 'embedded'
}

function localFileId(file: File): string {
  const suffix =
    typeof crypto !== 'undefined' && 'randomUUID' in crypto
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(36).slice(2)}`
  return `local:${file.name}:${file.size}:${file.lastModified}:${suffix}`
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

export class StandaloneOfficeHost implements OfficeHostApi {
  private readonly files = new Map<string, OfficeFile>()
  private dirty = false

  async getLocale(): Promise<string> {
    return document.documentElement.lang || navigator.language || 'en'
  }

  async saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult> {
    downloadBuffer(input.bytes, input.file.name, input.file.mimeType)
    return {
      ok: true,
      file: {
        ...input.file,
        version: input.baseVersion ?? input.file.version ?? null,
      },
    }
  }

  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {
    const input = document.createElement('input')
    input.type = 'file'
    input.multiple = options.multiple === true
    if (options.accept?.length) input.accept = options.accept.join(',')

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
      input.style.display = 'none'
      document.body.append(input)
      input.click()
    })

    if (!files?.length) return null

    const selected: SelectedOfficeFile[] = []
    for (const file of files) {
      const id = localFileId(file)
      const bytes = await file.arrayBuffer()
      const officeFile: OfficeFile = {
        id,
        name: file.name,
        mimeType: file.type || 'application/octet-stream',
        size: file.size,
        version: String(file.lastModified),
        bytes,
      }
      this.files.set(id, officeFile)
      selected.push({
        id,
        name: officeFile.name,
        mimeType: officeFile.mimeType,
        size: officeFile.size,
        version: officeFile.version,
        transport: 'buffer',
        bytes,
      })
    }
    return selected
  }

  async readFile(fileId: string): Promise<OfficeFile> {
    const file = this.files.get(fileId)
    if (!file) throw new Error(`Standalone file is not available: ${fileId}`)
    return file
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
    this.files.clear()
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
  constructor(
    private readonly bridge: EditorIframeBridge,
    private readonly locale?: string,
  ) {}

  async getLocale(): Promise<string> {
    return this.locale || document.documentElement.lang || navigator.language || 'en'
  }

  async saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult> {
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
        },
      },
      'office:save-document-result',
      [bytes],
    )
    return response.payload
  }

  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {
    const requestId = createOfficeRequestId('pick')
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
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:title-change',
      payload: { title },
    })
  }
}

export interface EmbeddedOfficeRuntime {
  bridge: EditorIframeBridge
  host: EmbeddedOfficeHost
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
  bridge.start()
  const host = new EmbeddedOfficeHost(bridge, options.locale)

  return {
    bridge,
    host,
    destroy: () => bridge.destroy(),
  }
}
