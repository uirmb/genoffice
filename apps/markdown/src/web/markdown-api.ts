import { normalizeLang, type Lang } from '@genoffice/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@genoffice/ai-provider'
import type {
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostApi,
  PickAssetsResult,
} from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION } from '@genoffice/office-protocol'
import type { EditorIframeBridge, WebRuntimeMode } from '@genoffice/web-runtime'
import type {
  ExportDocxRequest,
  ExportPdfRequest,
  ExportResult,
  ImageData,
  MarkdownApi,
  SaveMarkdownRequest,
  SaveMarkdownResult,
  SaveMode,
  WebSearchResult,
} from '../shared/ipc'

const MARKDOWN_ACCEPT = ['text/markdown', 'text/plain', '.md', '.markdown']
const IMAGE_ACCEPT = ['image/png', 'image/jpeg', 'image/gif', '.png', '.jpg', '.jpeg', '.gif']

function cloneFile(file: OfficeFile): OfficeFile {
  return { ...file, bytes: file.bytes.slice(0), transport: 'buffer' }
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

function virtualPath(file: OfficeFile | OfficeFileDescriptor): string {
  const safeId = encodeURIComponent(file.nodeId || file.id)
  return `/uc-web/${safeId}/${file.name}`
}

function ensureMarkdownName(name: string): string {
  const trimmed = name.trim() || 'Untitled'
  return /\.(md|markdown)$/i.test(trimmed) ? trimmed : `${trimmed}.md`
}

function bytesToBase64(bytes: ArrayBuffer): string {
  const view = new Uint8Array(bytes)
  let binary = ''
  for (let offset = 0; offset < view.length; offset += 0x8000) {
    binary += String.fromCharCode(...view.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function imageDataUrl(file: OfficeFile): string | null {
  if (!/^image\/(png|jpeg|gif)$/i.test(file.mimeType)) return null
  return `data:${file.mimeType};base64,${bytesToBase64(file.bytes)}`
}

function dataUrlToImageData(src: string): ImageData | null {
  const match = /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/i.exec(src)
  if (!match) return null
  return { base64: match[2]!, mime: match[1]!.toLowerCase() as ImageData['mime'] }
}

function unsupportedExport(): ExportResult {
  return { ok: false, error: 'Markdown Web export will be added after the core UC Host workflow.' }
}

export class MarkdownWebApi implements MarkdownApi {
  private currentFile: OfficeFile | null = null
  private dirty = false
  private pendingResolved = false
  private pendingResolve!: (value: string | null) => void
  private readonly pendingPath = new Promise<string | null>((resolve) => {
    this.pendingResolve = resolve
  })

  private saveHandler: ((mode: SaveMode) => void) | null = null
  private closeSaveHandler: (() => void) | null = null
  private saveRequestId: string | null = null
  private closeRequestId: string | null = null
  private languageListeners = new Set<(lang: Lang) => void>()
  private unsubscribeBridge: (() => void) | null = null

  constructor(
    private readonly host: OfficeHostApi,
    private readonly bridge: EditorIframeBridge | null,
    private readonly runtimeMode: WebRuntimeMode,
  ) {
    if (runtimeMode === 'standalone') this.resolvePending(null)

    if (bridge) {
      this.unsubscribeBridge = bridge.subscribe((message) => {
        switch (message.type) {
          case 'office:init':
            if (message.payload.kind !== 'markdown') {
              bridge.send({
                protocol: OFFICE_PROTOCOL_VERSION,
                type: 'office:error',
                requestId: message.requestId,
                payload: {
                  code: 'UNSUPPORTED_KIND',
                  message: 'Markdown Web only accepts Markdown documents.',
                },
              })
              return
            }
            if (this.pendingResolved) {
              bridge.send({
                protocol: OFFICE_PROTOCOL_VERSION,
                type: 'office:error',
                requestId: message.requestId,
                payload: {
                  code: 'RELOAD_REQUIRED',
                  message: 'Reload the Markdown iframe before binding another initial document.',
                },
              })
              return
            }
            this.currentFile = cloneFile(message.payload.file)
            this.host.setTitle(this.currentFile.name)
            this.resolvePending(virtualPath(this.currentFile))
            return
          case 'office:new':
            if (message.payload.kind !== 'markdown') return
            if (!this.pendingResolved) this.resolvePending(null)
            return
          case 'office:set-locale': {
            const lang = normalizeLang(message.payload.locale)
            for (const listener of this.languageListeners) listener(lang)
            return
          }
          case 'office:set-mode':
            return
          case 'office:save':
            this.saveRequestId = message.requestId
            if (this.saveHandler) this.saveHandler('save')
            else this.sendSaveRequestAck(false)
            return
          case 'office:query-state':
            bridge.send({
              protocol: OFFICE_PROTOCOL_VERSION,
              type: 'office:state-result',
              requestId: message.requestId,
              payload: {
                ready: true,
                dirty: this.dirty,
                saving: false,
                mode: 'edit',
                title: this.currentFile?.name,
              },
            })
            return
          case 'office:request-close':
            if (!this.dirty) {
              void this.host.approveClose?.(message.requestId)
              return
            }
            this.closeRequestId = message.requestId
            if (this.closeSaveHandler) this.closeSaveHandler()
            else void this.host.cancelClose?.(message.requestId)
            return
          default:
            return
        }
      })

      bridge.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:ready',
        payload: { kind: 'markdown' },
      })
    }
  }

  private resolvePending(value: string | null): void {
    if (this.pendingResolved) return
    this.pendingResolved = true
    this.pendingResolve(value)
  }

  async consumePending(): Promise<string | null> {
    return this.pendingPath
  }

  async readFile(_path: string): Promise<string> {
    if (!this.currentFile) return ''
    return new TextDecoder('utf-8').decode(this.currentFile.bytes)
  }

  async save(request: SaveMarkdownRequest): Promise<SaveMarkdownResult> {
    const bytes = new TextEncoder().encode(request.text).buffer
    const existing = this.currentFile
    const name = ensureMarkdownName(
      existing?.name || request.suggestedName || (this.runtimeMode === 'standalone' ? 'Untitled.md' : 'Untitled.md'),
    )
    const file: OfficeFileDescriptor = existing
      ? descriptorOf(existing)
      : {
          id: 'new:markdown',
          name,
          mimeType: 'text/markdown',
          size: 0,
          version: null,
          transport: 'buffer',
        }

    const result = await this.host.saveDocument({
      file,
      bytes,
      baseVersion: existing?.version,
      mode: request.mode,
      newDocument: !existing,
    })

    if (!result.ok) {
      if (result.code === 'CANCELLED') return { ok: true, canceled: true }
      return { ok: false, error: result.error || 'Markdown save failed.' }
    }

    const saved = result.file ?? { ...file, size: bytes.byteLength }
    this.currentFile = {
      ...saved,
      mimeType: saved.mimeType || 'text/markdown',
      size: bytes.byteLength,
      bytes: bytes.slice(0),
      transport: 'buffer',
    }
    this.host.setTitle(this.currentFile.name)
    return { ok: true, path: virtualPath(this.currentFile) }
  }

  setDirty(dirty: boolean): void {
    this.dirty = dirty
    this.host.setDirty(dirty)
  }

  onSaveRequest(handler: (mode: SaveMode) => void): () => void {
    this.saveHandler = handler
    return () => {
      if (this.saveHandler === handler) this.saveHandler = null
    }
  }

  sendSaveRequestAck(ok: boolean): void {
    const requestId = this.saveRequestId
    this.saveRequestId = null
    if (!requestId || !this.bridge) return
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:save-result',
      requestId,
      payload: { ok, ...(ok ? {} : { error: 'Markdown save failed.' }) },
    })
  }

  onCloseSaveRequest(handler: () => void): () => void {
    this.closeSaveHandler = handler
    return () => {
      if (this.closeSaveHandler === handler) this.closeSaveHandler = null
    }
  }

  sendCloseSaveResult(ok: boolean): void {
    const requestId = this.closeRequestId
    this.closeRequestId = null
    if (!requestId) return
    if (ok) void this.host.approveClose?.(requestId)
    else void this.host.cancelClose?.(requestId)
  }

  onFileRenamed(_handler: (newPath: string) => void): () => void {
    return () => {}
  }

  async pickImage(): Promise<string | null> {
    const result = await this.pickImages()
    if (result.status !== 'selected') return null
    return result.files[0] ? imageDataUrl(result.files[0]) : null
  }

  async saveImage(data: { base64: string; ext: string }): Promise<string | null> {
    const ext = data.ext.toLowerCase()
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/png'
    return `data:${mime};base64,${data.base64}`
  }

  async readImage(src: string): Promise<ImageData | null> {
    return dataUrlToImageData(src)
  }

  private async pickImages(): Promise<PickAssetsResult> {
    if (this.host.pickAssets) {
      return this.host.pickAssets({ multiple: false, accept: IMAGE_ACCEPT })
    }

    const files = await this.host.pickFile({ multiple: false, accept: IMAGE_ACCEPT, mode: 'file' })
    const first = files?.[0]
    if (!first) return { status: 'cancelled', files: [] }
    const file =
      first.transport === 'buffer' && first.bytes
        ? ({ ...first, bytes: first.bytes, transport: 'buffer' } as OfficeFile)
        : await this.host.readFile(first.id)
    return { status: 'selected', files: [file] }
  }

  onExportRequest(_handler: (format: 'pdf' | 'docx' | 'docs') => void): () => void {
    return () => {}
  }

  async exportDocx(_request: ExportDocxRequest): Promise<ExportResult> {
    return unsupportedExport()
  }

  async exportPdf(_request: ExportPdfRequest): Promise<ExportResult> {
    return unsupportedExport()
  }

  async getLanguage(): Promise<Lang> {
    return normalizeLang(await this.host.getLocale())
  }

  onLanguageChanged(handler: (lang: Lang) => void): () => void {
    this.languageListeners.add(handler)
    return () => this.languageListeners.delete(handler)
  }

  async getAiSettings(): Promise<AiSettings> {
    throw new Error('AI is disabled in Markdown Web until office-agent-protocol is integrated.')
  }

  async aiStream(_request: AiStreamRequest): Promise<void> {
    throw new Error('AI is disabled in Markdown Web until office-agent-protocol is integrated.')
  }

  async aiStreamCancel(_requestId: string): Promise<void> {}

  onAiStream(_handler: (chunk: AiStreamChunk) => void): () => void {
    return () => {}
  }

  async webSearch(_query: string, _maxResults?: number): Promise<WebSearchResult> {
    throw new Error('Web search is disabled in Markdown Web until office-agent-protocol is integrated.')
  }

  destroy(): void {
    this.unsubscribeBridge?.()
    this.unsubscribeBridge = null
    this.languageListeners.clear()
    this.host.setDirty(false)
  }
}

export { MARKDOWN_ACCEPT }
