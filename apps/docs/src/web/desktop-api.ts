import type { ProjectApi } from '@genoffice/project-store'
import { normalizeLang as normalizeUiLang, type Lang } from '@genoffice/i18n'
import type {
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostApi,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { OFFICE_PROTOCOL_VERSION, type HostToEditorMessage } from '@genoffice/office-protocol'
import type {
  AiSettings,
  AiStreamChunk,
  DesktopApi,
  MenuCommand,
  OpenFileResult,
  PickImageResult,
} from '../shared/ipc'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const IMAGE_MIMES = ['image/png', 'image/jpeg', 'image/gif'] as const

type DocsLang = Lang
type OpenHandler = Parameters<DesktopApi['onOpenDocx']>[0]
type RenameHandler = Parameters<DesktopApi['onRenamedDocx']>[0]
type LanguageHandler = Parameters<DesktopApi['onLanguageChanged']>[0]
type MenuHandler = Parameters<DesktopApi['onMenuCommand']>[0]
type VoidHandler = () => void

interface WebDocumentContext {
  file: OfficeFileDescriptor
  path: string
  hash: string
  bytes: ArrayBuffer
}

export interface DocsWebDesktopController {
  desktopApi: DesktopApi
  projectApi: ProjectApi
  notifyReady(): void
  destroy(): void
}

function emptyAiSettings(): AiSettings {
  const empty = { apiKey: '', model: '' }
  return {
    provider: 'custom',
    providers: {
      genspark: { ...empty },
      anthropic: { ...empty },
      gemini: { ...empty },
      deepseek: { ...empty },
      openai: { ...empty },
      custom: { ...empty, baseUrl: '' },
    },
  }
}

function normalizeLang(value: string): DocsLang {
  return normalizeUiLang(value)
}

function virtualPath(file: OfficeFileDescriptor): string {
  return `web-office://files/${encodeURIComponent(file.id)}/${encodeURIComponent(file.name)}`
}

async function sha256(bytes: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes)
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('')
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)
  let binary = ''
  const chunkSize = 0x8000
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize))
  }
  return btoa(binary)
}

async function selectedToOfficeFile(
  host: OfficeHostApi,
  selected: SelectedOfficeFile,
): Promise<OfficeFile> {
  if (selected.transport === 'buffer' && selected.bytes) {
    return { ...selected, transport: 'buffer', bytes: selected.bytes }
  }
  return host.readFile(selected.id)
}

function unavailable(message: string): Promise<never> {
  return Promise.reject(new Error(message))
}

function hostPickerFailure(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

function officeDescriptor(file: OfficeFile, fallbackMime: string): OfficeFileDescriptor {
  const { bytes: _bytes, ...descriptor } = file
  return {
    ...descriptor,
    mimeType: file.mimeType || fallbackMime,
    size: file.size ?? file.bytes.byteLength,
    version: file.version ?? null,
    transport: 'buffer',
  }
}

function createProjectApi(): ProjectApi {
  return {
    resolveChat: async ({ tempChatId }) => ({
      projectId: 'web-office',
      chatId: tempChatId ?? `web-${Date.now()}`,
    }),
    appendChat: async () => {},
    loadChat: async () => [],
    rebindChat: async ({ projectId, newChatId, tempChatId }) => ({
      projectId,
      chatId: newChatId ?? tempChatId,
    }),
    listProjects: async () => [],
    createProject: () => unavailable('Projects are not available in the web editor runtime'),
    renameProject: async () => {},
    deleteProject: async () => {},
    moveFile: async () => {},
    getTimeline: async () => [],
  }
}

export function createDocsWebDesktopController(
  host: OfficeHostApi,
  bridge?: EditorIframeBridge,
): DocsWebDesktopController {
  let current: WebDocumentContext | null = null
  const pendingDocumentSelections = new Map<string, OfficeFile>()
  let pendingOpen: OpenFileResult | null = null
  let initialOpenResolve: ((result: OpenFileResult | null) => void) | null = null
  let readyNotified = false
  let pendingStateRequestId: string | null = null
  let pendingHostCloseRequest: { requestId: string; reason: 'window-close' } | null = null
  let mode: 'view' | 'edit' = 'edit'
  let saving = false
  let currentLang: DocsLang = normalizeLang(
    document.documentElement.lang || navigator.language || 'en',
  )
  let aiSettings = emptyAiSettings()

  const openHandlers = new Set<OpenHandler>()
  const renameHandlers = new Set<RenameHandler>()
  const languageHandlers = new Set<LanguageHandler>()
  const modeHandlers = new Set<(nextMode: 'view' | 'edit') => void>()
  const teardownHandlers = new Set<VoidHandler>()
  const menuHandlers = new Set<MenuHandler>()
  const closeCheckHandlers = new Set<VoidHandler>()
  const closeSaveHandlers = new Set<VoidHandler>()
  const hostCloseRequestHandlers = new Set<VoidHandler>()
  const aiStreamHandlers = new Set<(chunk: AiStreamChunk) => void>()
  const pendingSaveRequestIds = new Set<string>()

  const setMode = (nextMode: 'view' | 'edit'): void => {
    if (mode === nextMode) return
    mode = nextMode
    for (const handler of modeHandlers) handler(mode)
  }

  const setLanguage = (locale: string): void => {
    const next = normalizeLang(locale)
    if (currentLang === next) return
    currentLang = next
    for (const handler of languageHandlers) handler(currentLang)
  }

  const notifyReady = (): void => {
    if (readyNotified) return
    readyNotified = true
    bridge?.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:ready',
      payload: { kind: 'docx' },
    })
  }

  const reportHostSaveResult = (ok: boolean, error?: string): void => {
    if (!bridge || pendingSaveRequestIds.size === 0) return
    for (const requestId of pendingSaveRequestIds) {
      bridge.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-result',
        requestId,
        payload: { ok, error },
      })
    }
    pendingSaveRequestIds.clear()
  }

  const contextToOpenResult = (): OpenFileResult | null =>
    current
      ? {
          path: current.path,
          name: current.file.name,
          data: current.bytes.slice(0),
          hash: current.hash,
        }
      : null

  const fileToOpenResult = async (
    file: OfficeFile,
    selectionId?: string,
  ): Promise<OpenFileResult> => {
    const bytes = file.bytes.slice(0)
    return {
      path: virtualPath(file),
      name: file.name,
      data: bytes,
      hash: await sha256(bytes),
      ...(selectionId ? { selectionId } : {}),
    }
  }

  const setCurrentFile = async (file: OfficeFile): Promise<OpenFileResult> => {
    const bytes = file.bytes.slice(0)
    current = {
      file: officeDescriptor(file, DOCX_MIME),
      path: virtualPath(file),
      hash: await sha256(bytes),
      bytes,
    }
    host.setTitle(file.name)
    return contextToOpenResult()!
  }

  const openSelectedDocx = async (): Promise<OpenFileResult | null> => {
    if (host.pickDocument && host.confirmDocumentOpened && host.releasePickedDocument) {
      const picked = await host.pickDocument({ accept: [DOCX_MIME, '.docx'] })
      if (picked.status === 'cancelled') return null
      if (picked.status === 'failed') throw hostPickerFailure(picked.code, picked.error)
      pendingDocumentSelections.set(picked.selectionId, picked.file)
      return fileToOpenResult(picked.file, picked.selectionId)
    }

    // Compatibility only for custom/older hosts that have not implemented the stable API.
    const selected = await host.pickFile({
      multiple: false,
      accept: [DOCX_MIME, '.docx'],
      mode: 'file',
    })
    if (!selected?.[0]) return null
    return setCurrentFile(await selectedToOfficeFile(host, selected[0]))
  }

  const saveWithName = async (
    name: string,
    data: ArrayBuffer,
    saveMode: 'save' | 'saveAs' = 'save',
  ): Promise<{ ok: boolean; path?: string; error?: string; reason?: 'external-modified' }> => {
    const fallbackFile: OfficeFileDescriptor = current?.file ?? {
      id: `new:${Date.now()}`,
      name,
      mimeType: DOCX_MIME,
      size: data.byteLength,
      version: null,
      transport: 'buffer',
    }
    const file: OfficeFileDescriptor = {
      ...fallbackFile,
      name,
      mimeType: DOCX_MIME,
      size: data.byteLength,
    }

    const newDocument = current === null
    saving = true
    try {
      const result = await host.saveDocument({
        file,
        bytes: data,
        baseVersion: current?.file.version ?? null,
        mode: saveMode,
        newDocument,
      })
      if (!result.ok) {
        reportHostSaveResult(false, result.error)
        return {
          ok: false,
          error: result.error,
          reason: result.code === 'VERSION_CONFLICT' ? 'external-modified' : undefined,
        }
      }

      if (!result.file) {
        const error = 'Host reported a successful save without the latest file descriptor.'
        reportHostSaveResult(false, error)
        return { ok: false, error }
      }

      const savedFile = result.file
      const oldPath = current?.path
      current = {
        file: savedFile,
        path: virtualPath(savedFile),
        hash: await sha256(data),
        bytes: data.slice(0),
      }
      host.setTitle(savedFile.name)
      if (oldPath && oldPath !== current.path) {
        for (const handler of renameHandlers) handler({ oldPath, newPath: current.path })
      }
      reportHostSaveResult(true)
      return { ok: true, path: current.path }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      reportHostSaveResult(false, message)
      return { ok: false, error: message }
    } finally {
      saving = false
    }
  }

  const emitAiUnavailable = (requestId: string): void => {
    queueMicrotask(() => {
      for (const handler of aiStreamHandlers) {
        handler({
          requestId,
          type: 'error',
          error: 'AI features are disabled in the standalone web editor runtime.',
        })
      }
    })
  }

  const desktopApi: DesktopApi = {
    getLanguage: async () => currentLang,
    onLanguageChanged: (handler) => {
      languageHandlers.add(handler)
      return () => languageHandlers.delete(handler)
    },
    getHostEditorMode: async () => mode,
    onHostEditorModeChanged: (handler) => {
      modeHandlers.add(handler)
      return () => modeHandlers.delete(handler)
    },
    reportDirtyChange: (dirty) => host.setDirty(dirty),
    canAutoPersistPathlessDocument: () => false,
    openDocx: openSelectedDocx,
    confirmOpenDocx: async (selectionId) => {
      const candidate = pendingDocumentSelections.get(selectionId)
      if (!candidate || !host.confirmDocumentOpened) {
        return { ok: false, error: 'The pending document selection is no longer available.' }
      }
      const result = await host.confirmDocumentOpened(selectionId)
      if (!result.ok) return { ok: false, error: result.error }

      const bytes = candidate.bytes.slice(0)
      const candidateDescriptor = officeDescriptor(candidate, DOCX_MIME)
      const bound: OfficeFileDescriptor = result.file
        ? { ...candidateDescriptor, ...result.file, transport: 'buffer' }
        : candidateDescriptor
      current = {
        file: bound,
        path: virtualPath(bound),
        hash: await sha256(bytes),
        bytes,
      }
      pendingDocumentSelections.delete(selectionId)
      host.setTitle(current.file.name)
      return { ok: true }
    },
    releaseOpenDocx: async (selectionId) => {
      pendingDocumentSelections.delete(selectionId)
      await host.releasePickedDocument?.(selectionId)
    },
    openDocxPath: async (path) => (current?.path === path ? contextToOpenResult() : null),
    consumePendingOpenDocx: async () => {
      const value = pendingOpen
      if (value) {
        pendingOpen = null
        return value
      }
      if (!bridge) return null

      // The App registers onOpenDocx before calling this method. Announce readiness
      // only now, then keep the initial boot open pending until the host replies
      // with office:init. This prevents a blank document from racing the real file.
      notifyReady()
      return new Promise<OpenFileResult | null>((resolve) => {
        initialOpenResolve = resolve
      })
    },
    consumeNewBlankDoc: async () => false,
    onOpenDocx: (handler) => {
      openHandlers.add(handler)
      return () => openHandlers.delete(handler)
    },
    onRenamedDocx: (handler) => {
      renameHandlers.add(handler)
      return () => renameHandlers.delete(handler)
    },
    saveDocx: async (_path, data) =>
      saveWithName(current?.file.name ?? 'Untitled.docx', data, 'save'),
    writeRecoveryCopy: async () => ({ ok: true }),
    onTeardown: (handler) => {
      teardownHandlers.add(handler)
      return () => teardownHandlers.delete(handler)
    },
    saveDocxAs: async (defaultName, data) => saveWithName(defaultName, data, 'saveAs'),
    saveDocxNew: async (defaultName, data) => saveWithName(defaultName, data, 'save'),
    saveHistoryVersion: async (_defaultName, data) => {
      if (!current) return { ok: false, error: 'Save the new document before creating history.' }
      if (!host.saveHistoryVersion) {
        return { ok: false, error: 'History versions are not supported by this host.' }
      }
      const result = await host.saveHistoryVersion({
        file: current.file,
        bytes: data,
        baseVersion: current.file.version ?? null,
      })
      if (result.ok && result.file) {
        current.file = { ...result.file, mimeType: result.file.mimeType || DOCX_MIME }
        current.path = virtualPath(current.file)
        host.setTitle(current.file.name)
      }
      return { ok: result.ok, error: result.error }
    },
    exportDocx: async (defaultName, data) => {
      const download = host.downloadDocument ?? host.exportDocument
      if (!download) {
        return { ok: false, error: 'DOCX export is not supported by this host.' }
      }
      const descriptor: OfficeFileDescriptor = current?.file ?? {
        id: `export:${Date.now()}`,
        name: defaultName,
        mimeType: DOCX_MIME,
        size: data.byteLength,
        version: null,
        transport: 'buffer',
      }
      const result = await download.call(host, {
        format: 'docx',
        file: { ...descriptor, name: defaultName, size: data.byteLength },
        bytes: data,
      })
      return { ok: result.ok, error: result.error }
    },
    requestHostClose: async () => {
      if (pendingHostCloseRequest) {
        const request = pendingHostCloseRequest
        pendingHostCloseRequest = null
        if (host.approveClose) {
          await host.approveClose(request.requestId)
          return
        }
        if (bridge) {
          bridge.send({
            protocol: OFFICE_PROTOCOL_VERSION,
            type: 'office:close-request',
            requestId: request.requestId,
            payload: { reason: request.reason },
          })
          return
        }
      }
      if (host.approveClose) await host.approveClose()
      else await host.requestClose?.()
    },
    onHostCloseRequest: (handler) => {
      hostCloseRequestHandlers.add(handler)
      if (pendingHostCloseRequest) queueMicrotask(handler)
      return () => hostCloseRequestHandlers.delete(handler)
    },
    cancelHostCloseRequest: () => {
      if (!pendingHostCloseRequest) return
      const request = pendingHostCloseRequest
      pendingHostCloseRequest = null
      if (host.cancelClose) {
        void host.cancelClose(request.requestId)
        return
      }
      bridge?.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:close-cancelled',
        requestId: request.requestId,
        payload: { reason: 'user-cancelled' },
      })
    },
    getRecentFiles: async () => [],
    pickImage: async (): Promise<PickImageResult | null> => {
      let file: OfficeFile
      if (host.pickAssets) {
        const picked = await host.pickAssets({ multiple: false, accept: [...IMAGE_MIMES] })
        if (picked.status === 'cancelled') return null
        if (picked.status === 'failed') throw hostPickerFailure(picked.code, picked.error)
        if (!picked.files[0]) throw new Error('Host returned an empty selected asset result.')
        file = picked.files[0]
      } else {
        const selected = await host.pickFile({ multiple: false, accept: [...IMAGE_MIMES] })
        if (!selected?.[0]) return null
        file = await selectedToOfficeFile(host, selected[0])
      }
      if (!IMAGE_MIMES.includes(file.mimeType as (typeof IMAGE_MIMES)[number])) {
        throw new Error(`Unsupported image type: ${file.mimeType}`)
      }
      return {
        base64: arrayBufferToBase64(file.bytes),
        mime: file.mimeType as PickImageResult['mime'],
        name: file.name,
      }
    },
    getAiSettings: async () => aiSettings,
    setAiSettings: async (settings) => {
      aiSettings = settings
    },
    print: async () => window.print(),
    exportPdf: async () => ({
      ok: false,
      error: 'PDF export is not available in the web runtime yet.',
    }),
    printPdfBuffer: async () => ({
      ok: false,
      error: 'PDF buffer export is not available in the web runtime yet.',
    }),
    saveMergedPdf: async () => ({
      ok: false,
      error: 'PDF merge is not available in the web runtime yet.',
    }),
    aiChat: async () => ({ ok: false, error: 'AI features are disabled in the web runtime.' }),
    aiStream: async (request) => emitAiUnavailable(request.requestId),
    aiStreamCancel: async () => {},
    aiGskStatus: async () => ({ loggedIn: false }),
    aiGskLogin: async () => {},
    webSearch: async () => ({ results: [], method: 'error', error: 'Web search is disabled.' }),
    imageSearch: async () => ({ images: [], method: 'error', error: 'Image search is disabled.' }),
    fetchImage: async () => null,
    pickAttachments: async () => null,
    addAttachmentPaths: async () => ({
      accepted: [],
      rejected: ['Local path attachments are unavailable on the web.'],
    }),
    addPastedImage: async () => ({
      accepted: [],
      rejected: ['AI attachments are disabled on the web.'],
    }),
    readAttachment: async () => ({ ok: false, error: 'AI attachments are disabled on the web.' }),
    readAttachmentImage: async () => ({
      ok: false,
      error: 'AI attachments are disabled on the web.',
    }),
    getPathForFile: (file) => `browser-file://${encodeURIComponent(file.name)}`,
    openNewTab: async () => {},
    listDocsTabs: async () => [],
    focusDocsTab: async () => {},
    onAiStream: (handler) => {
      aiStreamHandlers.add(handler)
      return () => aiStreamHandlers.delete(handler)
    },
    onMenuCommand: (handler) => {
      menuHandlers.add(handler)
      return () => menuHandlers.delete(handler)
    },
    onCloseCheck: (handler) => {
      closeCheckHandlers.add(handler)
      return () => closeCheckHandlers.delete(handler)
    },
    reportCloseCheck: (state) => {
      if (!bridge || !pendingStateRequestId) return
      bridge.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:state-result',
        requestId: pendingStateRequestId,
        payload: {
          ready: true,
          dirty: state.dirty,
          saving,
          mode,
          title: current?.file.name,
        },
      })
      pendingStateRequestId = null
    },
    onCloseSaveRequest: (handler) => {
      closeSaveHandlers.add(handler)
      return () => closeSaveHandlers.delete(handler)
    },
    reportCloseSaveResult: () => {},
  }

  const unsubscribeBridge = bridge?.subscribe((message: HostToEditorMessage) => {
    switch (message.type) {
      case 'office:init': {
        if (message.payload.kind !== 'docx') return
        setMode(message.payload.mode)
        if (message.payload.locale) setLanguage(message.payload.locale)
        void setCurrentFile(message.payload.file).then((result) => {
          if (initialOpenResolve) {
            const resolve = initialOpenResolve
            initialOpenResolve = null
            resolve(result)
          } else if (openHandlers.size > 0) {
            for (const handler of openHandlers) handler(result)
          } else {
            pendingOpen = result
          }
        })
        break
      }
      case 'office:new': {
        if (message.payload.kind !== 'docx') return
        setMode(message.payload.mode)
        if (message.payload.locale) setLanguage(message.payload.locale)
        current = null
        pendingOpen = null
        if (initialOpenResolve) {
          const resolve = initialOpenResolve
          initialOpenResolve = null
          resolve(null)
        } else {
          for (const handler of menuHandlers) handler('new' satisfies MenuCommand)
        }
        break
      }
      case 'office:set-locale':
        setLanguage(message.payload.locale)
        break
      case 'office:request-close': {
        if (pendingHostCloseRequest) break
        pendingHostCloseRequest = {
          requestId: message.requestId,
          reason: message.payload.reason,
        }
        for (const handler of hostCloseRequestHandlers) handler()
        break
      }
      case 'office:save': {
        const shouldTriggerSave = pendingSaveRequestIds.size === 0
        pendingSaveRequestIds.add(message.requestId)
        if (!shouldTriggerSave) break
        if (menuHandlers.size === 0) {
          reportHostSaveResult(false, 'Editor save handler is not ready.')
          break
        }
        for (const handler of menuHandlers) handler('save' satisfies MenuCommand)
        break
      }
      case 'office:query-state':
        pendingStateRequestId = message.requestId
        for (const handler of closeCheckHandlers) handler()
        break
      case 'office:set-mode':
        setMode(message.payload.mode)
        break
      case 'office:error':
        console.error(`[office host:${message.payload.code}] ${message.payload.message}`)
        break
      default:
        break
    }
  })

  return {
    desktopApi,
    projectApi: createProjectApi(),
    notifyReady,
    destroy: () => {
      unsubscribeBridge?.()
      for (const handler of teardownHandlers) handler()
      openHandlers.clear()
      renameHandlers.clear()
      languageHandlers.clear()
      modeHandlers.clear()
      pendingSaveRequestIds.clear()
      initialOpenResolve?.(null)
      initialOpenResolve = null
      teardownHandlers.clear()
      menuHandlers.clear()
      closeCheckHandlers.clear()
      closeSaveHandlers.clear()
      hostCloseRequestHandlers.clear()
      pendingHostCloseRequest = null
      for (const selectionId of pendingDocumentSelections.keys()) {
        void host.releasePickedDocument?.(selectionId)
      }
      pendingDocumentSelections.clear()
      aiStreamHandlers.clear()
    },
  }
}
