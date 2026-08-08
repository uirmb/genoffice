import type { ProjectApi } from '@genoffice/project-store'
import type {
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostApi,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import {
  OFFICE_PROTOCOL_VERSION,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'
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

type DocsLang = Awaited<ReturnType<DesktopApi['getLanguage']>>
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
  const lang = value.toLowerCase().split('-')[0]
  const supported: DocsLang[] = ['zh', 'en', 'ja', 'ko', 'fr', 'de', 'es', 'th', 'id', 'ru', 'ar']
  return supported.includes(lang as DocsLang) ? (lang as DocsLang) : 'en'
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

async function selectedToOfficeFile(host: OfficeHostApi, selected: SelectedOfficeFile): Promise<OfficeFile> {
  if (selected.transport === 'buffer' && selected.bytes) {
    return { ...selected, bytes: selected.bytes }
  }
  return host.readFile(selected.id)
}

function unavailable(message: string): Promise<never> {
  return Promise.reject(new Error(message))
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
  let pendingOpen: OpenFileResult | null = null
  let pendingStateRequestId: string | null = null
  let mode: 'view' | 'edit' = 'edit'
  let saving = false
  let currentLang: DocsLang = normalizeLang(document.documentElement.lang || navigator.language || 'en')
  let aiSettings = emptyAiSettings()

  const openHandlers = new Set<OpenHandler>()
  const renameHandlers = new Set<RenameHandler>()
  const languageHandlers = new Set<LanguageHandler>()
  const modeHandlers = new Set<(nextMode: 'view' | 'edit') => void>()
  const teardownHandlers = new Set<VoidHandler>()
  const menuHandlers = new Set<MenuHandler>()
  const closeCheckHandlers = new Set<VoidHandler>()
  const closeSaveHandlers = new Set<VoidHandler>()
  const aiStreamHandlers = new Set<(chunk: AiStreamChunk) => void>()
  const pendingSaveRequestIds = new Set<string>()

  const setMode = (nextMode: 'view' | 'edit'): void => {
    if (mode === nextMode) return
    mode = nextMode
    for (const handler of modeHandlers) handler(mode)
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

  const setCurrentFile = async (file: OfficeFile): Promise<OpenFileResult> => {
    const bytes = file.bytes.slice(0)
    current = {
      file: {
        id: file.id,
        name: file.name,
        mimeType: file.mimeType || DOCX_MIME,
        size: file.size ?? bytes.byteLength,
        version: file.version ?? null,
      },
      path: virtualPath(file),
      hash: await sha256(bytes),
      bytes,
    }
    host.setTitle(file.name)
    return contextToOpenResult()!
  }

  const openSelectedDocx = async (): Promise<OpenFileResult | null> => {
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
  ): Promise<{ ok: boolean; path?: string; error?: string; reason?: 'external-modified' }> => {
    const fallbackFile: OfficeFileDescriptor = current?.file ?? {
      id: `new:${Date.now()}`,
      name,
      mimeType: DOCX_MIME,
      size: data.byteLength,
      version: null,
    }
    const file: OfficeFileDescriptor = {
      ...fallbackFile,
      name,
      mimeType: DOCX_MIME,
      size: data.byteLength,
    }

    saving = true
    try {
      const result = await host.saveDocument({
        file,
        bytes: data,
        baseVersion: current?.file.version ?? null,
      })
      if (!result.ok) {
        reportHostSaveResult(false, result.error)
        return {
          ok: false,
          error: result.error,
          reason: result.code === 'VERSION_CONFLICT' ? 'external-modified' : undefined,
        }
      }

      const savedFile = result.file ?? file
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
    openDocx: openSelectedDocx,
    openDocxPath: async (path) => (current?.path === path ? contextToOpenResult() : null),
    consumePendingOpenDocx: async () => {
      const value = pendingOpen
      pendingOpen = null
      return value
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
    saveDocx: async (_path, data) => saveWithName(current?.file.name ?? 'Untitled.docx', data),
    writeRecoveryCopy: async () => ({ ok: true }),
    onTeardown: (handler) => {
      teardownHandlers.add(handler)
      return () => teardownHandlers.delete(handler)
    },
    saveDocxAs: async (defaultName, data) => saveWithName(defaultName, data),
    saveDocxNew: async (defaultName, data) => saveWithName(defaultName, data),
    getRecentFiles: async () => [],
    pickImage: async (): Promise<PickImageResult | null> => {
      const selected = await host.pickFile({ multiple: false, accept: [...IMAGE_MIMES] })
      if (!selected?.[0]) return null
      const file = await selectedToOfficeFile(host, selected[0])
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
    exportPdf: async () => ({ ok: false, error: 'PDF export is not available in the web runtime yet.' }),
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
    addAttachmentPaths: async () => ({ accepted: [], rejected: ['Local path attachments are unavailable on the web.'] }),
    addPastedImage: async () => ({ accepted: [], rejected: ['AI attachments are disabled on the web.'] }),
    readAttachment: async () => ({ ok: false, error: 'AI attachments are disabled on the web.' }),
    readAttachmentImage: async () => ({ ok: false, error: 'AI attachments are disabled on the web.' }),
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
        if (message.payload.locale) {
          currentLang = normalizeLang(message.payload.locale)
          for (const handler of languageHandlers) handler(currentLang)
        }
        void setCurrentFile(message.payload.file).then((result) => {
          if (openHandlers.size > 0) {
            for (const handler of openHandlers) handler(result)
          } else {
            pendingOpen = result
          }
        })
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
    notifyReady: () => {
      bridge?.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:ready',
        payload: { kind: 'docx' },
      })
    },
    destroy: () => {
      unsubscribeBridge?.()
      for (const handler of teardownHandlers) handler()
      openHandlers.clear()
      renameHandlers.clear()
      languageHandlers.clear()
      modeHandlers.clear()
      pendingSaveRequestIds.clear()
      teardownHandlers.clear()
      menuHandlers.clear()
      closeCheckHandlers.clear()
      closeSaveHandlers.clear()
      aiStreamHandlers.clear()
    },
  }
}
