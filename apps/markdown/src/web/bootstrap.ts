import { htmlLang, normalizeLang, type Lang } from '@genoffice/i18n'
import type { OfficeHostApi } from '@genoffice/office-host-api'
import {
  StandaloneOfficeHost,
  createEmbeddedOfficeRuntime,
  detectWebRuntimeMode,
  type EditorIframeBridge,
  type WebRuntimeMode,
} from '@genoffice/web-runtime'
import { MarkdownWebApi } from './markdown-api'
import './product-policy.css'

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:32px;color:#b42318"><h2>GenOffice Markdown Web failed to start</h2><pre style="white-space:pre-wrap">${message}</pre></div>`
}

function resolveHostOrigin(): string | null {
  const queryOrigin = new URL(window.location.href).searchParams.get('hostOrigin')
  if (queryOrigin) return queryOrigin
  const configured = import.meta.env.VITE_OFFICE_HOST_ORIGIN
  return typeof configured === 'string' && configured ? configured : null
}

function resolveRequestedLocale(): string | null {
  const locale = new URL(window.location.href).searchParams.get('locale')?.trim()
  return locale || null
}

/**
 * Markdown used to read <html lang="zh-CN"> before UC's initial message arrived,
 * so every newly-created iframe started in Chinese even when the UC plugin URL
 * already carried locale=en-US. Keep the current locale as runtime state and
 * accept all three Host sources: URL bootstrap, office:init/new, and live
 * office:set-locale updates.
 */
class LocalizedMarkdownWebApi extends MarkdownWebApi {
  private currentLang: Lang
  private readonly localeListeners = new Set<(lang: Lang) => void>()
  private initialLanguageResolved = false
  private resolveInitialLanguage!: () => void
  private readonly initialLanguage = new Promise<void>((resolve) => {
    this.resolveInitialLanguage = resolve
  })
  private unsubscribeLocaleBridge: (() => void) | null = null

  constructor(
    host: OfficeHostApi,
    bridge: EditorIframeBridge | null,
    runtimeMode: WebRuntimeMode,
    initialLocale: string | null,
  ) {
    super(host, bridge, runtimeMode)

    this.currentLang = normalizeLang(
      initialLocale || document.documentElement.lang || navigator.language || 'en',
    )
    this.applyLanguage(this.currentLang)

    if (runtimeMode === 'standalone' || initialLocale) this.resolveInitial()

    if (bridge) {
      this.unsubscribeLocaleBridge = bridge.subscribe((message) => {
        switch (message.type) {
          case 'office:init':
          case 'office:new':
            if (message.payload.kind !== 'markdown') return
            if (message.payload.locale) this.applyLanguage(message.payload.locale)
            this.resolveInitial()
            return
          case 'office:set-locale':
            this.applyLanguage(message.payload.locale)
            this.resolveInitial()
            return
          default:
            return
        }
      })
    } else {
      this.resolveInitial()
    }
  }

  private resolveInitial(): void {
    if (this.initialLanguageResolved) return
    this.initialLanguageResolved = true
    this.resolveInitialLanguage()
  }

  private applyLanguage(locale: string | Lang): void {
    const next = normalizeLang(locale)
    this.currentLang = next
    document.documentElement.lang = htmlLang(next)
    for (const listener of this.localeListeners) listener(next)
  }

  override async getLanguage(): Promise<Lang> {
    await this.initialLanguage
    return this.currentLang
  }

  override onLanguageChanged(handler: (lang: Lang) => void): () => void {
    this.localeListeners.add(handler)
    queueMicrotask(() => {
      if (this.localeListeners.has(handler)) handler(this.currentLang)
    })
    return () => this.localeListeners.delete(handler)
  }

  override destroy(): void {
    this.resolveInitial()
    this.unsubscribeLocaleBridge?.()
    this.unsubscribeLocaleBridge = null
    this.localeListeners.clear()
    super.destroy()
  }
}

async function bootstrapWeb(): Promise<void> {
  document.documentElement.classList.add('genoffice-markdown-web')

  const mode = detectWebRuntimeMode()
  const requestedLocale = resolveRequestedLocale()
  const embeddedRuntime =
    mode === 'embedded'
      ? createEmbeddedOfficeRuntime({
          hostOrigin:
            resolveHostOrigin() ??
            (() => {
              throw new Error(
                'Embedded Markdown requires ?hostOrigin=https://host.example.com or VITE_OFFICE_HOST_ORIGIN.',
              )
            })(),
        })
      : null

  const standaloneHost = mode === 'standalone' ? new StandaloneOfficeHost() : null
  const host = embeddedRuntime?.host ?? standaloneHost
  if (!host) throw new Error('Unable to initialize the Markdown web host runtime.')

  const markdownApi = new LocalizedMarkdownWebApi(
    host,
    embeddedRuntime?.bridge ?? null,
    mode,
    requestedLocale,
  )
  window.markdownApi = markdownApi

  const cleanup = () => {
    markdownApi.destroy()
    embeddedRuntime?.destroy()
    standaloneHost?.destroy()
  }
  window.addEventListener('pagehide', cleanup, { once: true })

  await import('../renderer/main')
}

void bootstrapWeb().catch((error) => {
  console.error(error)
  renderBootstrapError(error)
})
