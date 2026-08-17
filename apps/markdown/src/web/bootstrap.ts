import {
  StandaloneOfficeHost,
  createEmbeddedOfficeRuntime,
  detectWebRuntimeMode,
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

async function bootstrapWeb(): Promise<void> {
  document.documentElement.classList.add('genoffice-markdown-web')

  const mode = detectWebRuntimeMode()
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

  const markdownApi = new MarkdownWebApi(host, embeddedRuntime?.bridge ?? null, mode)
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
