import {
  StandaloneOfficeHost,
  createEmbeddedOfficeRuntime,
  detectWebRuntimeMode,
} from '@genoffice/web-runtime'
import type { DesktopFilesApi } from '../shared/ipc'
import { installSlidesWebHostPolicy } from './host-policy'
import './product-policy.css'
import { createSlidesWebController } from './slides-api'

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:32px;color:#b42318"><h2>GenOffice Slides Web failed to start</h2><pre style="white-space:pre-wrap">${message}</pre></div>`
}

function resolveHostOrigin(): string | null {
  const queryOrigin = new URL(window.location.href).searchParams.get('hostOrigin')
  if (queryOrigin) return queryOrigin
  const configured = import.meta.env.VITE_OFFICE_HOST_ORIGIN
  return typeof configured === 'string' && configured ? configured : null
}

function createDesktopFilesApi(): DesktopFilesApi {
  return {
    pickAttachments: async () => null,
    addAttachmentPaths: async () => ({
      accepted: [],
      rejected: ['Local path attachments are unavailable in Slides Web.'],
    }),
    addPastedImage: async () => ({
      accepted: [],
      rejected: ['AI attachments are disabled in Slides Web.'],
    }),
    readAttachment: async () => ({
      ok: false,
      error: 'AI attachments are disabled in Slides Web.',
    }),
    readAttachmentImage: async () => ({
      ok: false,
      error: 'AI attachments are disabled in Slides Web.',
    }),
    getPathForFile: (file) => `browser-file://${encodeURIComponent(file.name)}`,
  }
}

async function bootstrapWeb(): Promise<void> {
  const mode = detectWebRuntimeMode()

  // Match Docs Web product policy: platform owns AI/autosave decisions. Keep the
  // existing Slides renderer, but start it with those desktop-only surfaces off.
  localStorage.setItem('ai-slides-show-ai', '0')
  localStorage.setItem('ai-slides-auto-save', '0')

  const embeddedRuntime =
    mode === 'embedded'
      ? createEmbeddedOfficeRuntime({
          hostOrigin:
            resolveHostOrigin() ??
            (() => {
              throw new Error(
                'Embedded Slides requires ?hostOrigin=https://host.example.com or VITE_OFFICE_HOST_ORIGIN.',
              )
            })(),
        })
      : null

  const standaloneHost = mode === 'standalone' ? new StandaloneOfficeHost() : null
  const host = embeddedRuntime?.host ?? standaloneHost
  if (!host) throw new Error('Unable to initialize the Slides Web host runtime.')

  const policy = installSlidesWebHostPolicy(mode, embeddedRuntime?.bridge)
  const controller = createSlidesWebController(host, embeddedRuntime?.bridge)

  window.slidesApi = controller.slidesApi
  window.desktop = createDesktopFilesApi()

  const cleanup = () => {
    policy.destroy()
    controller.destroy()
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
