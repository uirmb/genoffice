import {
  StandaloneOfficeHost,
  createEmbeddedOfficeRuntime,
  detectWebRuntimeMode,
} from '@genoffice/web-runtime'
import './product-policy.css'
import { createDocsWebDesktopController } from './desktop-api'

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:32px;color:#b42318"><h2>GenOffice Web failed to start</h2><pre style="white-space:pre-wrap">${message}</pre></div>`
}

function resolveHostOrigin(): string | null {
  const queryOrigin = new URL(window.location.href).searchParams.get('hostOrigin')
  if (queryOrigin) return queryOrigin
  const configured = import.meta.env.VITE_OFFICE_HOST_ORIGIN
  return typeof configured === 'string' && configured ? configured : null
}

async function bootstrapWeb(): Promise<void> {
  const mode = detectWebRuntimeMode()

  // Apply the default Web product policy before React renders so legacy desktop
  // controls never flash on screen. office:init capabilities may refine it later.
  document.documentElement.classList.add(
    'office-web',
    'office-page-crop-marks',
    'office-can-open',
    'office-can-save',
    'office-can-save-as',
  )
  localStorage.setItem('aidocs.showAi', '0')
  localStorage.setItem('aidocs.autoSave', '0')

  const embeddedRuntime =
    mode === 'embedded'
      ? createEmbeddedOfficeRuntime({
          hostOrigin:
            resolveHostOrigin() ??
            (() => {
              throw new Error(
                'Embedded Docs requires ?hostOrigin=https://host.example.com or VITE_OFFICE_HOST_ORIGIN.',
              )
            })(),
        })
      : null

  const standaloneHost = mode === 'standalone' ? new StandaloneOfficeHost() : null
  const host = embeddedRuntime?.host ?? standaloneHost
  if (!host) throw new Error('Unable to initialize the Docs web host runtime.')

  const controller = createDocsWebDesktopController(host, embeddedRuntime?.bridge)
  window.desktop = controller.desktopApi
  window.projectApi = controller.projectApi

  const cleanup = () => {
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
