import {
  StandaloneOfficeHost,
  createEmbeddedOfficeRuntime,
  detectWebRuntimeMode,
} from '@genoffice/web-runtime'
import { installSheetsWebDesktopAdapters } from './desktop-adapters'
import { createSheetsWebDesktopController } from './desktop-api'
import { getXlsxEngineHealth } from './engine-client'
import './product-policy.css'

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:32px;color:#b42318"><h2>GenOffice Sheets Web failed to start</h2><pre style="white-space:pre-wrap">${message}</pre></div>`
}

function resolveHostOrigin(): string | null {
  const queryOrigin = new URL(window.location.href).searchParams.get('hostOrigin')
  if (queryOrigin) return queryOrigin
  const configured = import.meta.env.VITE_OFFICE_HOST_ORIGIN
  return typeof configured === 'string' && configured ? configured : null
}

async function bootstrapWeb(): Promise<void> {
  // Keep the Web product policy consistent with Docs/Slides: the platform owns
  // persistence and AI is not part of the embedded office surface.
  localStorage.setItem('ai-sheets-auto-save', '0')

  const health = await getXlsxEngineHealth()
  if (!health.ok) throw new Error('XLSX Engine Service reported an unhealthy state.')

  const mode = detectWebRuntimeMode()
  const embeddedRuntime =
    mode === 'embedded'
      ? createEmbeddedOfficeRuntime({
          hostOrigin:
            resolveHostOrigin() ??
            (() => {
              throw new Error(
                'Embedded Sheets requires ?hostOrigin=https://host.example.com or VITE_OFFICE_HOST_ORIGIN.',
              )
            })(),
        })
      : null
  const standaloneHost = mode === 'standalone' ? new StandaloneOfficeHost() : null
  const host = embeddedRuntime?.host ?? standaloneHost
  if (!host) throw new Error('Unable to initialize the Sheets web host runtime.')

  const controller = createSheetsWebDesktopController(host, embeddedRuntime?.bridge)
  installSheetsWebDesktopAdapters(controller.desktopApi, host)

  // Electron exposes this property as readonly through preload typings. Web
  // installs the same contract before importing the shared renderer.
  Object.defineProperty(window, 'desktopApi', {
    configurable: true,
    value: controller.desktopApi,
  })
  document.documentElement.classList.add('office-web')
  document.documentElement.dataset.xlsxSessionStore = health.sessionStore

  const cleanup = (): void => {
    controller.destroy()
    embeddedRuntime?.destroy()
    standaloneHost?.destroy()
  }
  window.addEventListener('pagehide', cleanup, { once: true })

  await import('../renderer/main')

  // Announce readiness only after the shared renderer has mounted and subscribed
  // to DesktopApi. This keeps office:init from racing the initial blank workbook.
  controller.notifyReady()
}

void bootstrapWeb().catch((error) => {
  console.error(error)
  renderBootstrapError(error)
})
