import { createRoot } from 'react-dom/client'
import {
  StandaloneOfficeHost,
  createEmbeddedOfficeRuntime,
  detectWebRuntimeMode,
} from '@genoffice/web-runtime'
import { PdfWebApp } from './PdfWebApp'
import './styles.css'

function resolveHostOrigin(): string | null {
  const queryOrigin = new URL(window.location.href).searchParams.get('hostOrigin')
  if (queryOrigin) return queryOrigin
  const configured = import.meta.env.VITE_OFFICE_HOST_ORIGIN
  return typeof configured === 'string' && configured ? configured : null
}

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div class="pdf-web-bootstrap-error"><h2>GenOffice PDF failed to start</h2><pre>${message}</pre></div>`
}

function bootstrap(): void {
  const rootElement = document.getElementById('root')
  if (!rootElement) throw new Error('Missing #root element.')

  const runtimeMode = detectWebRuntimeMode()
  const embeddedRuntime =
    runtimeMode === 'embedded'
      ? createEmbeddedOfficeRuntime({
          hostOrigin:
            resolveHostOrigin() ??
            (() => {
              throw new Error(
                'Embedded PDF requires ?hostOrigin=https://host.example.com or VITE_OFFICE_HOST_ORIGIN.',
              )
            })(),
        })
      : null
  const standaloneHost = runtimeMode === 'standalone' ? new StandaloneOfficeHost() : null
  const host = embeddedRuntime?.host ?? standaloneHost
  if (!host) throw new Error('Unable to initialize the PDF web host runtime.')

  const root = createRoot(rootElement)
  root.render(
    <PdfWebApp host={host} bridge={embeddedRuntime?.bridge ?? null} runtimeMode={runtimeMode} />,
  )

  window.addEventListener(
    'pagehide',
    () => {
      root.unmount()
      embeddedRuntime?.destroy()
      standaloneHost?.destroy()
    },
    { once: true },
  )
}

try {
  bootstrap()
} catch (error) {
  console.error(error)
  renderBootstrapError(error)
}
