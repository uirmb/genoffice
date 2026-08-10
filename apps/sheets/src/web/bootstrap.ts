import { getXlsxEngineHealth } from './engine-client'

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:32px;color:#b42318"><h2>GenOffice Sheets Web failed to start</h2><pre style="white-space:pre-wrap">${message}</pre></div>`
}

function renderFoundationReady(sessionStore: string): void {
  const root = document.getElementById('root')
  if (!root) return
  root.innerHTML = `
    <div style="font-family:system-ui,sans-serif;display:grid;place-items:center;min-height:100vh;background:#f5f7f8;color:#1f2328">
      <div style="padding:28px 32px;border:1px solid #dde2e7;border-radius:14px;background:#fff;box-shadow:0 18px 50px rgba(15,23,42,.08)">
        <h2 style="margin:0 0 8px">GenOffice Sheets Web</h2>
        <p style="margin:0;color:#667085">Web bootstrap and XLSX Engine Service are connected.</p>
        <p style="margin:12px 0 0;font-size:13px;color:#98a2b3">Session store: ${sessionStore}</p>
      </div>
    </div>
  `
}

async function bootstrapWeb(): Promise<void> {
  // The existing React + Univer renderer will be attached here after the
  // browser DesktopApi adapter is in place. Keep the first commit deliberately
  // isolated so Electron Sheets remains untouched while the service boundary
  // and deployment contract are established.
  const health = await getXlsxEngineHealth()
  if (!health.ok) throw new Error('XLSX Engine Service reported an unhealthy state.')
  renderFoundationReady(health.sessionStore)
}

void bootstrapWeb().catch((error) => {
  console.error(error)
  renderBootstrapError(error)
})
