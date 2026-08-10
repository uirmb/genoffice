import { getXlsxEngineHealth } from './engine-client'
import { createSheetsWebDesktopApi } from './desktop-api'

function renderBootstrapError(error: unknown): void {
  const root = document.getElementById('root')
  if (!root) return
  const message = error instanceof Error ? error.message : String(error)
  root.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:32px;color:#b42318"><h2>GenOffice Sheets Web failed to start</h2><pre style="white-space:pre-wrap">${message}</pre></div>`
}

async function bootstrapWeb(): Promise<void> {
  // Keep the Web product policy consistent with Docs/Slides: the platform owns
  // persistence and AI is not part of the embedded office surface.
  localStorage.setItem('ai-sheets-auto-save', '0')

  const health = await getXlsxEngineHealth()
  if (!health.ok) throw new Error('XLSX Engine Service reported an unhealthy state.')

  // The renderer stays shared with Electron. Browser-specific differences live
  // behind DesktopApi and, later, the office:* Host bridge.
  window.desktopApi = createSheetsWebDesktopApi()
  document.documentElement.classList.add('office-web')
  document.documentElement.dataset.xlsxSessionStore = health.sessionStore

  await import('../renderer/main')
}

void bootstrapWeb().catch((error) => {
  console.error(error)
  renderBootstrapError(error)
})
