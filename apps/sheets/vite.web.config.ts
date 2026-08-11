import react from '@vitejs/plugin-react'
import { defineConfig, type Plugin } from 'vite'

const XLSX_GATEWAY_NODE_MODULES = new Map<string, readonly string[]>([
  ['node:crypto', ['createHash']],
  ['node:fs/promises', ['open', 'readFile', 'rename', 'rm', 'writeFile']],
  ['node:path', ['dirname', 'join']],
])

function xlsxGatewayBrowserBoundary(): Plugin {
  const prefix = '\0genoffice-sheets-web-node-stub:'

  return {
    name: 'genoffice-sheets-web-xlsx-gateway-browser-boundary',
    enforce: 'pre',
    resolveId(source, importer) {
      if (!importer?.replaceAll('\\', '/').endsWith('/src/gateway/xlsx-gateway.ts')) return null
      if (!XLSX_GATEWAY_NODE_MODULES.has(source)) return null
      return `${prefix}${source}`
    },
    load(id) {
      if (!id.startsWith(prefix)) return null
      const source = id.slice(prefix.length)
      const exports = XLSX_GATEWAY_NODE_MODULES.get(source)
      if (!exports) return null
      const message = JSON.stringify(
        `${source} is Electron-only XLSX file I/O and must not execute in Sheets Web.`,
      )
      const declarations = exports
        .map(
          (name) =>
            `export function ${name}(..._args) { throw new Error(${message}) }`,
        )
        .join('\n')
      return declarations
    },
  }
}

export default defineConfig({
  root: 'src/web',
  plugins: [xlsxGatewayBrowserBoundary(), react()],
  server: {
    port: Number(process.env.SHEETS_WEB_PORT) || 5275,
    strictPort: true,
    proxy: {
      '/xlsx-engine': {
        target: process.env.XLSX_ENGINE_URL || 'http://127.0.0.1:7301',
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/xlsx-engine/, ''),
      },
    },
  },
  build: {
    outDir: '../../dist-web',
    emptyOutDir: true,
  },
})
