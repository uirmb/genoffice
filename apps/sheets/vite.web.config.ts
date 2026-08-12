import react from '@vitejs/plugin-react'
import { defineConfig, type Plugin } from 'vite'

const XLSX_GATEWAY_NODE_MODULES = new Map<string, readonly string[]>([
  ['node:crypto', ['createHash']],
  ['node:fs/promises', ['open', 'readFile', 'rename', 'rm', 'writeFile']],
  ['node:path', ['dirname', 'join']],
])

const xlsxEngineProxy = {
  '/xlsx-engine': {
    target: process.env.XLSX_ENGINE_URL || 'http://127.0.0.1:7301',
    changeOrigin: true,
    rewrite: (path: string) => path.replace(/^\/xlsx-engine/, ''),
  },
}

function xlsxGatewayBrowserBoundary(): Plugin {
  const prefix = '\0genoffice-sheets-web-node-stub:'
  const imageBase64Needle = "Uint8Array.from(Buffer.from(image.base64, 'base64'))"
  const imageBase64Replacement = '__genofficeDecodeBase64(image.base64)'
  const imageBase64Helper = `
function __genofficeDecodeBase64(input) {
  const normalized = input.replace(/\\s+/g, '').replace(/-/g, '+').replace(/_/g, '/')
  if (normalized.length % 4 === 1) throw new Error('Image payload is not valid base64.')
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=')
  let binary
  try {
    binary = atob(padded)
  } catch {
    throw new Error('Image payload is not valid base64.')
  }
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}
`

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
      return exports
        .map(
          (name) =>
            `export function ${name}(..._args) { throw new Error(${message}) }`,
        )
        .join('\n')
    },
    transform(code, id) {
      if (!id.replaceAll('\\', '/').endsWith('/src/gateway/xlsx-drawing-add.ts')) return null
      if (!code.includes(imageBase64Needle)) {
        throw new Error(
          'Sheets Web image base64 boundary no longer matches xlsx-drawing-add.ts; review the browser adaptation.',
        )
      }
      return {
        code: `${code.replace(imageBase64Needle, imageBase64Replacement)}\n${imageBase64Helper}`,
        map: null,
      }
    },
  }
}

export default defineConfig({
  root: 'src/web',
  plugins: [xlsxGatewayBrowserBoundary(), react()],
  server: {
    host: '0.0.0.0',
    port: Number(process.env.SHEETS_WEB_PORT) || 5275,
    strictPort: true,
    proxy: xlsxEngineProxy,
  },
  preview: {
    host: '0.0.0.0',
    port: Number(process.env.SHEETS_WEB_PORT) || 5275,
    strictPort: true,
    proxy: xlsxEngineProxy,
  },
  build: {
    outDir: '../../dist-web',
    emptyOutDir: true,
  },
})