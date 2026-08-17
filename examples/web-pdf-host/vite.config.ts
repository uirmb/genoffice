import { defineConfig } from 'vite'

export default defineConfig({
  root: 'examples/web-pdf-host',
  server: {
    port: Number(process.env.PDF_WEB_HOST_PORT) || 8083,
    strictPort: true,
  },
  build: {
    outDir: '../../dist-web-pdf-host',
    emptyOutDir: true,
  },
})
