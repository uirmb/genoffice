import { defineConfig } from 'vite'

export default defineConfig({
  server: {
    port: Number(process.env.MARKDOWN_WEB_HOST_PORT) || 8084,
    strictPort: true,
  },
  build: {
    outDir: '../../dist-web-markdown-host',
    emptyOutDir: true,
  },
})
