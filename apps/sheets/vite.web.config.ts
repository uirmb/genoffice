import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  root: 'src/web',
  plugins: [react()],
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
