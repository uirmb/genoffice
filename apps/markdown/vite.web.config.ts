import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  root: 'src/web',
  plugins: [react()],
  server: {
    port: Number(process.env.MARKDOWN_WEB_PORT) || 5277,
    strictPort: true,
  },
  build: {
    outDir: '../../dist-web',
    emptyOutDir: true,
  },
})
