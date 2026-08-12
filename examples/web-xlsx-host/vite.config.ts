import { defineConfig } from 'vite'

export default defineConfig({
  root: __dirname,
  server: {
    host: '0.0.0.0',
    port: 8082,
    strictPort: true,
  },
})
