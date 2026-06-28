import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // ⚠️  Cambia "infratec-wr" con il nome ESATTO del tuo repository GitHub
  base: '/infratec-wr/',
})
