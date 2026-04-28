import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173, // ← local dev, ignored by Vercel
  },
  build: {
    outDir: 'dist', // ← Vercel uses this
  },
})