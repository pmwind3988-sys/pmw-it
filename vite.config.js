import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
  },
  build: {
    outDir: 'dist', // ← Vercel uses this
  },
  // NOTE: `assetsInclude: ['**/*.html']` used to be here. It matched index.html
  // itself, so Vite stopped treating it as the HTML entry and emitted it as a
  // static asset — `npm run build` produced a dist/index.html containing
  // `export default "/assets/index-….html"` and no bundle at all. Nothing in
  // src imports an .html file, so the option had nothing to do.
  optimizeDeps: {
    entries: ['src/main.jsx'],
  },
})
