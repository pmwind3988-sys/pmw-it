import { fileURLToPath } from 'node:url'
import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

const here = (path) => fileURLToPath(new URL(path, import.meta.url))

/**
 * Shared checklist links in `npm run dev`, as Vercel serves them in production:
 * `/c/<code>` is the public page (`checklist.html`, not the portal), and
 * `/api/c/<code>` runs the same handler as `api/c/[code].js`.
 *
 * With `CHECKLIST_FAKE=1` the handler talks to an in-memory SharePoint seeded
 * with three demo links (`server/devApi.js`), so the page can be driven
 * without the app secret.
 */
function checklistLinks(env) {
  let devApi = null

  return {
    name: 'checklist-links',
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        const path = (req.url ?? '').split('?')[0]

        const api = /^\/api\/c\/([^/]+)$/.exec(path)
        if (api) {
          try {
            const { handleLinkRequest, apiFromEnv } = await server.ssrLoadModule('/server/linkHandler.js')
            if (env.CHECKLIST_FAKE === '1' && !devApi) {
              const { createDevApi } = await server.ssrLoadModule('/server/devApi.js')
              devApi = await createDevApi()
            }
            await handleLinkRequest(req, res, decodeURIComponent(api[1]), {
              api: devApi ?? apiFromEnv(env),
            })
          } catch (error) {
            next(error)
          }
          return
        }

        if (/^\/c\/[^/]+\/?$/.test(path)) req.url = '/checklist.html'
        next()
      })
    },
  }
}

export default defineConfig(({ mode }) => {
  const env = { ...process.env, ...loadEnv(mode, process.cwd(), '') }

  return {
    plugins: [react(), checklistLinks(env)],
    server: {
      port: 5173,
    },
    build: {
      outDir: 'dist', // ← Vercel uses this
      rollupOptions: {
        // Two pages: the portal, and the page a shared checklist link opens.
        // The second carries the form kit and nothing else — no sign-in, no
        // portal routes — so a link cannot be walked back into the portal.
        input: {
          main: here('./index.html'),
          checklist: here('./checklist.html'),
        },
      },
    },
    // NOTE: `assetsInclude: ['**/*.html']` used to be here. It matched index.html
    // itself, so Vite stopped treating it as the HTML entry and emitted it as a
    // static asset — `npm run build` produced a dist/index.html containing
    // `export default "/assets/index-….html"` and no bundle at all. Nothing in
    // src imports an .html file, so the option had nothing to do.
    optimizeDeps: {
      entries: ['src/main.jsx', 'src/public/checklist.jsx'],
    },
    test: {
      environment: 'node',
      include: ['src/**/*.test.js', 'server/**/*.test.js'],
    },
  }
})
