import { fileURLToPath, URL } from 'node:url'

import { defineConfig, type Plugin } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueJsx from '@vitejs/plugin-vue-jsx'
import vueDevTools from 'vite-plugin-vue-devtools'

function certificateApiPlugin(): Plugin {
  return {
    name: 'certificate-api',
    // Only active during `vite dev`, skipped entirely during `vite build`
    apply: 'serve',
    async configureServer(server) {
      const { default: expressApp } = await import('./server/index')
      server.middlewares.use((req, res, next) => {
        if (req.url?.startsWith('/api')) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          expressApp(req as any, res as any, next)
          return
        }
        next()
      })
    },
  }
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    vueJsx(),
    vueDevTools(),
    certificateApiPlugin(),
  ],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
})
