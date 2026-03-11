import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import wasm from 'vite-plugin-wasm'
import topLevelAwait from 'vite-plugin-top-level-await'

export default defineConfig({
  plugins: [react(), wasm(), topLevelAwait()],
  optimizeDeps: {
    include: ['react-dom/client'],
  },
  test: {
    browser: {
      enabled: true,
      provider: 'playwright',
      instances: [{ browser: 'chromium' }],
    },
    testTimeout: 180_000, // GLA tests run ~2 million cycles in-browser
    onConsoleLog: () => (process.env.VERBOSE_TESTS ? undefined : false), // suppress WASM console.log noise
  },
})
