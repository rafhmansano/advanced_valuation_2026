import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/advanced_valuation_2026/',
  esbuild: {
    charset: 'utf8',
  },
})
