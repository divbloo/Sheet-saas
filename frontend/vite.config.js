import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  build: {
    // ExcelJS is loaded only when a user imports or exports Excel files.
    chunkSizeWarningLimit: 1000,
  },
})
