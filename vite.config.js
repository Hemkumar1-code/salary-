import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 8000,
    allowedHosts: ['5ac5e4092b0e.ngrok-free.app']
  }
})
