import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // Relative assets work whether GitHub Pages publishes at the repository root
  // or under a project sub-path, preventing a blank page after deployment.
  base: './',
})
