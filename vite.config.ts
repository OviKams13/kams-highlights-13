import { fileURLToPath, URL } from 'node:url';
import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) },
  },
  test: {
    // Builds the synthetic .docx fixtures once, before any test file runs.
    globalSetup: ['./tests/global-setup.ts'],
    projects: [
      {
        // Default: plain Node (src/core, src/storage, tests/).
        extends: true,
        test: {
          name: 'node',
          environment: 'node',
          include: ['src/**/*.test.{ts,tsx}', 'tests/**/*.test.ts'],
          exclude: [...configDefaults.exclude, 'src/docx/**', 'src/ui/**', 'src/app/**'],
        },
      },
      {
        // Browser code runs in a simulated browser: React (src/ui, src/app) and the
        // .docx reader, which uses the browser DOMParser (src/docx).
        extends: true,
        test: {
          name: 'dom',
          environment: 'jsdom',
          include: [
            'src/docx/**/*.test.ts',
            'src/ui/**/*.test.{ts,tsx}',
            'src/app/**/*.test.{ts,tsx}',
          ],
        },
      },
    ],
  },
});
