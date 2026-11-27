import js from '@eslint/js';
import prettier from 'eslint-config-prettier';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';
import { defineConfig, globalIgnores } from 'eslint/config';
import globals from 'globals';
import tseslint from 'typescript-eslint';

// Browser / runtime globals that pure logic in src/core must never touch.
const NON_PURE_GLOBALS = [
  'window',
  'document',
  'navigator',
  'location',
  'history',
  'localStorage',
  'sessionStorage',
  'indexedDB',
  'DOMParser',
  'XMLSerializer',
  'Node',
  'Element',
  'HTMLElement',
  'Document',
  'XMLDocument',
  'Blob',
  'File',
  'FileReader',
  'fetch',
  'process',
  'Buffer',
].map((name) => ({
  name,
  message:
    'src/core is pure logic: no DOM, browser or Node APIs (feasibility report §5, execution context §6).',
}));

export default defineConfig([
  globalIgnores(['dist', 'coverage', 'node_modules', 'tests/fixtures/generated', 'a.*', 'a.*/**']),

  {
    files: ['**/*.{js,ts,tsx}'],
    extends: [js.configs.recommended, tseslint.configs.recommended],
    languageOptions: { ecmaVersion: 2023 },
  },

  // App code runs in the browser.
  {
    files: ['src/**/*.{ts,tsx}'],
    languageOptions: { globals: globals.browser },
  },
  {
    files: ['src/**/*.tsx'],
    extends: [reactHooks.configs.flat.recommended, reactRefresh.configs.vite],
  },

  // src/core: pure functions only (no React, no DOM, no I/O, no other app layers).
  {
    files: ['src/core/**/*.ts'],
    ignores: ['src/core/**/*.test.ts'],
    rules: {
      'no-restricted-globals': ['error', ...NON_PURE_GLOBALS],
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['react', 'react/*', 'react-dom', 'react-dom/*'],
              message: 'src/core must not depend on React.',
            },
            {
              group: ['jszip', 'idb', 'node:*'],
              message:
                'src/core must not do I/O: file and storage work belongs in src/docx or src/storage.',
            },
            {
              group: [
                '@/docx',
                '@/docx/*',
                '@/storage',
                '@/storage/*',
                '@/ui',
                '@/ui/*',
                '@/app',
                '@/app/*',
                '**/docx/**',
                '**/storage/**',
                '**/ui/**',
                '**/app/**',
              ],
              message: 'src/core is the bottom layer: it must not import other app layers.',
            },
          ],
        },
      ],
    },
  },

  // Tooling and tests run in Node.
  {
    files: ['tests/**/*.ts', '*.config.{js,ts}'],
    languageOptions: { globals: globals.node },
  },

  // Must stay last: turns off rules that clash with Prettier.
  prettier,
]);
