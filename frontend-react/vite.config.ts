import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  root: path.resolve(projectRoot, 'apps/web'),
  plugins: [react()],
  resolve: {
    alias: {
      '@react-sheets/core-model': path.resolve(projectRoot, 'packages/core-model/src'),
      '@react-sheets/number-format': path.resolve(projectRoot, 'packages/number-format/src'),
      '@react-sheets/storage': path.resolve(projectRoot, 'packages/storage/src'),
      '@react-sheets/command-runtime': path.resolve(projectRoot, 'packages/command-runtime/src'),
      '@react-sheets/protocol': path.resolve(projectRoot, 'packages/protocol/src'),
      '@react-sheets/render-engine': path.resolve(projectRoot, 'packages/render-engine/src'),
      '@react-sheets/formula-engine': path.resolve(projectRoot, 'packages/formula-engine/src'),
      '@react-sheets/ui-system': path.resolve(projectRoot, 'packages/ui-system/src'),
      '@react-sheets/sheet-features': path.resolve(projectRoot, 'packages/sheet-features/src'),
      '@react-sheets/pro-features': path.resolve(projectRoot, 'packages/pro-features/src'),
    },
  },
  server: {
    proxy: {
      '/api': 'http://127.0.0.1:4181',
      '/ws': {
        target: 'ws://127.0.0.1:4181',
        ws: true,
      },
    },
  },
  build: {
    outDir: path.resolve(projectRoot, 'dist/web'),
    emptyOutDir: true,
  },
});
