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
      '/api': {
        target: 'http://127.0.0.1:4181',
        ws: true,
      },
      '/ws': {
        target: 'ws://127.0.0.1:4181',
        ws: true,
      },
    },
  },
  build: {
    outDir: path.resolve(projectRoot, 'dist/web'),
    emptyOutDir: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          const normalized = id.replaceAll('\\', '/');
          if (normalized.includes('/packages/formula-engine/') || normalized.includes('/packages/render-engine/')) return 'sheet-engine';
          if (normalized.includes('/packages/pro-features/')) return 'pro-features';
          if (normalized.includes('/packages/core-model/') || normalized.includes('/packages/command-runtime/')) return 'sheet-model';
          return undefined;
        },
      },
    },
  },
});
