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
      '@react-sheets/exchange-xlsx': path.resolve(projectRoot, 'packages/exchange-xlsx/src'),
      '@react-sheets/spreadsheet-app': path.resolve(projectRoot, 'packages/spreadsheet-app/src'),
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
          if (normalized.includes('/packages/core-model/') || normalized.includes('/packages/command-runtime/')) return 'sheet-model';
          if (normalized.includes('/packages/ui-system/')) return 'ui-system';
          if (normalized.includes('/packages/spreadsheet-app/src/features/pivot/')) return 'feature-pivot';
          if (normalized.includes('/packages/spreadsheet-app/src/features/query/')) return 'feature-query';
          if (normalized.includes('/packages/spreadsheet-app/src/features/review/')) return 'feature-review';
          if (normalized.includes('/packages/spreadsheet-app/src/features/print/')) return 'feature-print';
          if (normalized.includes('/packages/spreadsheet-app/src/features/automation/')) return 'feature-automation';
          if (normalized.includes('/packages/spreadsheet-app/src/features/extended/')) return 'feature-what-if';
          if (normalized.includes('/packages/spreadsheet-app/src/features/drawing/') || normalized.includes('/packages/spreadsheet-app/src/features/chart/')) return 'feature-drawing';
          if (normalized.includes('/packages/spreadsheet-app/')) return 'workbook-runtime';
          return undefined;
        },
      },
    },
  },
});
