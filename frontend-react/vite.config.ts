import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFile, readdir, writeFile } from 'node:fs/promises';

const projectRoot = path.dirname(fileURLToPath(import.meta.url));
const webOutputDirectory = path.resolve(projectRoot, 'dist/web');

async function listOutputAssets(directory: string, relative = ''): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true });
  const paths: string[] = [];
  for (const entry of entries) {
    const nextRelative = path.posix.join(relative, entry.name);
    const nextPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      paths.push(...await listOutputAssets(nextPath, nextRelative));
    } else if (/\.(?:css|js|mjs|wasm|woff2?|ttf|otf)$/i.test(entry.name)) {
      paths.push(`/${nextRelative}`);
    }
  }
  return paths;
}

function offlineShellManifestPlugin() {
  return {
    name: 'offline-shell-manifest',
    async closeBundle() {
      const assets = await listOutputAssets(webOutputDirectory);
      const shellUrls = ['/', '/index.html', '/manifest.webmanifest', ...assets.sort()];
      const workerPath = path.join(webOutputDirectory, 'sw.js');
      const source = await readFile(workerPath, 'utf8');
      const next = source.replace(
        /const SHELL_URLS = \[[^;]+\];/,
        `const SHELL_URLS = ${JSON.stringify(shellUrls)};`,
      );
      if (next === source) throw new Error('Offline shell manifest placeholder was not found');
      await writeFile(workerPath, next, 'utf8');
    },
  };
}

export default defineConfig({
  root: path.resolve(projectRoot, 'apps/web'),
  plugins: [react(), offlineShellManifestPlugin()],
  build: {
    outDir: webOutputDirectory,
    emptyOutDir: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          const normalized = id.replaceAll('\\', '/');
          if (normalized.includes('/packages/formula-engine/') || normalized.includes('/packages/render-engine/')) return 'sheet-engine';
          if (normalized.includes('/packages/core-model/') || normalized.includes('/packages/command-runtime/')) return 'sheet-model';
          if (normalized.includes('/packages/ui-system/')) return 'ui-system';
          if (normalized.includes('/packages/spreadsheet-app/')) return 'workbook-runtime';
          return undefined;
        },
      },
    },
  },
});
