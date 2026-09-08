import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFile, readdir, writeFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';

const projectRoot = path.dirname(fileURLToPath(import.meta.url));
const webOutputDirectory = path.resolve(projectRoot, 'dist/web');
const offlineShellSource = path.resolve(projectRoot, 'apps/web/public/sw.js');
const kernelManifestPath = path.resolve(projectRoot, 'apps/web/public/kernel/kernel-manifest.json');
const kernelBuildManifest = JSON.parse(readFileSync(kernelManifestPath, 'utf8')) as {
  schema?: unknown;
  artifact?: unknown;
  expectedSha256?: unknown;
};
if (kernelBuildManifest.schema !== 'react-sheets.kernel-build.v1'
  || kernelBuildManifest.artifact !== 'kernel_host.wasm'
  || typeof kernelBuildManifest.expectedSha256 !== 'string'
  || !/^[a-f0-9]{64}$/.test(kernelBuildManifest.expectedSha256)) {
  throw new Error('Kernel build manifest is invalid or is not bound to the frontend build');
}
const kernelBuildId = kernelBuildManifest.expectedSha256;

async function listOutputAssets(directory: string, relative = ''): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true });
  const paths: string[] = [];
  for (const entry of entries) {
    const nextRelative = path.posix.join(relative, entry.name);
    const nextPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      paths.push(...await listOutputAssets(nextPath, nextRelative));
    } else if (/\.(?:css|js|mjs|wasm|woff2?|ttf|otf|svg)$/i.test(entry.name)) {
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
      const shellUrls = [
        '/',
        '/index.html',
        '/manifest.webmanifest',
        `/kernel/kernel-manifest.json?build=${kernelBuildId}`,
        `/kernel/kernel_host.wasm?sha256=${kernelBuildId}`,
        ...assets.filter((asset) => !asset.startsWith('/kernel/')).sort(),
      ];
      const workerPath = path.join(webOutputDirectory, 'sw.js');
      // Vite keeps an out-of-root output directory between some development
      // builds. Always render from the tracked shell template instead of a
      // previously generated manifest, otherwise a second package build has
      // no placeholder left to replace.
      const source = await readFile(offlineShellSource, 'utf8');
      const next = source.replace(
        /const BUILD_ID = '__REACT_SHEETS_BUILD_ID__';/,
        `const BUILD_ID = '${kernelBuildId}';`,
      ).replace(
        /const SHELL_URLS = \[[^;]+\];/,
        `const SHELL_URLS = ${JSON.stringify(shellUrls)};`,
      );
      if (next === source || next.includes('__REACT_SHEETS_BUILD_ID__')) throw new Error('Offline shell build placeholders were not found');
      await writeFile(workerPath, next, 'utf8');
    },
  };
}

export default defineConfig({
  root: path.resolve(projectRoot, 'apps/web'),
  plugins: [react(), offlineShellManifestPlugin()],
  define: {
    'import.meta.env.VITE_KERNEL_BUILD_ID': JSON.stringify(kernelBuildId),
  },
  resolve: {
    alias: [
      { find: 'react/jsx-dev-runtime', replacement: path.resolve(projectRoot, 'node_modules/react/jsx-dev-runtime.js') },
      { find: 'react/jsx-runtime', replacement: path.resolve(projectRoot, 'node_modules/react/jsx-runtime.js') },
      { find: 'react-dom', replacement: path.resolve(projectRoot, 'node_modules/react-dom') },
      { find: 'react', replacement: path.resolve(projectRoot, 'node_modules/react') },
    ],
    dedupe: ['react', 'react-dom'],
  },
  optimizeDeps: {
    include: ['react', 'react-dom', 'react-dom/client', 'react/jsx-runtime', 'react/jsx-dev-runtime'],
  },
  server: {
    proxy: {
      '/api': {
        changeOrigin: true,
        target: process.env.REACT_SHEETS_API_ORIGIN ?? 'http://127.0.0.1:8082',
      },
      '/ws': {
        changeOrigin: true,
        target: process.env.REACT_SHEETS_API_ORIGIN ?? 'http://127.0.0.1:8082',
        ws: true,
      },
    },
  },
  build: {
    outDir: webOutputDirectory,
    emptyOutDir: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          const normalized = id.replaceAll('\\', '/');
          if (normalized.includes('/packages/formula-engine/')) return 'formula-engine';
          if (normalized.includes('/packages/render-engine/')) return 'sheet-render';
          if (normalized.includes('/packages/core-model/') || normalized.includes('/packages/command-runtime/')) return 'sheet-model';
          if (normalized.includes('/packages/ui-system/')) return 'ui-system';
          return undefined;
        },
      },
    },
  },
});
