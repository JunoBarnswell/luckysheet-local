import { app, BrowserWindow, net, protocol } from 'electron';
import path from 'node:path';
import { existsSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'app',
    privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true },
  },
]);

const electronDirectory = path.dirname(fileURLToPath(import.meta.url));
const webDirectory = path.resolve(electronDirectory, '..', 'dist', 'web');

function resolveWebAsset(requestUrl) {
  const requestPath = decodeURIComponent(new URL(requestUrl).pathname);
  const relativePath = requestPath === '/' ? 'index.html' : requestPath.replace(/^\/+/, '');
  const candidate = path.resolve(webDirectory, relativePath);
  if (candidate.startsWith(`${webDirectory}${path.sep}`) && existsSync(candidate)) return candidate;
  return path.join(webDirectory, 'index.html');
}

async function createWindow() {
  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 680,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  window.once('ready-to-show', () => window.show());
  await window.loadURL('app://-/');
}

app.whenReady().then(async () => {
  protocol.handle('app', (request) => net.fetch(pathToFileURL(resolveWebAsset(request.url)).toString()));
  await createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) void createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
