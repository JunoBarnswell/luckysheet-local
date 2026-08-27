import { app, BrowserWindow, net, protocol } from 'electron';
import path from 'node:path';
import { existsSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { collaborationArgument, resolveBackendOrigin, resolveCollaborationUrl } from './desktop-config.mjs';

protocol.registerSchemesAsPrivileged([{
  scheme: 'app',
  privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true },
}]);

const electronDirectory = path.dirname(fileURLToPath(import.meta.url));
const webDirectory = path.resolve(electronDirectory, '..', 'dist', 'web');

function resolveWebAsset(requestUrl) {
  const requestPath = decodeURIComponent(new URL(requestUrl).pathname);
  const relativePath = requestPath === '/' ? 'index.html' : requestPath.replace(/^\/+/, '');
  const candidate = path.resolve(webDirectory, relativePath);
  if (!candidate.startsWith(`${webDirectory}${path.sep}`) || !existsSync(candidate)) return undefined;
  return candidate;
}

async function handleAppRequest(request, backendOrigin) {
  const requestUrl = new URL(request.url);
  if (requestUrl.pathname === '/api' || requestUrl.pathname.startsWith('/api/')) {
    const target = new URL(`${requestUrl.pathname}${requestUrl.search}`, backendOrigin);
    const hasBody = request.method !== 'GET' && request.method !== 'HEAD';
    return net.fetch(target.toString(), {
      method: request.method,
      headers: request.headers,
      ...(hasBody ? { body: request.body, duplex: 'half' } : {}),
    });
  }
  if (requestUrl.pathname === '/ws' || requestUrl.pathname.startsWith('/ws/')) {
    return new Response('WebSocket connections must use the configured backend ws endpoint', { status: 426 });
  }
  const asset = resolveWebAsset(request.url);
  if (asset) return net.fetch(pathToFileURL(asset).toString());
  if (request.headers.get('accept')?.includes('text/html')) {
    return net.fetch(pathToFileURL(path.join(webDirectory, 'index.html')).toString());
  }
  return new Response('Not found', { status: 404 });
}

async function createWindow(backendOrigin) {
  const collaborationUrl = resolveCollaborationUrl(backendOrigin);
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
      preload: path.join(electronDirectory, 'preload.mjs'),
      additionalArguments: [collaborationArgument(collaborationUrl)],
    },
  });
  window.once('ready-to-show', () => window.show());
  await window.loadURL('app://-/');
}

app.whenReady().then(async () => {
  const backendOrigin = resolveBackendOrigin(process.env.REACT_SHEETS_API_ORIGIN);
  protocol.handle('app', (request) => handleAppRequest(request, backendOrigin));
  await createWindow(backendOrigin);
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) void createWindow(backendOrigin);
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
