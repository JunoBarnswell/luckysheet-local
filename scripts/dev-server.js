const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const repositoryRoot = path.resolve(__dirname, '..');
const webRoot = path.resolve(process.env.REACT_SHEETS_WEB_ROOT ?? path.join(repositoryRoot, 'frontend-react', 'dist', 'web'));
const backend = {
  host: process.env.REACT_SHEETS_BACKEND_HOST ?? '127.0.0.1',
  port: Number.parseInt(process.env.REACT_SHEETS_BACKEND_PORT ?? '8082', 10),
};
const port = Number.parseInt(process.env.REACT_SHEETS_PREVIEW_PORT ?? '4181', 10);

const contentTypes = new Map([
  ['.css', 'text/css; charset=utf-8'],
  ['.html', 'text/html; charset=utf-8'],
  ['.ico', 'image/x-icon'],
  ['.js', 'application/javascript; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'],
  ['.map', 'application/json; charset=utf-8'],
  ['.png', 'image/png'],
  ['.svg', 'image/svg+xml'],
  ['.ttf', 'font/ttf'],
  ['.webmanifest', 'application/manifest+json; charset=utf-8'],
  ['.woff', 'font/woff'],
  ['.woff2', 'font/woff2'],
]);

function isBackendRequest(url) {
  return url === '/health' || url === '/health/' || url === '/api' || url.startsWith('/api/') || url === '/ws' || url.startsWith('/ws/');
}

function proxyHttp(request, response) {
  const options = {
    hostname: backend.host,
    port: backend.port,
    path: request.url,
    method: request.method,
    headers: { ...request.headers, host: `${backend.host}:${backend.port}` },
  };
  const proxy = http.request(options, (upstream) => {
    response.writeHead(upstream.statusCode ?? 502, upstream.headers);
    upstream.pipe(response);
  });
  proxy.on('error', () => {
    if (!response.headersSent) response.writeHead(502, { 'content-type': 'application/json; charset=utf-8' });
    response.end(JSON.stringify({ error: 'BACKEND_UNAVAILABLE' }));
  });
  request.pipe(proxy);
}

function safeStaticPath(requestUrl) {
  const pathname = decodeURIComponent((requestUrl ?? '/').split('?', 1)[0]);
  const requested = pathname === '/' ? '/index.html' : pathname;
  const candidate = path.resolve(webRoot, `.${requested}`);
  const rootWithSeparator = webRoot.endsWith(path.sep) ? webRoot : `${webRoot}${path.sep}`;
  return candidate === webRoot || candidate.startsWith(rootWithSeparator) ? candidate : null;
}

function serveStatic(request, response) {
  const candidate = safeStaticPath(request.url);
  const indexPath = path.join(webRoot, 'index.html');
  const filePath = candidate && fs.existsSync(candidate) && fs.statSync(candidate).isFile() ? candidate : indexPath;
  fs.readFile(filePath, (error, data) => {
    if (error) {
      response.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
      response.end('Frontend build output is missing. Run scripts/build.ps1 first.');
      return;
    }
    response.writeHead(200, { 'content-type': contentTypes.get(path.extname(filePath).toLowerCase()) ?? 'application/octet-stream' });
    response.end(data);
  });
}

const server = http.createServer((request, response) => {
  if (isBackendRequest(request.url ?? '/')) {
    proxyHttp(request, response);
    return;
  }
  serveStatic(request, response);
});

server.on('upgrade', (request, socket, head) => {
  if (!(request.url ?? '').startsWith('/ws')) {
    socket.destroy();
    return;
  }
  const proxy = http.request({
    hostname: backend.host,
    port: backend.port,
    path: request.url,
    method: request.method,
    headers: { ...request.headers, host: `${backend.host}:${backend.port}` },
  });
  proxy.on('upgrade', (upstreamResponse, upstreamSocket, upstreamHead) => {
    socket.write(`HTTP/1.1 ${upstreamResponse.statusCode} ${upstreamResponse.statusMessage}\r\n`);
    for (const [key, value] of Object.entries(upstreamResponse.headers)) {
      socket.write(`${key}: ${value}\r\n`);
    }
    socket.write('\r\n');
    if (upstreamHead.length > 0) socket.write(upstreamHead);
    upstreamSocket.pipe(socket);
    socket.pipe(upstreamSocket);
  });
  proxy.on('error', () => socket.destroy());
  proxy.end();
});

server.listen(port, '127.0.0.1', () => {
  console.log(`React Sheets preview: http://127.0.0.1:${port}/`);
  console.log(`Backend proxy: http://${backend.host}:${backend.port}`);
});

function shutdown() {
  server.close(() => process.exit(0));
}
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
