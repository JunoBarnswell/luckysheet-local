const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const dist = path.join(root, 'frontend-react', 'dist', 'web');
const backend = { host: '127.0.0.1', port: Number(process.env.BACKEND_PORT ?? 8082) };
const port = Number(process.env.PORT ?? 4180);
const types = { '.html': 'text/html; charset=utf-8', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json', '.wasm': 'application/wasm', '.png': 'image/png', '.svg': 'image/svg+xml' };

function proxy(req, res) {
  const upstream = http.request({ hostname: backend.host, port: backend.port, path: req.url, method: req.method, headers: { ...req.headers, host: `${backend.host}:${backend.port}` } }, (reply) => { res.writeHead(reply.statusCode ?? 502, reply.headers); reply.pipe(res); });
  upstream.on('error', (error) => { res.writeHead(502, { 'content-type': 'text/plain; charset=utf-8' }); res.end(`Backend unavailable: ${error.message}`); });
  req.pipe(upstream);
}

const server = http.createServer((req, res) => {
  if (req.url?.startsWith('/api/') || req.url?.startsWith('/ws')) return proxy(req, res);
  const pathname = (req.url ?? '/').split('?')[0];
  const candidate = path.resolve(dist, `.${pathname === '/' ? '/index.html' : pathname}`);
  const file = candidate.startsWith(`${dist}${path.sep}`) ? candidate : path.join(dist, 'index.html');
  const actual = fs.existsSync(file) && fs.statSync(file).isFile() ? file : path.join(dist, 'index.html');
  fs.readFile(actual, (error, data) => { if (error) { res.writeHead(404); res.end('Build output is unavailable; run scripts/build.ps1'); return; } res.writeHead(200, { 'content-type': types[path.extname(actual)] ?? 'application/octet-stream' }); res.end(data); });
});

server.on('upgrade', (req, socket, head) => {
  if (!req.url?.startsWith('/ws')) {
    socket.destroy();
    return;
  }
  const upstream = http.request({
    hostname: backend.host,
    port: backend.port,
    path: req.url,
    method: req.method,
    headers: { ...req.headers, host: `${backend.host}:${backend.port}` },
  });
  upstream.on('upgrade', (reply, upstreamSocket, upstreamHead) => {
    const headers = Object.entries(reply.headers)
      .flatMap(([name, value]) => Array.isArray(value) ? value.map((entry) => `${name}: ${entry}`) : value === undefined ? [] : [`${name}: ${value}`]);
    socket.write(`HTTP/${reply.httpVersion} ${reply.statusCode} ${reply.statusMessage}\r\n${headers.join('\r\n')}\r\n\r\n`);
    if (head.length > 0) upstreamSocket.write(head);
    if (upstreamHead.length > 0) socket.write(upstreamHead);
    upstreamSocket.pipe(socket);
    socket.pipe(upstreamSocket);
  });
  upstream.on('response', (reply) => {
    socket.write(`HTTP/${reply.httpVersion} ${reply.statusCode} ${reply.statusMessage}\r\nConnection: close\r\n\r\n`);
    socket.destroy();
  });
  upstream.on('error', () => socket.destroy());
  upstream.end();
});
server.listen(port, '127.0.0.1', () => console.log(`React web app: http://127.0.0.1:${port}/`));
