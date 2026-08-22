const http = require('http');
const fs = require('fs');
const path = require('path');

const dist = path.join(__dirname, '..', 'frontend', 'dist');
const backend = { host: '127.0.0.1', port: 9004 };
const PORT = 3000;

const types = {
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
};

function proxyHttp(req, res, target = backend, backendPath) {
  const options = {
    hostname: target.host,
    port: target.port,
    path: backendPath || req.url,
    method: req.method,
    headers: { ...req.headers, host: `${target.host}:${target.port}` },
  };

  const proxyReq = http.request(options, (proxyRes) => {
    res.writeHead(proxyRes.statusCode, proxyRes.headers);
    proxyRes.pipe(res);
  });

  proxyReq.on('error', (err) => {
    res.writeHead(502);
    res.end(`Backend proxy error: ${err.message}`);
  });

  req.pipe(proxyReq);
}

function serveStatic(req, res) {
  let filePath = path.join(dist, req.url === '/' ? 'index.html' : req.url.split('?')[0]);
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    filePath = path.join(dist, 'index.html');
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    const ext = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': types[ext] || 'application/octet-stream' });
    res.end(data);
  });
}

const server = http.createServer((req, res) => {
  if (req.url.startsWith('/luckysheet/')) {
    return proxyHttp(req, res, backend);
  }
  // 兼容旧 dist 仍请求 /luckyToXlsx、/luckyexcel/*
  if (req.url.startsWith('/luckyToXlsx') || req.url.startsWith('/luckyexcel')) {
    const queryIndex = req.url.indexOf('?');
    const pathname = queryIndex >= 0 ? req.url.slice(0, queryIndex) : req.url;
    const query = queryIndex >= 0 ? req.url.slice(queryIndex) : '';
    return proxyHttp(req, res, backend, `/luckysheet${pathname}${query}`);
  }
  return serveStatic(req, res);
});

server.on('upgrade', (req, socket, head) => {
  if (!req.url.startsWith('/luckysheet/')) {
    socket.destroy();
    return;
  }

  const options = {
    hostname: backend.host,
    port: backend.port,
    path: req.url,
    method: req.method,
    headers: { ...req.headers, host: `${backend.host}:${backend.port}` },
  };

  const proxyReq = http.request(options);
  proxyReq.on('upgrade', (proxyRes, proxySocket, proxyHead) => {
    socket.write(`HTTP/1.1 ${proxyRes.statusCode} ${proxyRes.statusMessage}\r\n`);
    Object.entries(proxyRes.headers).forEach(([key, value]) => {
      socket.write(`${key}: ${value}\r\n`);
    });
    socket.write('\r\n');
    if (proxyHead && proxyHead.length) {
      socket.write(proxyHead);
    }
    proxySocket.pipe(socket);
    socket.pipe(proxySocket);
  });

  proxyReq.on('error', () => socket.destroy());
  proxyReq.end();
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`Luckysheet frontend: http://127.0.0.1:${PORT}/`);
  console.log(
    `Collaboration: http://127.0.0.1:${PORT}/?share=1&gridKey=1079500#-8803#7c45f52b7d01486d88bc53cb17dcd2c3`
  );
  console.log('Excel import/export: Java /luckysheet/luckyToXlsx and /luckysheet/luckyexcel/upload');
});
