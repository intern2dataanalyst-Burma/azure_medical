const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname);
const port = Number(process.env.PORT ?? 4173);
const mimeTypes = {
  '.css': 'text/css',
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.json': 'application/json',
};

http.createServer((request, response) => {
  const url = new URL(request.url ?? '/', 'http://localhost');
  if (url.pathname === '/api' || url.pathname.startsWith('/api/')) {
    const upstream = http.request({
      hostname: '127.0.0.1',
      port: 7071,
      path: `${url.pathname}${url.search}`,
      method: request.method,
      headers: { ...request.headers, host: '127.0.0.1:7071' },
    }, (upstreamResponse) => {
      response.writeHead(upstreamResponse.statusCode ?? 502, upstreamResponse.headers);
      upstreamResponse.pipe(response);
    });
    upstream.on('error', () => {
      response.writeHead(502, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ error: { message: 'API server is unavailable' } }));
    });
    request.pipe(upstream);
    return;
  }

  let pathname;
  try {
    pathname = decodeURIComponent(url.pathname === '/' ? '/index.html' : url.pathname);
  } catch {
    response.writeHead(400);
    response.end('Bad request');
    return;
  }
  const target = path.resolve(root, `.${pathname}`);
  if (target !== root && !target.startsWith(`${root}${path.sep}`)) {
    response.writeHead(403);
    response.end('Forbidden');
    return;
  }
  fs.readFile(target, (error, content) => {
    if (error) {
      response.writeHead(404);
      response.end('Not found');
      return;
    }
    response.writeHead(200, { 'Content-Type': mimeTypes[path.extname(target)] ?? 'application/octet-stream' });
    response.end(content);
  });
}).listen(port, '0.0.0.0', () => {
  console.log(`Frontend available at http://localhost:${port}`);
});