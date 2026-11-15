const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = parseInt(process.env.OFFICE_PORT || '3002', 10);
const dist = path.join(__dirname, '..', 'dist');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.xml': 'application/xml; charset=utf-8',
  '.svg': 'image/svg+xml',
};

function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', '*');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const cleanUrl = req.url.split('?')[0];
  let filePath = path.join(dist, cleanUrl === '/' ? 'word/taskpane.html' : cleanUrl);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) {
    filePath = path.join(filePath, 'taskpane.html');
  }

  if (!fs.existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('404 Not Found: ' + cleanUrl);
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  res.writeHead(200, { 'Content-Type': contentType });
  fs.createReadStream(filePath).pipe(res);
}

// Check for local SSL certs in certs/ directory if provided
const certPath = path.join(__dirname, '..', 'certs', 'localhost.crt');
const keyPath = path.join(__dirname, '..', 'certs', 'localhost.key');

if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
  const options = {
    key: fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath),
  };
  https.createServer(options, handler).listen(PORT, () => {
    console.log(`[OrcaX Office.js] HTTPS Server đang chạy tại https://localhost:${PORT}`);
  });
} else {
  http.createServer(handler).listen(PORT, () => {
    console.log(`[OrcaX Office.js] HTTP Server đang chạy tại http://localhost:${PORT}`);
    console.log(
      `(Lưu ý: Để sideload vào Microsoft Word Desktop cần HTTPS. Xem hướng dẫn tại docs/architecture/add-ins.md)`,
    );
  });
}
