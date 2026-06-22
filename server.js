// Simple HTTPS server with self-signed cert
const https = require('https');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const certDir = __dirname;

// Generate cert if not exists
if (!fs.existsSync(path.join(certDir, 'server.crt'))) {
  execSync(`openssl req -x509 -newkey rsa:2048 -keyout "${certDir}/server.key" -out "${certDir}/server.crt" -days 365 -nodes -subj "/CN=localhost"`, { stdio: 'inherit' });
}

const options = {
  key: fs.readFileSync(path.join(certDir, 'server.key')),
  cert: fs.readFileSync(path.join(certDir, 'server.crt')),
};

const mimeTypes = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.wasm': 'application/wasm',
};

https.createServer(options, (req, res) => {
  let filePath = path.join(certDir, req.url === '/' ? 'index.html' : req.url);
  const ext = path.extname(filePath);

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': mimeTypes[ext] || 'text/plain',
      'Access-Control-Allow-Origin': '*',
    });
    res.end(data);
  });
}).listen(9092, () => {
  console.log('HTTPS server on https://localhost:9092');
});
