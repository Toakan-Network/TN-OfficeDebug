const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

// Simple HTTP server for Office Add-in development (use HTTP for now, Office will show security warning but it works)
const PORT = 3000;

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.gif': 'image/gif',
  '.ico': 'image/x-icon',
  '.svg': 'image/svg+xml'
};

function getMimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return mimeTypes[ext] || 'application/octet-stream';
}

const server = http.createServer((req, res) => {
  // Enable CORS for Office Add-ins
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  
  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  const parsedUrl = url.parse(req.url);
  let pathname = parsedUrl.pathname;
  
  // Default to index.html
  if (pathname === '/') {
    pathname = '/index.html';
  }
  
  const filePath = path.join(__dirname, pathname);
  
  console.log(`${new Date().toISOString()} - Request: ${req.method} ${pathname}`);
  
  fs.readFile(filePath, (err, data) => {
    if (err) {
      console.error(`Error reading file ${filePath}:`, err.message);
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end(`File not found: ${pathname}`);
      return;
    }
    
    const mimeType = getMimeType(filePath);
    res.writeHead(200, { 'Content-Type': mimeType });
    res.end(data);
    console.log(`  -> Served ${pathname} (${data.length} bytes)`);
  });
});

server.on('error', (err) => {
  if (err.code === 'EACCES') {
    console.error(`Permission denied. Try running as administrator or use a different port.`);
  } else if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use. Please stop other applications using this port.`);
  } else {
    console.error('Server error:', err.message);
  }
  process.exit(1);
});

server.listen(PORT, () => {
  console.log(`\n🚀 Office Add-in HTTP Server running at http://localhost:${PORT}`);
  console.log(`📁 Serving files from: ${__dirname}`);
  console.log(`\n📋 To test your add-in:`);
  console.log(`   1. Update manifest.xml URLs to use http://localhost:${PORT} if needed`);
  console.log(`   2. Open Word`);
  console.log(`   3. Go to Insert > My Add-ins > Upload My Add-in`);
  console.log(`   4. Select manifest.xml from this folder`);
  console.log(`\n⚠️  Note: HTTP (not HTTPS) - Office will show security warnings but add-in should work`);
  console.log(`\n🛑 Press Ctrl+C to stop the server\n`);
});

process.on('SIGINT', () => {
  console.log('\n🛑 Shutting down server...');
  server.close(() => {
    console.log('✅ Server stopped');
    process.exit(0);
  });
});