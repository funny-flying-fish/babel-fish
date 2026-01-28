/**
 * Local development server with Lokalise API proxy
 * Handles CORS by proxying requests to api.lokalise.com
 * 
 * Usage: node server.js
 */

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = 8080;
const LOKALISE_API_HOST = 'api.lokalise.com';

const MIME_TYPES = {
    '.html': 'text/html',
    '.css': 'text/css',
    '.js': 'application/javascript',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon'
};

function serveStatic(req, res) {
    let filePath = req.url === '/' ? '/index.html' : req.url;
    filePath = filePath.split('?')[0]; // Remove query string
    const fullPath = path.join(__dirname, filePath);
    const ext = path.extname(fullPath);
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(fullPath, (err, data) => {
        if (err) {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('File not found');
            return;
        }
        res.writeHead(200, { 
            'Content-Type': contentType,
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        });
        res.end(data);
    });
}

function proxyToLokalise(req, res) {
    const apiPath = req.url.replace('/api/lokalise', '/api2');
    
    // Collect request body
    let body = '';
    req.on('data', chunk => {
        body += chunk.toString();
    });

    req.on('end', () => {
        const options = {
            hostname: LOKALISE_API_HOST,
            port: 443,
            path: apiPath,
            method: req.method,
            headers: {
                'Content-Type': 'application/json',
                'X-Api-Token': req.headers['x-api-token'] || ''
            }
        };

        if (body && req.method !== 'GET') {
            options.headers['Content-Length'] = Buffer.byteLength(body);
        }

        const proxyReq = https.request(options, (proxyRes) => {
            // Add CORS headers
            res.writeHead(proxyRes.statusCode, {
                ...proxyRes.headers,
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type, X-Api-Token'
            });

            proxyRes.pipe(res);
        });

        proxyReq.on('error', (err) => {
            console.error('Proxy error:', err.message);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: { message: err.message } }));
        });

        if (body && req.method !== 'GET') {
            proxyReq.write(body);
        }
        proxyReq.end();
    });
}

function handleCORS(req, res) {
    res.writeHead(200, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, X-Api-Token',
        'Access-Control-Max-Age': '86400'
    });
    res.end();
}

const server = http.createServer((req, res) => {
    console.log(`${req.method} ${req.url}`);

    // Handle CORS preflight
    if (req.method === 'OPTIONS') {
        return handleCORS(req, res);
    }

    // Proxy API requests to Lokalise
    if (req.url.startsWith('/api/lokalise')) {
        return proxyToLokalise(req, res);
    }

    // Serve static files
    serveStatic(req, res);
});

server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}/`);
    console.log(`Lokalise API proxy available at http://localhost:${PORT}/api/lokalise/`);
});
