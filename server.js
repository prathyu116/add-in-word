const express = require('express');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Enhanced CORS configuration for SharePoint
app.use(cors({
    origin: function (origin, callback) {
        // Allow requests with no origin (like mobile apps or curl requests)
        if (!origin) return callback(null, true);
        
        // List of allowed origins
        const allowedOrigins = [
            'https://localhost:3000',
            /^https:\/\/.*\.sharepoint\.com$/,
            /^https:\/\/.*\.office\.com$/,
            /^https:\/\/.*\.officeapps\.live\.com$/,
            'https://login.microsoftonline.com',
            'https://login.windows.net'
        ];
        
        const allowed = allowedOrigins.some(allowedOrigin => {
            if (allowedOrigin instanceof RegExp) {
                return allowedOrigin.test(origin);
            }
            return allowedOrigin === origin;
        });
        
        if (allowed) {
            callback(null, true);
        } else {
            console.log('Blocked by CORS:', origin);
            callback(null, true); // For development, allow all origins
        }
    },
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

// Serve static files
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/scripts', express.static(path.join(__dirname, 'scripts')));
app.use('/styles', express.static(path.join(__dirname, 'styles')));

// Serve HTML files
app.get('/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'taskpane.html'));
});

app.get('/test.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'test.html'));
});

app.get('/function-file/function-file.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'function-file', 'function-file.html'));
});

// Default route
app.get('/', (req, res) => {
    res.send('Word Add-in Server is running!');
});

// Start server with HTTPS for development (Office Add-ins require HTTPS)
const https = require('https');
const fs = require('fs');

// Check if certificates exist, if not, use HTTP for initial setup
try {
    const privateKey = fs.readFileSync('certs/server.key', 'utf8');
    const certificate = fs.readFileSync('certs/server.crt', 'utf8');
    const credentials = { key: privateKey, cert: certificate };
    
    const httpsServer = https.createServer(credentials, app);
    httpsServer.listen(PORT, () => {
        console.log(`HTTPS Server running on https://localhost:${PORT}`);
    });
} catch (err) {
    console.log('HTTPS certificates not found. Running on HTTP.');
    console.log('For production, please generate SSL certificates.');
    app.listen(PORT, () => {
        console.log(`HTTP Server running on http://localhost:${PORT}`);
    });
}
