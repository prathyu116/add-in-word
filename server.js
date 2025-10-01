const express = require('express');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS for all origins during development
app.use(cors());

// Serve static files
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/scripts', express.static(path.join(__dirname, 'scripts')));
app.use('/styles', express.static(path.join(__dirname, 'styles')));

// Serve HTML files
app.get('/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'taskpane.html'));
});

app.get('/function-file/function-file.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'function-file', 'function-file.html'));
});

// Default route
app.get('/', (req, res) => {
    res.send('Word Add-in Server is running!');
});

// Configuration endpoint for SharePoint settings
app.get('/config', (req, res) => {
    res.json({
        sharePointSiteUrl: process.env.SHAREPOINT_SITE_URL || '',
        sharePointListName: process.env.SHAREPOINT_LIST_NAME || ''
    });
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
