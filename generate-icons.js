const fs = require('fs');
const path = require('path');

// Create assets directory if it doesn't exist
const assetsDir = path.join(__dirname, 'assets');
if (!fs.existsSync(assetsDir)) {
    fs.mkdirSync(assetsDir);
}

// Simple SVG icon template
const createSvgIcon = (size) => {
    return `<?xml version="1.0" encoding="UTF-8"?>
<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" xmlns="http://www.w3.org/2000/svg">
    <defs>
        <linearGradient id="gradient" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" style="stop-color:#667eea;stop-opacity:1" />
            <stop offset="100%" style="stop-color:#764ba2;stop-opacity:1" />
        </linearGradient>
    </defs>
    <rect width="${size}" height="${size}" rx="${size/8}" fill="url(#gradient)"/>
    <g transform="translate(${size/2}, ${size/2})">
        <!-- SharePoint icon representation -->
        <circle cx="0" cy="-${size/6}" r="${size/12}" fill="white" opacity="0.9"/>
        <circle cx="-${size/6}" cy="${size/8}" r="${size/12}" fill="white" opacity="0.9"/>
        <circle cx="${size/6}" cy="${size/8}" r="${size/12}" fill="white" opacity="0.9"/>
        <path d="M 0,-${size/6} L -${size/6},${size/8} L ${size/6},${size/8} Z" 
              stroke="white" stroke-width="${size/32}" fill="none" opacity="0.7"/>
    </g>
    <text x="${size/2}" y="${size*0.85}" font-family="Arial, sans-serif" font-size="${size/4}" 
          font-weight="bold" fill="white" text-anchor="middle">SP</text>
</svg>`;
};

// Generate icons in different sizes
const sizes = [16, 32, 64, 80, 128];

sizes.forEach(size => {
    const svg = createSvgIcon(size);
    const filename = path.join(assetsDir, `icon-${size}.svg`);
    fs.writeFileSync(filename, svg);
    console.log(`Created ${filename}`);
});

// Create a simple PNG placeholder (base64 encoded 1x1 transparent pixel)
// In production, you would convert the SVGs to PNGs using a proper image library
const transparentPixel = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');

sizes.forEach(size => {
    const filename = path.join(assetsDir, `icon-${size}.png`);
    // For now, create placeholder PNGs
    // In production, use a library like sharp or canvas to convert SVG to PNG
    fs.writeFileSync(filename, transparentPixel);
    console.log(`Created placeholder ${filename}`);
});

console.log('\nIcon generation complete!');
console.log('Note: PNG files are placeholders. For production, convert SVGs to PNGs using an image processing library.');
console.log('\nTo convert SVGs to PNGs, you can install and use sharp:');
console.log('npm install sharp');
console.log('Then use sharp to convert the SVG files to PNG format.');
