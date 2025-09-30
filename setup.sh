#!/bin/bash

echo "========================================="
echo "SharePoint Word Add-in Setup Script"
echo "========================================="
echo ""

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js first."
    echo "   Visit: https://nodejs.org/"
    exit 1
fi

echo "✅ Node.js is installed: $(node -v)"
echo ""

# Install dependencies
echo "📦 Installing npm dependencies..."
npm install
echo ""

# Generate icons
echo "🎨 Generating icon files..."
node generate-icons.js
echo ""

# Create .env file if it doesn't exist
if [ ! -f .env ]; then
    echo "📝 Creating .env file from template..."
    cp .env.example .env
    echo "✅ .env file created. Please update it with your Azure AD credentials."
    echo ""
else
    echo "✅ .env file already exists."
    echo ""
fi

# Create SSL certificates directory
if [ ! -d certs ]; then
    echo "🔐 Creating SSL certificates..."
    mkdir certs
    cd certs
    
    # Generate private key
    openssl genrsa -out server.key 2048 2>/dev/null
    
    # Generate certificate signing request with default values
    openssl req -new -key server.key -out server.csr -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost" 2>/dev/null
    
    # Generate self-signed certificate
    openssl x509 -req -days 365 -in server.csr -signkey server.key -out server.crt 2>/dev/null
    
    cd ..
    echo "✅ SSL certificates created in ./certs directory"
    echo ""
    
    # Prompt to trust certificate on macOS
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "🔒 To trust the certificate on macOS, run:"
        echo "   sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/server.crt"
        echo ""
    fi
else
    echo "✅ SSL certificates directory already exists."
    echo ""
fi

echo "========================================="
echo "📋 Next Steps:"
echo "========================================="
echo ""
echo "1. Update Azure AD Configuration:"
echo "   Edit the .env file and add your Azure AD credentials"
echo "   Edit scripts/taskpane.js and update:"
echo "   - YOUR_APP_CLIENT_ID"
echo "   - YOUR_TENANT_ID"
echo ""
echo "2. Trust the SSL certificate (if on macOS):"
echo "   sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/server.crt"
echo ""
echo "3. Start the server:"
echo "   npm start"
echo ""
echo "4. Sideload the add-in in Word:"
echo "   - Open Word"
echo "   - Go to Insert > Add-ins > My Add-ins"
echo "   - Click 'Upload My Add-in'"
echo "   - Select the manifest.xml file"
echo ""
echo "========================================="
echo "✨ Setup complete! Happy coding!"
echo "========================================="
