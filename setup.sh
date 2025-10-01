# Create SSL certificates directory
if [ ! -d certs ]; then
    echo "🔐 Creating SSL certificates..."
    mkdir certs
    cd certs
    
    # Check if OpenSSL is available
    if ! command -v openssl &> /dev/null; then
        echo "❌ OpenSSL is not installed or not in PATH."
        echo "   On Windows, you can:"
        echo "   - Install Git for Windows (includes OpenSSL)"
        echo "   - Install OpenSSL directly"
        echo "   - Use Windows Subsystem for Linux (WSL)"
        cd ..
        exit 1
    fi
    
    # Generate private key
    echo "   Generating private key..."
    openssl genrsa -out server.key 2048
    
    # Generate certificate signing request with default values
    echo "   Generating certificate signing request..."
    openssl req -new -key server.key -out server.csr -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost"
    
    # Generate self-signed certificate
    echo "   Generating self-signed certificate..."
    openssl x509 -req -days 365 -in server.csr -signkey server.key -out server.crt
    
    # Clean up CSR file
    rm -f server.csr
    
    cd ..
    echo "✅ SSL certificates created in ./certs directory"
    echo ""
    
    # Platform-specific certificate trust instructions
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "🔒 To trust the certificate on macOS, run:"
        echo "   sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/server.crt"
        echo ""
    elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" ]]; then
        echo "🔒 On Windows, you may need to:"
        echo "   1. Double-click certs/server.crt to install it"
        echo "   2. Choose 'Local Machine' and 'Trusted Root Certification Authorities'"
        echo "   3. Or run: certlm.msc and manually import the certificate"
        echo ""
    fi
else
    echo "✅ SSL certificates directory already exists."
    echo ""
fi
