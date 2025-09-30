#!/bin/bash

echo "========================================="
echo "Starting Word Add-in with ngrok"
echo "========================================="
echo ""

# Check if ngrok is installed
if ! command -v ngrok &> /dev/null; then
    echo "❌ ngrok is not installed."
    echo "Installing ngrok..."
    npm install -g ngrok
fi

# Start the server in background
echo "🚀 Starting local server..."
npm start &
SERVER_PID=$!

# Wait for server to start
sleep 3

# Start ngrok
echo "🌐 Starting ngrok tunnel..."
ngrok http https://localhost:3000 &
NGROK_PID=$!

# Wait for ngrok to start
sleep 3

# Get ngrok URL
echo ""
echo "========================================="
echo "📋 Next Steps:"
echo "========================================="
echo ""
echo "1. Copy the ngrok HTTPS URL from the terminal above"
echo "   (looks like: https://xxxxx.ngrok.io)"
echo ""
echo "2. Update manifest-sharepoint.xml:"
echo "   Replace all instances of https://localhost:3000"
echo "   with your ngrok URL"
echo ""
echo "3. Also update scripts/taskpane.js:"
echo "   Update redirectUri to use your ngrok URL"
echo ""
echo "4. Upload the updated manifest to SharePoint:"
echo "   - Go to SharePoint App Catalog"
echo "   - Upload manifest-sharepoint.xml"
echo "   - Or sideload in Word Online"
echo ""
echo "5. Press Ctrl+C to stop both servers when done"
echo ""
echo "========================================="

# Wait for user to stop
wait $SERVER_PID
wait $NGROK_PID
