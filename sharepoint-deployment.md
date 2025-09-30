# SharePoint Deployment Guide

## Issue: Add-in Error in SharePoint

When you see an "add-in error" in SharePoint while the add-in works locally, it's typically due to one of these issues:

### Common Causes & Solutions

## 1. **Certificate Trust Issues**

SharePoint Online requires HTTPS with a valid certificate. Self-signed certificates often cause issues.

### Solution A: Use ngrok for Testing (Recommended for Development)
```bash
# Install ngrok
npm install -g ngrok

# Start your local server
npm start

# In another terminal, expose your local server
ngrok http https://localhost:3000
```

Then update your manifest to use the ngrok URL (e.g., `https://abc123.ngrok.io`)

### Solution B: Deploy to Azure App Service
Deploy your add-in to Azure for a proper HTTPS endpoint with valid certificates.

## 2. **App Catalog Deployment**

For SharePoint Online, you need to deploy through the App Catalog:

1. **Access the App Catalog**:
   - Go to SharePoint Admin Center
   - Navigate to More features → Apps → App Catalog
   - If no catalog exists, create one

2. **Upload the Add-in**:
   - Go to "Apps for Office"
   - Click "New" → "Upload"
   - Select your `manifest-sharepoint.xml`
   - Click "OK" to deploy

3. **Trust the Add-in**:
   - After upload, click on the add-in
   - Select "Trust It"

## 3. **Authentication Configuration**

### Update Redirect URI in Azure AD:
1. Go to Azure Portal → App Registrations
2. Select your app (ID: 02fb26ab-51b3-4b53-ad60-1ad1270db6dd)
3. Go to Authentication
4. Add these redirect URIs:
   - `https://localhost:3000/taskpane.html`
   - `https://your-sharepoint-site.sharepoint.com/_layouts/15/appregnew.aspx`
   - Your ngrok URL if using ngrok

### Update API Permissions:
Ensure these permissions are granted with Admin Consent:
- Sites.Read.All
- Sites.ReadWrite.All
- Files.Read.All
- User.Read

## 4. **Manifest Configuration for SharePoint**

Use the `manifest-sharepoint.xml` file which includes:
- All necessary AppDomains
- Proper authentication endpoints
- SharePoint-compatible settings

## 5. **Testing in SharePoint**

### Method 1: Test in Word Online
1. Open Word Online in SharePoint
2. Go to Insert → Add-ins → My Organization
3. Your add-in should appear if properly deployed

### Method 2: Direct Sideloading
1. Open a Word document in SharePoint
2. Go to Insert → Add-ins → Upload My Add-in
3. Upload `manifest-sharepoint.xml`

## Quick Fix Steps

1. **Run with ngrok**:
```bash
# Terminal 1
npm start

# Terminal 2
ngrok http https://localhost:3000
```

2. **Update manifest with ngrok URL**:
   - Replace all `https://localhost:3000` with your ngrok URL
   - Save as `manifest-ngrok.xml`

3. **Upload to SharePoint**:
   - Use the App Catalog method above
   - Or sideload directly in Word Online

## Debugging Tips

1. **Check Browser Console**:
   - Press F12 in browser
   - Look for CORS errors or authentication failures

2. **Verify Certificate**:
   - Navigate to https://localhost:3000 directly
   - Ensure certificate is trusted

3. **Test Authentication**:
   - Try signing in from the add-in
   - Check if token is acquired successfully

4. **SharePoint Logs**:
   - Check SharePoint ULS logs for detailed errors
   - Use Fiddler to inspect network traffic

## Alternative: Use Office Add-in Debugger

For easier testing without SharePoint deployment:

1. Install Office Add-in Debugger VS Code extension
2. Configure launch.json:
```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "type": "office-addin",
            "request": "launch",
            "name": "Word Desktop",
            "addinType": "taskpane",
            "addinId": "e504fb41-a92a-4526-b101-542f357b7acb",
            "addinCommandId": "Contoso.TaskpaneButton",
            "addinFile": "${workspaceFolder}/manifest.xml",
            "appType": "desktop",
            "app": "word"
        }
    ]
}
```

3. Press F5 to debug in Word Desktop

## Production Deployment

For production use:
1. Deploy to Azure App Service or similar
2. Get a proper SSL certificate
3. Update manifest with production URLs
4. Submit to Microsoft AppSource or deploy via SharePoint App Catalog

## Need Help?

If you're still seeing errors:
1. Check the specific error message in browser console
2. Verify all URLs in manifest are accessible
3. Ensure Azure AD app is properly configured
4. Test with a simple "Hello World" add-in first to isolate issues
