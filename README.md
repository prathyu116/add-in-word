# SharePoint Word Add-in

A Microsoft Word add-in that allows users to fetch data from SharePoint lists/libraries and insert it directly into Word documents.

## Features

- **SharePoint Authentication**: Secure sign-in using Microsoft Authentication Library (MSAL)
- **Data Fetching**: Retrieve data from any SharePoint list or document library
- **Data Preview**: View fetched data in a formatted table before insertion
- **Document Integration**: Insert SharePoint data as formatted tables in Word documents
- **Persistent Configuration**: Saves your SharePoint site URL and list name for convenience

## Prerequisites

1. **Node.js** (version 14 or higher)
2. **Microsoft 365 Account** with access to SharePoint
3. **Azure AD App Registration** for authentication
4. **SSL Certificate** for local development (Office Add-ins require HTTPS)

## Setup Instructions

### 1. Clone and Install Dependencies

```bash
cd /Users/prathyu/ccs/add-in
npm install
```

### 2. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure the app:
   - Name: `SharePoint Word Add-in`
   - Supported account types: Choose based on your needs
   - Redirect URI: `Single-page application` - `https://localhost:3000/taskpane.html`
5. After registration, note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**
6. Go to **API permissions** and add:
   - Microsoft Graph:
     - `Sites.Read.All`
     - `Sites.ReadWrite.All`
     - `Files.Read.All`
     - `User.Read`
7. Grant admin consent for the permissions

### 3. Configure the Add-in

1. Copy `.env.example` to `.env`:
```bash
cp .env.example .env
```

2. Update the `.env` file with your Azure AD details

3. Update `scripts/taskpane.js`:
   - Replace `YOUR_APP_CLIENT_ID` with your Application ID
   - Replace `YOUR_TENANT_ID` with your Directory ID

### 4. Generate SSL Certificates (for local development)

```bash
mkdir certs
cd certs

# Generate private key
openssl genrsa -out server.key 2048

# Generate certificate signing request
openssl req -new -key server.key -out server.csr

# Generate self-signed certificate
openssl x509 -req -days 365 -in server.csr -signkey server.key -out server.crt

cd ..
```

### 5. Trust the Certificate (macOS)

```bash
# Add certificate to keychain
sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/server.crt
```

### 6. Start the Add-in Server

```bash
npm start
```

The server will run on `https://localhost:3000`

### 7. Sideload the Add-in in Word

#### Method 1: Using Office Add-in Debugger (Recommended)

1. Install the Office Add-in Debugger extension for VS Code
2. Open the command palette (Cmd+Shift+P)
3. Run: `Office: Sideload Add-in`
4. Select Word and choose the `manifest.xml` file

#### Method 2: Manual Sideloading

1. Open Word
2. Go to **Insert** > **Add-ins** > **My Add-ins**
3. Click **Upload My Add-in**
4. Browse and select the `manifest.xml` file
5. Click **Upload**

## Usage

1. **Open the Add-in**:
   - In Word, go to the **Home** tab
   - Click **Fetch SharePoint Data** button in the ribbon

2. **Sign In**:
   - Click **Sign In to SharePoint**
   - Enter your Microsoft 365 credentials
   - Grant permissions if prompted

3. **Configure SharePoint Source**:
   - Enter your SharePoint site URL (e.g., `https://yourcompany.sharepoint.com/sites/yoursite`)
   - Enter the list or library name (e.g., `Documents`, `Tasks`, etc.)

4. **Fetch Data**:
   - Click **Fetch Data from SharePoint**
   - Review the data in the preview section

5. **Insert into Document**:
   - Click **Insert Data into Document**
   - The data will be inserted as a formatted table at the cursor position

## Project Structure

```
/Users/prathyu/ccs/add-in/
├── manifest.xml           # Office Add-in manifest
├── package.json          # Node.js dependencies
├── server.js            # Express server
├── taskpane.html        # Main add-in UI
├── scripts/
│   └── taskpane.js      # Add-in functionality
├── styles/
│   └── taskpane.css     # Add-in styling
├── function-file/
│   └── function-file.html # Ribbon command functions
├── assets/              # Icons (to be added)
├── certs/              # SSL certificates (generated)
└── .env                # Environment variables (create from .env.example)
```

## Troubleshooting

### Certificate Issues
- Make sure the certificate is trusted in your system
- Try using Chrome or Edge if Safari has issues

### Authentication Issues
- Verify Azure AD app registration settings
- Check that all required permissions are granted
- Ensure redirect URI matches exactly

### SharePoint Access Issues
- Verify you have access to the SharePoint site
- Check the site URL and list name are correct
- Ensure the list/library exists and is accessible

### Add-in Loading Issues
- Clear Office cache: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
- Restart Word
- Re-sideload the add-in

## Security Notes

- Never commit `.env` file or expose your Azure AD credentials
- Use environment variables for sensitive information
- Implement proper token refresh logic for production
- Consider implementing server-side proxy for enhanced security

## Development

To run in development mode with auto-reload:

```bash
npm run dev
```

To validate the manifest:

```bash
npm run validate
```

## Support

For issues or questions, please check:
1. [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
2. [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)
3. [MSAL.js documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)

## License

MIT
