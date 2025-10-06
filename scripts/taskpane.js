// MSAL configuration
const msalConfig = {
    auth: {
        clientId: "6677d310-6f8a-48d7-a24d-f9a8a56ce49b",
        authority: "https://login.microsoftonline.com/15dc6d2d-a6f3-4081-ad30-aa3d729969ba",
        redirectUri: "https://localhost:3000/taskpane.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

// Initialize MSAL
const msalInstance = new msal.PublicClientApplication(msalConfig);

// SharePoint data storage
let sharePointData = [];
let accessToken = null;
let currentUser = null;

// Office initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Initialize UI event handlers
        document.getElementById("fetchDataButton").onclick = fetchSharePointData;
        document.getElementById("insertDataButton").onclick = insertDataIntoDocument;
        
        // Load saved configuration
        loadConfiguration();
        
        console.log("Add-in ready - authentication will happen automatically when fetching data");
    }
});


// Get user info from Microsoft Graph (for SSO scenarios)
async function getUserInfo() {
    try {
        const graphClient = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
        
        const userInfo = await graphClient.api('/me').get();
        
        // Create a user object similar to MSAL format
        currentUser = {
            username: userInfo.userPrincipalName || userInfo.mail,
            name: userInfo.displayName,
            id: userInfo.id
        };
        
    } catch (error) {
        console.error("Failed to get user info:", error);
        // Set a default user object if Graph call fails
        currentUser = {
            username: "Current User",
            name: "Current User"
        };
    }
}


// Get access token using Office SSO first, fallback to MSAL
async function getAccessToken() {
    try {
        // Try Office SSO first (reuses Word's authentication)
        console.log("Attempting Office SSO authentication...");
        const ssoToken = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true
        });
        
        console.log("Office SSO successful");
        accessToken = ssoToken;
        return accessToken;
        
    } catch (ssoError) {
        console.log("SSO failed, falling back to MSAL:", ssoError);
        
        // Define tokenRequest outside the nested try blocks
        const tokenRequest = {
            scopes: ["Sites.Read.All", "Sites.ReadWrite.All", "Files.Read.All"],
            account: currentUser
        };
        
        // Fallback to MSAL authentication
        try {
            const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
            accessToken = tokenResponse.accessToken;
            return accessToken;
            
        } catch (error) {
            console.error("Token acquisition failed:", error);
            // Try to acquire token interactively
            try {
                const tokenResponse = await msalInstance.acquireTokenPopup(tokenRequest);
                accessToken = tokenResponse.accessToken;
                return accessToken;
            } catch (popupError) {
                console.error("Interactive token acquisition failed:", popupError);
                throw popupError;
            }
        }
    }
}

// Fetch SharePoint data
async function fetchSharePointData() {
    try {
        showLoading(true);
        clearMessages();
        
        // Authenticate automatically in background
        console.log("Authenticating automatically...");
        await getAccessToken();
        
        // Get user info if not already available
        if (!currentUser) {
            await getUserInfo();
        }
        
        const configResponse = await fetch('/config');
        const config = await configResponse.json();
        
        const siteUrl = config.sharePointSiteUrl;
        const listName = config.sharePointListName;
        
        if (!siteUrl || !listName) {
            showError("Please configure SharePoint site URL and list name in environment variables.");
            showLoading(false);
            return;
        }
        
        // Save configuration
        saveConfiguration(siteUrl, listName);
        
        // Parse site URL to get site ID
        const siteId = await getSiteId(siteUrl);
        
        // Fetch list items
        const graphClient = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
        
        // Get list ID
        const listsResponse = await graphClient
            .api(`/sites/${siteId}/lists`)
            .filter(`displayName eq '${listName}'`)
            .get();
        
        if (!listsResponse.value || listsResponse.value.length === 0) {
            throw new Error(`List '${listName}' not found in the SharePoint site.`);
        }
        
        const listId = listsResponse.value[0].id;
        
        // Get list items
        const itemsResponse = await graphClient
            .api(`/sites/${siteId}/lists/${listId}/items`)
            .expand("fields")
            .top(50) // Limit to 50 items for demo
            .get();
        
        sharePointData = itemsResponse.value;
        
        // Display data preview
        displayDataPreview(sharePointData);
        
        // Enable insert button
        document.getElementById("insertDataButton").disabled = false;
        
        showSuccess(`Successfully fetched ${sharePointData.length} items from SharePoint!`);
        
    } catch (error) {
        console.error("Failed to fetch SharePoint data:", error);
        showError("Failed to fetch data: " + error.message);
    } finally {
        showLoading(false);
    }
}

// Get SharePoint site ID from URL
async function getSiteId(siteUrl) {
    try {
        const url = new URL(siteUrl);
        const hostname = url.hostname;
        const sitePath = url.pathname;
        
        const graphClient = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
        
        // Construct site path for Graph API
        let siteApiPath;
        if (sitePath && sitePath !== '/') {
            // Remove /sites/ prefix if present
            const cleanPath = sitePath.replace(/^\/sites\//, '');
            siteApiPath = `/sites/${hostname}:/sites/${cleanPath}`;
        } else {
            siteApiPath = `/sites/${hostname}`;
        }
        
        const siteResponse = await graphClient
            .api(siteApiPath)
            .get();
        
        return siteResponse.id;
        
    } catch (error) {
        console.error("Failed to get site ID:", error);
        throw new Error("Failed to connect to SharePoint site. Please check the URL.");
    }
}

// Display data preview
function displayDataPreview(data) {
    const previewDiv = document.getElementById("dataPreview");
    
    if (!data || data.length === 0) {
        previewDiv.innerHTML = '<p class="placeholder">No data found in the specified list.</p>';
        return;
    }
    
    let html = '<div class="data-table-container"><table class="data-table"><thead><tr>';
    
    // Get column headers from the first item's fields
    const firstItem = data[0];
    const fields = firstItem.fields || {};
    const columns = Object.keys(fields).filter(key => !key.startsWith('@'));
    
    // Add headers
    columns.forEach(column => {
        html += `<th>${column}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // Add data rows (limit to first 10 for preview)
    const previewData = data.slice(0, 10);
    previewData.forEach(item => {
        html += '<tr>';
        columns.forEach(column => {
            const value = item.fields[column];
            const displayValue = value !== null && value !== undefined ? value : '';
            html += `<td>${escapeHtml(String(displayValue))}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    
    if (data.length > 10) {
        html += `<p class="preview-note">Showing first 10 of ${data.length} items</p>`;
    }
    
    previewDiv.innerHTML = html;
}

// Insert data into Word document
async function insertDataIntoDocument() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            
            // Clear existing content (optional - remove if you want to append)
            // body.clear();
            
            // Insert title
            const title = body.insertParagraph("SharePoint Data", Word.InsertLocation.end);
            title.styleBuiltIn = Word.Style.heading1;
            
            // Insert timestamp
            const timestamp = body.insertParagraph(
                `Data fetched on: ${new Date().toLocaleString()}`,
                Word.InsertLocation.end
            );
            timestamp.styleBuiltIn = Word.Style.subtitle;
            
            // Insert source information
            const siteUrl = document.getElementById("siteUrl").value;
            const listName = document.getElementById("listName").value;
            const sourceInfo = body.insertParagraph(
                `Source: ${listName} from ${siteUrl}`,
                Word.InsertLocation.end
            );
            sourceInfo.styleBuiltIn = Word.Style.quote;
            
            // Add space
            body.insertParagraph("", Word.InsertLocation.end);
            
            if (sharePointData.length === 0) {
                body.insertParagraph("No data to insert.", Word.InsertLocation.end);
                await context.sync();
                return;
            }
            
            // Get columns from first item
            const firstItem = sharePointData[0];
            const fields = firstItem.fields || {};
            const columns = Object.keys(fields).filter(key => !key.startsWith('@'));
            
            // Create table data
            const tableData = [];
            
            // Add header row
            tableData.push(columns);
            
            // Add data rows
            sharePointData.forEach(item => {
                const row = columns.map(column => {
                    const value = item.fields[column];
                    return value !== null && value !== undefined ? String(value) : '';
                });
                tableData.push(row);
            });
            
            // Insert table
            const table = body.insertTable(
                tableData.length,
                columns.length,
                Word.InsertLocation.end,
                tableData
            );
            
            // Style the table
            table.styleBuiltIn = Word.Style.gridTable4_Accent5;
            table.headerRowCount = 1;
            
            await context.sync();
            
            showSuccess(`Successfully inserted ${sharePointData.length} items into the document!`);
        });
    } catch (error) {
        console.error("Failed to insert data:", error);
        showError("Failed to insert data into document: " + error.message);
    }
}


// Configuration management
function saveConfiguration(siteUrl, listName) {
    localStorage.setItem("sharepoint_siteUrl", siteUrl);
    localStorage.setItem("sharepoint_listName", listName);
}

function loadConfiguration() {
    const savedSiteUrl = localStorage.getItem("sharepoint_siteUrl");
    const savedListName = localStorage.getItem("sharepoint_listName");
    
    if (savedSiteUrl) {
        document.getElementById("siteUrl").value = savedSiteUrl;
    }
    if (savedListName) {
        document.getElementById("listName").value = savedListName;
    }
}

// Utility functions
function showLoading(show) {
    document.getElementById("loadingIndicator").style.display = show ? "flex" : "none";
}

function showError(message) {
    const errorDiv = document.getElementById("errorMessage");
    errorDiv.textContent = message;
    errorDiv.style.display = "block";
    setTimeout(() => {
        errorDiv.style.display = "none";
    }, 5000);
}

function showSuccess(message) {
    const successDiv = document.getElementById("successMessage");
    successDiv.textContent = message;
    successDiv.style.display = "block";
    setTimeout(() => {
        successDiv.style.display = "none";
    }, 5000);
}

function clearMessages() {
    document.getElementById("errorMessage").style.display = "none";
    document.getElementById("successMessage").style.display = "none";
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}
