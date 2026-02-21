const config = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "http://localhost:5500" // Critical for GitHub Pages
    }
};

//const msalInstance = new msal.PublicClientApplication(config);

let account = null;
let activeSheet = null;
const fileName = "MAE_Master_Inventory_Template.xlsx"

let msalInstance;
async function initializeMsal() {
    msalInstance = new msal.PublicClientApplication(config);
    await msalInstance.initialize();

// Required for Redirect flows, good practice to run on every load
    msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        account = response.account;
        updateUI();
    } else {
        const currentAccounts = msalInstance.getAllAccounts();
        if (currentAccounts.length > 0) {
            account = currentAccounts[0];
            updateUI();
        }
    }
    });

async function signIn() {
    const loginRequest = {
        scopes: ["User.Read", "Files.ReadWrite.All"] // Request permissions for Excel
    };

    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        account = loginResponse.account;
        updateUI();
    } catch (error) {
        console.error("Login failed:", error);
    }
}

function updateUI() {
    const authBtn = document.getElementById('auth-btn');
    if (account) {
        authBtn.innerText = `Connected: ${account.username}`;
        authBtn.style.background = "#27ae60";
        
        // NEW: Scan the Excel file to see what features this client has
        loadDynamicMenu(); 
    }
}


// Attach the event listener to your existing button
document.getElementById('auth-btn').addEventListener('click', signIn);

async function loadDynamicMenu() {
    const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Files.ReadWrite.All"]
    });

    // 1. Ask Graph for ALL tables in the workbook
    // Use backticks (`) for the URL to allow the ${variable} syntax
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${fileName}:/workbook/tables`;

    
    const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
    });
    const data = await response.json();

    // 2. Clear and rebuild the sidebar menu based on what's actually there
    const menu = document.getElementById('menu');
    menu.innerHTML = ""; // Clear existing

    data.value.forEach(table => {
        const li = document.createElement('li');
        // Clean up the name for the button (e.g., "Shop_Machinery_Table" -> "Shop Machinery")
        li.innerText = table.name.replace(/_/g, ' ').replace('Table', '');
        li.onclick = () => fetchTableData(table.name);
        menu.appendChild(li);
    });
}


async function switchSheet(sheetName) {
    activeSheet = sheetName;
    document.getElementById('current-view-title').innerText = sheetName;
    
    // Map internal names to your actual Excel Worksheet names
    const sheetMap = {
        'ResellInventory': 'Resell_Inventory_Table',
        'Tools': 'Shop_Tools_Table',
        'Suppliers': 'Supplier_Contact_List'
    };
    
    fetchTableData(sheetMap[sheetName]);
}

async function fetchTableData(tableName) {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root/:${fileName}:/workbook/tables/${tableName}/rows`;

    // ... same fetch logic as previous example ...
}

