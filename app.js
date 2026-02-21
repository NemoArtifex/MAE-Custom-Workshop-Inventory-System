
// =============CONFIGURATION: The "Blueprint"  ======================
const msalConfig = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "http://localhost:5500" // Critical for GitHub Pages
    },
    cache: {
        cacheLocation: "sessionStorage",// use sessionStorage to avoid issues with multiple tabs/GitHub Pages
        storeAuthStateInCookie: false,
    }
};
// ===========END CONFIGURATION =============

// =========== STARTUP LOGIC ============
// Global instance variable
let myMSALObj;
let account = null;

// Startup Function: runs on page load =
async function startup(){
    try {
// Instantiate and Initialize (the v3 way)
        myMSALObj = new msal.PublicClientApplication(msalConfig);
        await myMSALObj.initialize();
// Handle returning from a redirect (ifnot using Popups)
        const response = await myMSALObj.handleRedirectPromise();
// UI Setup
        const authButton = document.getElementById("auth-btn");
        authButton.disabled = false; // Enable the button now that MSAL is ready
// Check if user is already signed in (handles page refresh)
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0) {
            myMSALObj.setActiveAccount(accounts[0]); // Set the first account as active
            updateUIForLoggedInUser(accounts[0]);
        }
// Attach the click event
        authButton.addEventListener("click", signIn);
    } catch (error) {
        console.error("Error during MSAL initialization:", error);
    }
}

//==========END STARTUP LOGIC ===========

let activeSheet = null;
const fileName = "MAE_Master_Inventory_Template.xlsx"

//===========SIGN-IN FUNCTION ==========
async function signIn() {
    const loginRequest = {
// Scopes configured in Azure
        scopes: ["User.Read", "Files.ReadWrite"] // Request permissions for Excel
    };

    try {
// Standard Pop Login
        const loginResponse = await myMSALObj.loginPopup(loginRequest);
        console.log("Login Successful:", loginResponse);
        myMSALObj.setActiveAccount(loginResponse.account); // Set the active account
        updateUIForLoggedInUser(loginResponse.account);
    } catch (error) {
        console.error("Login failed:", error);
    }
}

//=========END SIGN-IN FUNCTION ===========

// ======== FUNCTION TO UPDATE UI BASED ON LOGIN STATUS ========
function updateUIForLoggedInUser(userAccount) {
    account = userAccount; // Store the account info globally
    const authButton = document.getElementById("auth-btn");
    authButton.innerText = `Connected: ${account.username}`;
    authButton.style.background = "#27ae60"; // Change color to indicate success

    // Load the dynamic menu based on the user's Excel file
    console.log("Loading dynamic menu for user:", account.username);
    loadDynamicMenu();
}

//===========END UI UPDATE FUNCTION ===========

// Attach the event listener to your existing button
document.getElementById('auth-btn').addEventListener('click', signIn);

//======= FUNCTION Load Dynamic Menu ================
async function loadDynamicMenu() {
    const tokenResponse = await myMSALObj.acquireTokenSilent({
        scopes: ["Files.ReadWrite"]
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
        const displayName = table.name.replace(/_/g, ' ').replace('Table', '');
        li.innerText = displayName;

        li.onclick = () => {
            document.getElementById('current-view-title').innerText = displayName;
            fetchTableData(table.name);// Functon to get actual Excel rows
        };
        menu.appendChild(li);
    });
}

//=============END DYNAMIC MENU FUNCTION ==============

//======= FUNCTION SWITCH SHEET ===============
// ======  TO DO ==============
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

// ============END SWITCH SHEET FUNCTION ==============

//======= FUNCTION TO FETCH TABLE DATA ==============
async function fetchTableData(tableName) {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root/:${fileName}:/workbook/tables/${tableName}/rows`;

    // ... same fetch logic as previous example ...
}

// ============END FETCH TABLE DATA FUNCTION ==============


