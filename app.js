import { maeSystemConfig } from './config.js'
import { UI} from './ui.js';
const fileName = maeSystemConfig.spreadsheetName;
window.currentTable = "";
// =============CONFIGURATION: The "Blueprint"  ======================
// Defines the configuration object for the Microsoft Authentication Libray (MSAL)
// Used to integrate Microsoft's identity and sign-in features into web apps
const msalConfig = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://login.microsoftonline.com/common",
        //redirectUri: "http://localhost:5500" ,
        redirectUri: "https://nemoartifex.github.io/MAE-Custom-Workshop-Inventory-System/",
        navigateToLoginRequestUrl: false 
    },
    // Defines how and where the app stores security tokens after received
    // Tokens stored for duration of browser's tab life 
    // "false": tells MSAL NOT to store the auth state in browser cookies  
    cache: {
        cacheLocation: "sessionStorage", // Simple and effective for workshop environments
        storeAuthStateInCookie: false,
    }
    
};
// ===========END CONFIGURATION =============

// =========== STARTUP LOGIC ============
//Initializes the authentication flow for app. Handles the moment page
//first loads, specifically checking if user is returning from a login 
//attempt or has an existing session (ie, clicked refresh)  

let myMSALObj;
let account = null;

async function startup() {
    try {
        //Intialize the PublicClientApplication
        //  MSAL V2 uses 'msal.PublicClientApplication'
        myMSALObj = new window.msal.PublicClientApplication(msalConfig);

        const response = await myMSALObj.handleRedirectPromise();
    
        if (response) {
        account = response.account;
        console.log("Login successful via redirect. Account:", account.username);
        } else {
            const accounts = myMSALObj.getAllAccounts();
            if (accounts.length > 0) account = accounts[0];
        }

        if (account) {
           updateUIForLoggedInUser(account); 
        } else {
            const authButton = document.getElementById("auth-btn");
            authButton.addEventListener("click", signIn);
        }

    } catch (error) {
        console.error("Error during MSAL startup:", error);
    }
}
//========END STARTUP LOGIC ===========
startup();

//===========SIGN-IN FUNCTION ==========
//Initiated after pushing the authButton after made active at the end of the Startup function
async function signIn() {
    const loginRequest = {
        scopes: ["User.Read", "Files.ReadWrite"],
        prompt: "select_account" // Always prompt user to select account on sign-in
    };

    try {
        console.log("Redirecting for a fresh login...");
        await myMSALObj.loginRedirect(loginRequest);
    } catch (error) {
        console.error("Login failed:", error);
    }
}
//===========END SIGN-IN FUNCTION ==========

// ======== FUNCTION TO UPDATE UI BASED ON LOGIN STATUS ========
// the startup() function calls updateUIForLoggedInUser() if successful 'login'
// changes text on button and triggers loadDynamicMenu() function  
function updateUIForLoggedInUser(userAccount) {

    UI.setConnected(userAccount.username, signOut);
    loadDynamicMenu();
}

//=====END UPDATE UI BASED ON LOGIN STATUS ========

//========SIGN-OUT FUNCTION ===========
async function signOut() {
    console.log("Starting sign-out process via redirect...");
    
    if (!account) {
        resetUI();
        return;
    }

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(account.username),
        // After Microsoft logs you out, it will send the browser back here
        postLogoutRedirectUri: window.location.origin + window.location.pathname
    };

    try {
        // This clears the session storage and redirects the whole tab
        // to the Microsoft logout page.
        await myMSALObj.logoutRedirect(logoutRequest);
    } catch (error) {
        console.error("Sign-out redirect failed:", error);
        // Fallback: If the redirect fails, at least clean up the local UI
        account = null;
        sessionStorage.clear();
        resetUI();
    }
}
//======END  SIGN-OUT FUNCTION ===========

//========FUNCTION TO RESET UI AFTER SIGN-OUT =============
// Removes signout event listener, reverts button to "login" , Clears UI   
function resetUI() {
    account = null;
    sessionStorage.clear();
    UI.setDisconnected(signIn);
}
//========END FUNCTION TO RESET UI AFTER SIGN-OUT =============

//======= FUNCTION Load Dynamic Menu ================
async function loadDynamicMenu() {
   //const menu = document.getElementById("menu");
    //menu.innerHTML = ""; // Clear any existing menu items
    
    //console.log("Building dynamic menu from config...");
    // We iterate through the CONFIG, not the Excel file. 
    // This ensures the App stays "Locked" to the business agreement.

    //FILTER: only show worksheets with active:true
    const activeWorksheets = maeSystemConfig.worksheets.filter(sheet => sheet.active !==false);

    // UI handles the creation of buttons
    UI.renderMenu(activeWorksheets, (tableName) => {
        loadTableData(tableName);
    });


    verifySpreadsheetExists();
    
}
//=======END FUNCTION Load Dynamic Menu ================

// ======= FUNCTION verifySpreadSheetExists =============
async function verifySpreadsheetExists(){
    // Logic here to check if maeSystemConfig.spreadsheetName exists
    // If 404: Call a function to CREATE the workbook using the config
    // If 200: All good, ready to work.
    const tokenResponse = await myMSALObj.acquireTokenSilent({
        scopes: ["Files.ReadWrite"],
        account: account
    });

    // Check if file exists in the root of OneDrive
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}`;

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
        });

        if (response.status === 404) {
            console.warn("File not found. MAE System: Initializing new workbook...");
            await createInitialWorkbook(tokenResponse.accessToken);
        } else {
            console.log("MAE System: Workbook verified and ready.");
         //  Run health check even if file exists to ensure customer didn't break it
            await initializeSheetAndTable(tokenResponse.accessToken);
        }
    } catch (error) {
        console.error("Verification Error:", error);
    }
}
//======END FUNCTION verifySpreadSheetExists =============

//=========FUNCTION createInitialWorkbook =============
// Practical: Instead of building via API (brittle), we upload your Master Template.
async function createInitialWorkbook(accessToken) {
    const statusTitle = document.getElementById("current-view-title");
    statusTitle.innerText = "Initializing your custom workshop system... please wait.";

    // 1. Path to your Master Template in your GitHub Repo
    // Assumes the .xlsx is in the root of your project directory
    const MASTER_TEMPLATE_URL = `./${maeSystemConfig.spreadsheetName}`;

    try {
        // 2. Fetch the physical file from your GitHub server
        const response = await fetch(MASTER_TEMPLATE_URL);
        if (!response.ok) throw new Error("Could not find the Master Template on the server.");
        const fileBlob = await response.blob();

        // 3. Upload to Customer's OneDrive Root
        const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(maeSystemConfig.spreadsheetName)}:/content`;
        
        const uploadResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            },
            body: fileBlob
        });

        if (uploadResponse.ok) {
            console.log("MAE System: Master Workbook uploaded successfully.");
            statusTitle.innerText = "System Initialized! Selecting a module to begin.";
            // Now that the file exists, we perform a quick verification check
            await initializeSheetAndTable(accessToken);
        } else {
            const errorData = await uploadResponse.json();
            throw new Error(`Upload failed: ${errorData.error.message}`);
        }

    } catch (error) {
        console.error("Critical Error during initialization:", error);
        statusTitle.innerText = "Setup Error: Please contact MAE Support.";
    }
}

//=======END FUNCTION createInitialWorkbook ==============

//=========FUNCTION initializeSheetAndTable =============
async function initializeSheetAndTable(accessToken) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");

    console.log("MAE System: Running Health Check...");
    
    // Check the first table in config to see if the file is "healthy"
    const firstTableName = maeSystemConfig.worksheets[0].tableName;
    const checkUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${firstTableName}`;

    try {
        const response = await fetch(checkUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        if (response.ok) {
            title.innerText = "System Ready: Select a Category";
            container.innerHTML = `<p style="padding:20px;">Workbook verified. Use the sidebar to manage your ${maeSystemConfig.worksheets.length} workshop modules.</p>`;
        } else {
            // Error handling for the "Bowing Out" strategy
            title.innerText = "System Integrity Alert";
            container.innerHTML = `
                <div style="padding:20px; color: #c0392b;">
                    <p><strong>Warning:</strong> The spreadsheet structure has been modified outside the app.</p>
                    <p>Please ensure the table <b>${firstTableName}</b> has not been renamed or deleted in Excel.</p>
                    <hr>
                    <p>To reset, you may delete the file from OneDrive and refresh this page to re-deploy the master template.</p>
                </div>`;
        }
    } catch (error) {
        console.error("Health check error:", error);
    }
}

//========END FUNCTION initializeSheetAndTable===========

//========== FUNCTION  loadTableData ==================
/**
 * REVISED loadTableData
 * Logic: Fetch from Microsoft Graph
 * UI: Hand off to UI.renderTable
 */
async function loadTableData(tableName) {
   window.currentTable = tableName;
   
   const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
   UI.showLoading(tableName);

   try {
    //Get fresh token
    const tokenResponse = await myMSALObj.acquireTokenSilent({
        scopes: ["Files.ReadWrite"],
        account: account
    });

    //API request to Microsoft Graph
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;

    const response = await fetch(url, {
        headers: {'Authorization' : `Bearer ${tokenResponse.accessToken}`}
    });

    if (!response.ok){
        throw new Error(`Failed to fetch table data: ${response.statusText}`);  
    }

    const data = await response.json();

    // Hand off to UI module
    // Find the specific sheet config to pass along for column filtering
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);

    //Pass 1. The Rows, 2. The Table Name, 3. The Config Blueprint
    UI.renderTable(data.value, tableName, sheetConfig);

    // Draw command bar at the bottom
    UI.renderCommandBar(tableName);

   } catch (error) {
    console.error("MAE System: Error loading table data:", error);
    UI.showError("Error: Could not load data.  Ensure spreadsheet is not open in another tab.");
   }
} 

//=========== END loadTableData ===================

//=========== GLOBAL CLICK LISTENER FUNCTION===========
// 1. GLOBAL CLICK LISTENER (Event Delegation)
// This stays active even when buttons are deleted/recreated
document.getElementById('action-bar-zone').addEventListener('click', (event) => {
    // We check the ID of what was actually clicked
    const btn = event.target.closest('button');
    if (!btn) return;// Exit if they clicked the bar but not the button

    if (btn.id === 'btn-add') {
        // Trigger the Add Item flow
        handleAddClick(window.currentTable); 
    } else if (btn.id === 'btn-edit') {
        handleEditClick(window.currentTable);
    }
    // You can add more 'else if' blocks here later for Edit/Print/Delete
});

// handle ADD Click function
async function handleAddClick(tableName) {
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    
    // Call the UI tool to draw the form
    // We pass a 'callback' function so the UI knows what to do when 'Save' is clicked
    UI.renderEntryForm('add',tableName, sheetConfig, () => {
        submitNewRow(tableName, sheetConfig);
    });
}

// Handle EDIT CLICK function
// app.js - REPLACES the previous handleEditClick
function handleEditClick(tableName) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    // 1. VISUAL FEEDBACK: Add the "is-editing" class to show borders & delete icons
    table.classList.add("is-editing");
    
    // 2. UNLOCK CELLS: Find all cells marked as 'editable' and turn on browser editing
    const cells = table.querySelectorAll(".editable-cell");
    cells.forEach(cell => {
        cell.contentEditable = "true";
    });

    // 3. RUGGED PROTECTION: Click-Outside-To-Save logic
    const handleOutsideClick = (e) => {
        // If the user clicks something that ISN'T the table or the Edit button itself...
        const isClickInsideTable = table.contains(e.target);
        const isClickEditBtn = e.target.closest('#btn-edit');

        if (!isClickInsideTable && !isClickEditBtn) {
            // ...Save the changes and lock the table back up
            processInPlaceTableUpdate(tableName); 
            exitEditMode();
            document.removeEventListener('click', handleOutsideClick);
        }
    };

    // Timeout (100ms) ensures the click that opened the mode doesn't immediately close it
    setTimeout(() => {
        document.addEventListener('click', handleOutsideClick);
    }, 100);
}

// Helper to lock the UI back down
function exitEditMode() {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    table.classList.remove("is-editing");
    const cells = table.querySelectorAll(".editable-cell");
    cells.forEach(cell => cell.contentEditable = "false");
}


//===========FUNCTION submitNewRow====to send data to Microsoft========
// app.js

async function submitNewRow(tableName, sheetConfig) {
    // 1. MAP DATA: Order matches config.js exactly
    const rowData = sheetConfig.columns.map(col => {
        // Handle Auto-ID
        if (col.header === "mae_id") return `MAE-${Date.now()}`;
        
        // Handle Formulas (Excel must calculate these)
        if (col.type === "formula") return null;

        // Find the input field by its cleaned ID
        const fieldId = `field-${col.header.replace(/\s+/g, '')}`;
        const input = document.getElementById(fieldId);

        // RUGGED: Handle Empty Inputs
        // Sending null for empty numbers/dates keeps Excel calculations accurate
        if (!input || input.value === "") {
            return (col.type === "number" || col.type === "date") ? null : "";
        }

        // Handle Numbers & Currency
        if (col.type === "number" || (col.format && col.format.includes("$"))) {
            const num = parseFloat(input.value);
            return isNaN(num) ? null : num;
        }

        // Handle Dates, Dropdowns, and Strings
        // HTML5 date inputs (YYYY-MM-DD) are natively accepted by Excel
        return input.value;
    });

    try {
        // 2. AUTH: Get fresh token
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        // 3. API CALL: Corrected URL path for Table Rows
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows`;
        //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowData] }) // Values must be 2D array
        });

        if (response.ok) {
            console.log(`MAE System: Row successfully added to ${tableName}`);
            alert("Entry Saved Successfully!");
            
            // Clean up: Remove form and refresh table view
            const form = document.getElementById("add-entry-form");
            if (form) form.remove();
            
            loadTableData(tableName); 
        } else {
            const error = await response.json();
            throw new Error(error.error.message || "Unknown API Error");
        }

    } catch (err) {
        console.error("MAE System - Save failed:", err);
        UI.showError(`Failed to save: ${err.message}`);
    }
}


//===========END function to send data to Microsoft


//===========END GLOBAL CLICK LISTENER============

// ======= FUNCTION to scan html table and send updates to OneDrive ==========
//=======  function processInPlaceTableUpdate ========

// app.js - The "Brain" for In-Place Saving

async function processInPlaceTableUpdate(tableName) {
    const table = document.getElementById("main-data-table");
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    const rows = table.querySelectorAll("tbody tr");
    
    // We only want to send rows that have actually been modified to save on API calls
    const updates = [];

    rows.forEach(tr => {
        const rowIndex = tr.getAttribute("data-row-index");
        const rowValues = [];

        // Build the row array based on config order
        sheetConfig.columns.forEach((col, index) => {
            // Find the cell in this row that matches the config column index
            const cell = tr.querySelector(`td[data-col-index="${index}"]`);
            
            if (col.type === "formula") {
                rowValues.push(null); // Excel will recalculate formulas
            } else if (cell) {
                let val = cell.innerText.trim();
                // Rugged: Convert to numbers where required so Excel math doesn't break
                if (col.type === "number") {
                    val = val === "" ? null : parseFloat(val.replace(/[^0-9.-]+/g,""));
                }
                rowValues.push(val);
            } else {
                rowValues.push(null);
            }
        });
        
        updates.push({ index: rowIndex, values: [rowValues] });
    });

    // Send to Microsoft Graph
    try {
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        // Professional Approach: Update rows one by one or in a batch if supported
        for (const update of updates) {
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows/itemAt(index=${update.index})`;
            //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
            await fetch(url, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${tokenResponse.accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: update.values })
            });
        }
        console.log("MAE System: All changes synced to OneDrive.");
    } catch (err) {
        console.error("Sync Error:", err);
        UI.showError("Failed to sync changes. Check internet connection.");
    }
}

// =====  END  function processInPlaceTableUpdate   ==========

//======== FUNCTION delete Excel Row ===========
// app.js - Rugged Row Deletion

async function deleteExcelRow(tableName, rowIndex) {
    try {
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
         //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
        const response = await fetch(url, {
            method: 'DELETE',
            headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
        });

        if (response.ok) {
            alert("Row deleted successfully.");
            loadTableData(tableName); // Refresh the UI to reflect the removal
        } else {
            throw new Error("Failed to delete row.");
        }
    } catch (err) {
        console.error("Delete Error:", err);
        alert("Could not delete row: " + err.message);
    }
}

//====== END delete Excel Row ============
