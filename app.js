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
    UI.renderCommandBar("");

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

// ========== DATE CONVERSION HELPER ===============
/**
 * Converts an Excel serial date number to a formatted MM/DD/YYYY string.
 * @param {number} serial - The Excel serial date (e.g., 44562).
 * @returns {string} - The formatted date string.
 */
function excelSerialToDate(serial) {
    if (!serial || isNaN(serial)) return serial; // Return as-is if not a valid number
    
    // Formula: (Serial - 25569) * milliseconds in a day
    const jsDate = new Date(Math.round((serial - 25569) * 86400 * 1000));
    
    const mm = String(jsDate.getMonth() + 1).padStart(2, '0');
    const dd = String(jsDate.getDate()).padStart(2, '0');
    const yyyy = jsDate.getFullYear();

    return `${mm}/${dd}/${yyyy}`;
}


// ======= END DATE CONVERSION HELPER =========

//========== FUNCTION  loadTableData ==================
/**
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

    if (!response.ok) throw new Error(`Graph API error: ${response.status}`);
    const data = await response.json();

    // Process rows to format dates before rendering
    const formattedRows = data.value.map(row => {
    return row.values[0].map((cellValue, index) => {
        const colDef = sheetConfig.columns[index];
        // Check if this column is marked as a date in your manifest
        if (colDef && colDef.type === 'date') {
            return excelSerialToDate(cellValue);
        }
        return cellValue;
    });
});

    // Hand off cleaned data to to UI module
    //Pass 1. The Rows, 2. The Table Name, 3. The Config Blueprint
    UI.renderTable(formattedRows, tableName, sheetConfig);

    // Draw command bar at the bottom
    UI.renderCommandBar(tableName);

   } catch (error) {
    console.error("MAE System: Error loading table data:", error);
    UI.showError("Error: Could not load data.  Ensure spreadsheet is closed in Excel");
   }
} 

//=========== END loadTableData ===================

//=========== GLOBAL CLICK LISTENER FUNCTION===========
// 1. GLOBAL CLICK LISTENER (Event Delegation)
// This stays active even when buttons are deleted/recreated
// app.js - Refined Global Listener
document.getElementById('action-bar-zone').addEventListener('click', (event) => {
    const btn = event.target.closest('button');
    if (!btn) return;

    // Use the global window object for consistency
    const config = window.maeSystemConfig;
    const currentTable = window.currentTable;

    if (btn.id === 'btn-add') {
        handleAddClick(currentTable); 
    } 
    else if (btn.id === 'btn-edit') {
        handleEditClick(currentTable);
    } 
    else if (btn.id === 'btn-print') {
        const sheetConfig = config.worksheets.find(s => s.tableName === currentTable);
        UI.printTable(currentTable, sheetConfig);
    } 
    else if (btn.id === 'btn-manual-print') {
        const sheetConfig = config.worksheets.find(s => s.tableName === currentTable);
        UI.printManualLog(currentTable, sheetConfig);
    } 
    else if (btn.id === 'btn-inventory-update') {
        handleQuickUpdate(currentTable);
    }
    // You can add more 'else if' blocks here later for Edit/Print/Delete
});

//============= handle ADD Click function ===================
async function handleAddClick(tableName) {
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    
    UI.renderEntryForm('add', tableName, sheetConfig, async () => {
        // 1. Submit the data
        const success = await submitNewRow(tableName, sheetConfig);
        
        // 2. If successful, clear the inputs so the form stays open for more items
        if (success) {
            const formContainer = document.getElementById("entry-form");
            if (formContainer) {
                const inputs = formContainer.querySelectorAll("input, select");
                inputs.forEach(input => {
                    input.value = ""; // Clear the text
                });

                // RUGGED: Auto-focus the first visible input (skips hidden IDs)
                const firstVisible = formContainer.querySelector("input:not([type='hidden']), select");
                if (firstVisible) firstVisible.focus();

                console.log("MAE System: Entry saved. Form cleared for next item.");
            }
        }
    });
}


//========== Handle EDIT CLICK function ================

function handleEditClick(tableName) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    table.classList.add("is-editing");
    
    const cells = table.querySelectorAll(".editable-cell");

    cells.forEach(cell => {
        const colIdx = parseInt(cell.getAttribute('data-col-index'));
        const colDef = sheetConfig.columns[colIdx];

        // --- RUGGED RULE: NO contentEditable for Dropdowns ---
        if (colDef.type === "dropdown") {
            cell.contentEditable = "false"; 
            cell.setAttribute('tabindex', '0'); 
            cell.classList.add("dropdown-edit-zone");

            // Event listener to inject the select menu
            const startDropdownEdit = (e) => {
                e.stopPropagation();
                if (cell.querySelector('select')) return; // Already editing

                const currentVal = cell.innerText.trim();
                let selectHtml = `<select class="edit-dropdown" style="width:100%; height:100%; border:none; background:#fffde7; font:inherit; cursor:pointer;">`;
                
                colDef.options.forEach(opt => {
                    selectHtml += `<option value="${opt}" ${opt === currentVal ? 'selected' : ''}>${opt}</option>`;
                });
                selectHtml += `</select>`;

                cell.innerHTML = selectHtml;
                const select = cell.querySelector('select');
                select.focus();

                // Logic to finish editing
                const finishEdit = () => {
                    cell.innerText = select.value; // Store value as text for processInPlaceTableUpdate
                };

                select.onchange = finishEdit;
                select.onblur = finishEdit;
                // Allow "Enter" to finish edit
                select.onkeydown = (k) => { if(k.key === 'Enter') finishEdit(); };
            };

            cell.onclick = startDropdownEdit;
            // Also trigger on 'Enter' key for keyboard accessibility (Rugged)
            cell.onkeydown = (k) => { if(k.key === 'Enter') startDropdownEdit(k); };
        } 
        // --- STANDARD TEXT/NUMBERS ---
        else {
            cell.contentEditable = "true";
            cell.setAttribute('tabindex', '0');

            // Quick Update: Arrow keys
            if (colDef.header === "Quantity" || colDef.header === "Current Stock") {
                cell.onkeydown = (e) => {
                    if (e.key === "ArrowUp" || e.key === "ArrowDown") {
                        e.preventDefault();
                        let val = parseInt(cell.innerText) || 0;
                        cell.innerText = (e.key === "ArrowUp") ? val + 1 : Math.max(0, val - 1);
                    }
                };
            }
        }

        cell.onmousedown = (e) => e.stopPropagation();
    });

    // Global Click-Outside logic
    const handleOutsideClick = (e) => {
    // Guards
    if (e.target.closest('.delete-row-btn')) return;
    const isStartBtn = e.target.id === 'btn-inventory-update';

    if (table && !table.contains(e.target) && !isStartBtn) {
        // STEP A: Send the data to Microsoft in the background
        processInPlaceTableUpdate(tableName); 

        // STEP B: Instantly move the span text into the cell (Instant Persistence)
        exitEditMode();

        // STEP C: Remove the listener
        document.removeEventListener('mousedown', handleOutsideClick);
        
        // CRITICAL: We do NOT call loadTableData(). 
        // By NOT calling it, the app never asks Excel for the "old" data.
        console.log("MAE System: UI Updated locally. Syncing to OneDrive in background...");
    }
};


    setTimeout(() => {
        document.addEventListener('mousedown', handleOutsideClick);
    }, 150);
}


// ===========   FUNCTION Exit Edit Mode ===========
function exitEditMode() {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    table.classList.remove("is-editing", "is-quick-updating");
    
    const cells = table.querySelectorAll("td");
    cells.forEach(cell => {
        // 1. Find our special Qty span
        const qtySpan = cell.querySelector('.qty-value');
        
        if (qtySpan) {
            // 2. CAPTURE the user's manual entry
            const newValue = qtySpan.innerText.trim();
            
            // 3. LOCK IT IN: Replace the HTML with just the plain text
            // This prevents the "revert" because the span is gone, 
            // but the number remains.
            cell.innerText = newValue; 
        }

        // 4. Remove all temporary "Edit Mode" styling
        cell.contentEditable = "false";
        cell.style.opacity = "";
        cell.style.backgroundColor = "";
        cell.style.pointerEvents = "";
        cell.classList.remove("quick-edit-focus");
    });
}


//=========  END Exit Edit Mode ===============


//===========FUNCTION submitNewRow====to send data to Microsoft========

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
            //const form = document.getElementById("add-entry-form");
           // if (form) form.remove();
            
            loadTableData(tableName); 
            return true;
        } else {
            const error = await response.json();
            throw new Error(error.error.message || "Unknown API Error");
        }

    } catch (err) {
        console.error("MAE System - Save failed:", err);
        UI.showError(`Failed to save: ${err.message}`);
    }

    return false;
}


//===========END function to send data to Microsoft
//===========END GLOBAL CLICK LISTENER============

// ======= FUNCTION to scan html table and send updates to OneDrive ==========
//=======  function processInPlaceTableUpdate ========

// app.js - The "Brain" for In-Place Saving

async function processInPlaceTableUpdate(tableName) {
    const table = document.getElementById("main-data-table");
    if (!table) return;
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
                let val =   cell.querySelector('.qty-value') ?
                            cell.querySelector('.qty-value').innerText.trim() :
                            cell.innerText.trim();
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

// ========= Standalone Saving Single Row Update =======
// app.js - Standalone Optimization Function
async function saveSingleRowUpdate(tableName, rowIndex, rowValues) {
    try {
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
        //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
        const response = await fetch(url, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowValues] }) // Values must be in a nested array
        });

        if (response.ok) {
            console.log(`MAE System: Row ${rowIndex} updated successfully.`);
        } else {
            const error = await response.json();
            throw new Error(error.error.message);
        }
    } catch (err) {
        console.error("Single Row Sync Error:", err);
    }
}

// ======= END Saving Single Row Update ===========

//======== FUNCTION delete Excel Row ==========

// 1. The Wrapper (The button calls this)
async function requestDelete(rowIndex) {
    // RUGGED: Always warn the user before a destructive action
    const confirmed = confirm(`WARNING: Are you sure you want to delete row ${rowIndex + 1}? This cannot be undone.`);
    
    if (confirmed) {
        // If they say yes, call your API function
        await deleteExcelRow(window.currentTable, rowIndex);
    }
}

// 2. EXPOSE TO WINDOW: This fixes the ReferenceError
window.requestDelete = requestDelete;


// API logic to delete an Excel Row
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

//======= FUNCTION handleQuickUpdate ================
function handleQuickUpdate(tableName) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    table.classList.add("is-editing", "is-quick-updating");
    
    const cells = table.querySelectorAll("td");

    cells.forEach(cell => {
        const colIdx = parseInt(cell.getAttribute('data-col-index'));
        if (isNaN(colIdx)) return; 

        const colDef = sheetConfig.columns[colIdx];
        const isInventoryCol = colDef.header === "Quantity" || colDef.header === "Current Stock";

        if (isInventoryCol) {
            const currentVal = cell.innerText.trim();
            cell.classList.add("quick-edit-focus");
            
            // Inject Visible UI: Value + Up/Down Buttons
            cell.innerHTML = `
                <div class="qty-editor">
                    <span class="qty-value" contenteditable="true" tabindex="0">${currentVal}</span>
                    <div class="qty-controls">
                        <button class="qty-up">▲</button>
                        <button class="qty-down">▼</button>
                    </div>
                </div>
            `;

            const valSpan = cell.querySelector('.qty-value');
            // NEW: Instant save when focus is lost (blur)
            valSpan.onblur = () => {
                // 1. Immediately extract the value
                const newVal = valSpan.innerText.trim();
    
                // 2. Trigger the sync to OneDrive without waiting for a global click
                processInPlaceTableUpdate(tableName);
    
                // 3. (Optional) Provide a small visual 'success' flash
                valSpan.style.backgroundColor = "#d4edda"; // Light green
                setTimeout(() => valSpan.style.backgroundColor = "transparent", 500);
            };

            const adjust = async (amt) => {
                let val = parseInt(valSpan.innerText) || 0;
                const newQty = Math.max(0, val + amt);
                valSpan.innerText = newQty;

                // RUGGED: Get all values for this row to satisfy the Excel API requirement
                const rowValues = extractCurrentRowValues(cell.parentElement); 
                const rowIndex = cell.parentElement.getAttribute("data-row-index");

                // TRIGGER INSTANT SAVE
                saveSingleRowUpdate(tableName, rowIndex, rowValues);

            };

            // Click Logic for Arrows
            cell.querySelector('.qty-up').onclick = (e) => { e.stopPropagation(); adjust(1); };
            cell.querySelector('.qty-down').onclick = (e) => { e.stopPropagation(); adjust(-1); };

            // Keyboard Arrow Support
            valSpan.onkeydown = (e) => {
                if (e.key === "ArrowUp") { e.preventDefault(); adjust(1); }
                if (e.key === "ArrowDown") { e.preventDefault(); adjust(-1); }
                if (e.key === "Enter") { e.preventDefault(); valSpan.blur(); }
            };

            // Focus the number immediately for rapid entry
            setTimeout(() => valSpan.focus(), 50);

        } else {
            cell.contentEditable = "false";
            cell.style.opacity = "0.4";
            cell.style.backgroundColor = "#f9f9f9";
            cell.style.pointerEvents = "none"; 
        }
    });

    const handleOutsideClick = (e) => {
        // ADDED: Guard to ignore clicks on the Quick Update button itself
        const isBtn = e.target.id === 'btn-inventory-update';
        
        if (table && !table.contains(e.target) && !isBtn) {
            processInPlaceTableUpdate(tableName); 
            table.classList.remove("is-quick-updating", "is-editing");
            exitEditMode();
            document.removeEventListener('mousedown', handleOutsideClick);
            //loadTableData(tableName); 
        }
    };

    setTimeout(() => document.addEventListener('mousedown', handleOutsideClick), 150);
}


// ========END handleQuickUpdate ============

// ==========  extract Current Row Values =======
function extractCurrentRowValues(trElement) {
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === window.currentTable);
    return sheetConfig.columns.map((col, index) => {
        const cell = trElement.querySelector(`td[data-col-index="${index}"]`);
        if (col.type === "formula") return null;
        
        // Use your existing logic to find the value (span vs plain text)
        return cell.querySelector('.qty-value') ? 
               cell.querySelector('.qty-value').innerText.trim() : 
               cell.innerText.trim();
    });
}

//======= extract Current Row Values ==========

window.handleEditClick = handleEditClick;
window.handleQuickUpdate = handleQuickUpdate;
window.handleAddClick = handleAddClick;
window.requestDelete = requestDelete;
