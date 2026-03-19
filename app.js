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
    // 1. FILTER: only show worksheets with active:true
    const activeWorksheets = maeSystemConfig.worksheets.filter(sheet => sheet.active !== false);

    // 2. UI handles the creation of buttons
    UI.renderMenu(activeWorksheets, (tableName) => {
        
        // RUGGED RESET: Call the centralized UI cleanup
        // This handles: Dropdowns, Edit Mode Classes, and the Add Item Form
        UI.exitEditMode(); 

        // 3. Load the new module
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
     // RUGGED: We update the values but keep the row object so UI.js 
    // can still find the row index and metadata.
    const formattedRows = data.value.map(rowObj => {
        // Map the inner values array (Graph API returns values as a 2D array [[]])
        const cleanValues = rowObj.values[0].map((cellValue, index) => {
            const colDef = sheetConfig.columns[index];
            if (colDef && colDef.type === 'date') {
                return excelSerialToDate(cellValue);
            }
            return cellValue;
        });

        // Return the original object but with the formatted values
        return {
            ...rowObj,
            values: [cleanValues] 
        };
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
// ========== END ADD Click Function ====================

// ============ GLOBAL "CLICK OFF" HANDLER for all Edit Modes ==========
// Centralized "Click-Off" handler for all edit modes
async function globalClickOffHandler(e) {
    const table = document.getElementById("main-data-table");
    const title = document.getElementById("current-view-title");
    if (!table || !title) return;

    // 1. IDENTIFY SAFE ZONES
    const isInsideTable = table.contains(e.target);
    const isCommandBtn = e.target.closest('.action-btn');
    const isDeleteBtn = e.target.closest('.delete-row-btn');

    // 2. TRIGGER SYNC ONLY ON BACKGROUND CLICK
    if (!isInsideTable && !isCommandBtn && !isDeleteBtn) {
        // DETACH IMMEDIATELY: Prevents double-syncing if they click twice
        document.removeEventListener('mousedown', globalClickOffHandler);
        
        console.log("MAE System: Outside click detected. Syncing and Closing.");
        
        // 3. CAPTURE ORIGINAL STATE
        const originalTitle = title.innerText;

        // 4. SET VISUAL SAVING STATE
        title.innerText = "💾 Saving changes to OneDrive... Please wait.";
        title.classList.add("is-syncing");
        table.classList.add("saving-active");
        table.style.opacity = "0.5";
        table.style.pointerEvents = "none";

        try {
            // 5. ATTEMPT SYNC
            await processInPlaceTableUpdate(window.currentTable); 
            console.log("MAE System: Sync successful.");
        } catch (err) {
            // 6. ERROR HANDLING
            console.error("MAE System: Sync failed, forcing UI reset.", err);
        } finally {
            // 7. THE SAFETY RESET: This runs NO MATTER WHAT (success or fail)
            title.innerText = originalTitle;
            title.classList.remove("is-syncing");
            
            // Force styles back to normal in case UI.exitEditMode misses them
            table.style.opacity = "1";
            table.style.pointerEvents = "auto";
            table.classList.remove("saving-active");

            UI.exitEditMode(); 
        }
    }
}
//=======END GLOBAL "CLICK OFF" HANDLER ===========

//========== Handle EDIT CLICK function ================
function handleEditClick(tableName) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    window.currentTable = tableName; // update global state
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    table.classList.add("is-editing");
    
    const cells = table.querySelectorAll(".editable-cell");

    cells.forEach(cell => {
        const colIdx = parseInt(cell.getAttribute('data-col-index'));
        const colDef = sheetConfig.columns[colIdx];

        // --- RUGGED RULE: Dropdowns ---

        if (colDef.type === "dropdown") {
            cell.contentEditable = "false"; 
            cell.setAttribute('tabindex', '0'); 
            cell.classList.add("dropdown-edit-zone");

            const startDropdownEdit = (e) => {
                e.stopPropagation();
                if (cell.querySelector('select')) return;

                const currentVal = cell.innerText.trim();
                let selectHtml = `<select class="edit-dropdown" style="width:100%; height:100%; border:none; background:#fffde7; font:inherit; cursor:pointer;">`;
                
                colDef.options.forEach(opt => {
                    selectHtml += `<option value="${opt}" ${opt === currentVal ? 'selected' : ''}>${opt}</option>`;
                });
                selectHtml += `</select>`;

                cell.innerHTML = selectHtml;
                const select = cell.querySelector('select');
                select.focus();
       
                const finishEdit = () => {
                   // 1. Explicitly grab the selected TEXT, not the entire HTML block
                    const selectedText = select.options[select.selectedIndex].text;
    
                    // 2. NUCLEAR RESET: Wipe the cell and set it to just the text
                    cell.innerHTML = selectedText; 
    
                    // 3. Clean up the cell's temporary edit classes
                    cell.classList.remove("dropdown-edit-zone");  
                };
       
                select.onchange = finishEdit;
                select.onblur = finishEdit;
                select.onkeydown = (k) => { if(k.key === 'Enter') finishEdit(); };
            };

            cell.onclick = startDropdownEdit;
            cell.onkeydown = (k) => { if(k.key === 'Enter') startDropdownEdit(k); };
        } 

        // --- NUMBERS (Integer & Currency) ---
        else if (colDef.type === "number") {
            const isCurrency = colDef.format && colDef.format.includes("$");
            const currentVal = cell.innerText.replace(/[^0-9.-]+/g, "") || 0;

            // Inject the native input (This gives you the arrows)
            cell.contentEditable = "false"; 
            cell.innerHTML = `<input type="number" class="edit-number-input" value="${currentVal}" step="${isCurrency ? '0.01' : '1'}" min="0">`;
    
            const input = cell.querySelector('input');
            input.focus();

            // RE-APPLY YOUR PROTECTION LOGIC
            input.onkeydown = (e) => {
            // Block scientific 'e'
                if (e.key.toLowerCase() === "e") e.preventDefault();

                // Block decimals for Integers (Qty/Stock)
                if (!isCurrency && (e.key === "." || e.key === ",")) {
                    e.preventDefault();
            }
        };

            // CLEANUP ON BLUR
            input.onblur = () => {
                let val = parseFloat(input.value);
                if (isNaN(val)) val = 0;

                // Standardize the text in the cell for the Sync function
                cell.innerText = isCurrency ? val.toFixed(2) : Math.floor(val).toString();
        };
    }




 /*       else if (colDef.type === "number") {
            const isCurrency = colDef.format && colDef.format.includes("$");
            const isInteger = !isCurrency; // Qty, Stock, Reorder Point

                if (isInteger){

                }





            cell.contentEditable = "true";
            cell.setAttribute('tabindex', '0');
            const isCurrency = colDef.format && colDef.format.includes("$");
            const isQtyField = colDef.header === "Quantity" || colDef.header === "Current Stock" || colDef.header === "Reorder Point";

            cell.onkeydown = (e) => {
            // 1. Block scientific 'e'
            if (e.key.toLowerCase() === "e") e.preventDefault();

            // 2. Block decimals for Integers
            if (!isCurrency && (e.key === "." || e.key === ",")) {
                e.preventDefault();
            }

            // 3. RUGGED ARROW LOGIC
            if (isQtyField && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
             e.preventDefault();
            
                // Clean the text of any whitespace/hidden chars before parsing
                let currentText = cell.innerText.replace(/\s/g, '');
                let val = parseInt(currentText) || 0;
            
            const newVal = (e.key === "ArrowUp") ? val + 1 : Math.max(0, val - 1);
            
            // Update the UI
            cell.innerText = newVal;

            // IMPORTANT: Move cursor to the end so they can keep typing if they want
            const range = document.createRange();
            const sel = window.getSelection();
            range.selectNodeContents(cell);
            range.collapse(false);
            sel.removeAllRanges();
            sel.addRange(range);
        }
    };

    cell.onblur = () => {
        // Remove everything except numbers and decimals
        let raw = cell.innerText.replace(/[^0-9.-]+/g, "");
        let num = parseFloat(raw);
        
        if (isNaN(num)) {
            cell.innerText = "0";
        } else {
            // Standardize format: Currency gets decimals, Qty gets rounded down
            cell.innerText = isCurrency ? num.toFixed(2) : Math.floor(num).toString();
        }
    };
}*/
        // --- STANDARD TEXT ---
        else {
            cell.contentEditable = "true";
            cell.setAttribute('tabindex', '0');
        }

        cell.onmousedown = (e) => e.stopPropagation();
    });

    // ATTACH THE CENTRAL LISTENER (Replace your internal handleOutsideClick)
    setTimeout(() => {
        document.addEventListener('mousedown', globalClickOffHandler);
    }, 150);
}

// ====== END handleEditClick function =================


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
            if (col.type === "formula") {
                rowValues.push(null);
                return; // Move to next column
            }

            const cell = tr.querySelector(`td[data-col-index="${index}"]`);

            if (!cell) {
                const isNumeric = col.type === "number" || col.type === "date";
                rowValues.push(isNumeric ? null : "");
                return;
            }


             // 1. RUGGED SCRUB: Prioritize UI elements over raw text
                const select = cell.querySelector('select');
                const qtySpan = cell.querySelector('.qty-value');
            
                let val = "";
                if (select) {
                    val = select.value;
                } else if (qtySpan) {
                    val = qtySpan.innerText.trim();
                } else {
                    val = cell.innerText.trim();
                }

                // 2. TYPE ENFORCEMENT (Step 4: Currency vs Integer)
                if (col.type === "number") {
                    const isCurrency = col.format && col.format.includes("$");
                    // Strip all non-numeric characters except decimal/minus
                    let cleanNum = parseFloat(val.replace(/[^0-9.-]+/g,""));
                
                    if (isNaN(cleanNum)) {
                    val = 0; // Prevent Excel from rejecting a "blank" string
                    } else {
                    // Force whole numbers for Qty/Stock, 2-decimals for Currency
                    val = isCurrency ? parseFloat(cleanNum.toFixed(2)) : Math.floor(cleanNum);
                    }
                }
                rowValues.push(val);
            
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
            // NEW: Add the Worksheet segment between the Filename and the Table
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${update.index})`;
            //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows/itemAt(index=${update.index})`;
           
            
            const response = await fetch(url, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${tokenResponse.accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: update.values })
            });

            if (!response.ok){
                const errorBody = await response.json();
                console.error("Microsoft Graph Error:", errorBody);
            }
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

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${update.index})`;
        //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
       
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
async function handleQuickUpdate(tableName) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    window.currentTable = tableName; //update global state
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    
    // 1. Set Visual States
    table.classList.add("is-editing", "is-quick-updating");

    const cells = table.querySelectorAll("td");
    cells.forEach(cell => {
        const colIdxAttr = cell.getAttribute('data-col-index');
        
        // RUGGED: Skip the 'Delete' column or any cell without an index
        if (colIdxAttr === null) return; 

        const colIdx = parseInt(colIdxAttr);
        const colDef = sheetConfig.columns[colIdx];

        // Safety check to prevent the 'undefined' crash
        if (!colDef) return; 

        const isQtyField = colDef.header === "Quantity" || colDef.header === "Current Stock";
        
        if (isQtyField) {
            cell.classList.add("quick-edit-focus");
            cell.contentEditable = "true";
            // Note: Your existing Up/Down button injection logic should be called here
        } else {
            cell.contentEditable = "false";
            cell.style.backgroundColor = "#f9f9f9";
            cell.style.color = "#999";
        }
    });
    
    // ATTACH THE CENTRAL LISTENER
    setTimeout(() => {
        document.addEventListener('mousedown', globalClickOffHandler);
    }, 150);
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
