window.maeLocations =["TBD"]; // Global cache default for intake workflow
import { maeSystemConfig } from './config.js'
import { UI} from './ui.js';
import { Labels } from './labels.js';
import { Dashboard } from './dashboard.js';
import { myMSALObj } from './auth.js';
const fileName = maeSystemConfig.spreadsheetName;
window.currentTable = "";


// =========== STARTUP LOGIC ============
//Initializes the authentication flow for app. Handles the moment page
//first loads, specifically checking if user is returning from a login 
//attempt or has an existing session (ie, clicked refresh)  

window.account = null;
window.isEditing = false; // Initialize the global state

async function startup() {
    try {
        // Initialize the PublicClientApplication
        //  MSAL V2 uses 'msal.PublicClientApplication'
        //myMSALObj = new window.msal.PublicClientApplication(msalConfig);

        const response = await myMSALObj.handleRedirectPromise();
    
        if (response) {
            window.account = response.account;
            console.log("Login successful via redirect. Account:", window.account.username);
        } else {
            const accounts = myMSALObj.getAllAccounts();
            if (accounts.length > 0) account = accounts[0];
        }

        if (window.account) {
            // SCENARIO 1: USER IS LOGGED IN
            updateUIForLoggedInUser(window.account); 

            // --- INDUSTRIAL SCANNER INTEGRATION ---
            // Purged old URL/QR lookup logic.
            // Initializing the background listener for the Inateck-75S.
            console.log("MAE System: Initializing Industrial HID Listener...");
            Labels.initHIDScanner((scannedId) => {
                // When a scan is detected, it triggers the universal search
                handleUniversalLookup(scannedId);
            });

        } else {
            // SCENARIO 2: USER IS NOT LOGGED IN
            const authButton = document.getElementById("auth-btn");
            if (authButton) {
                authButton.addEventListener("click", signIn);
                console.log("MAE System: Auth button ready.");
            }
        }
    } catch (error) {
        console.error("Error during MSAL startup:", error);
    }
}

//========END STARTUP LOGIC ===========


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

// ===== Centralized Authentication Token Helper =============
async function getGraphToken() {
    try {
        // Find the active account. Fallback to retrieving all accounts if global state is missing
        let activeAccount = window.account;
        if (!activeAccount) {
            const accounts = myMSALObj.getAllAccounts();
            if (accounts.length > 0) activeAccount = accounts[0];
        }

        if (!activeAccount) {
            throw new Error("No active user account found. Please sign in.");
        }

        // Fetch the token silently
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: activeAccount
        });

        // Return just the string token
        return tokenResponse.accessToken;

    } catch (error) {
        console.error("MAE System Fail: Could not acquire token silently.", error);
        UI.showError("Session expired. Please reconnect to Office 365.");
        throw error;
    }
}

//==== END Centralized Authentication Token Helper =============

// ======== FUNCTION TO UPDATE UI BASED ON LOGIN STATUS ========
// the startup() function calls updateUIForLoggedInUser() if successful 'login'
// changes text on button and triggers loadDynamicMenu() function  
function updateUIForLoggedInUser(userAccount) {

    UI.setConnected(userAccount.username, signOut);
    loadDynamicMenu();

    loadTableData("Master_Dashboard");
}

//=====END UPDATE UI BASED ON LOGIN STATUS ========

//========SIGN-OUT FUNCTION ===========
async function signOut() {
    console.log("Starting sign-out process via redirect...");
    
    if (!window.account) {
        resetUI();
        return;
    }

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(window.account.username),
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
        window.account = null;
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

    await refreshLocationCache(); // fetch "control tower" date (locations)



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

//===== NEW WORKER FUNCTION: refresh location cache; specifically pulls you Location List for the dropdowns
// Worker: Turns [["Value1", "Value2"]] into { "mae_id": "Value1", "Location_ID": "Value2" }
function mapRowToHeaders(rowValues, sheetConfig) {
    const data = {};
    sheetConfig.columns.forEach((col, index) => {
        data[col.header] = rowValues[index];
    });
    return data;
}

async function refreshLocationCache() {
    try {
        const locationConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        const data = await Dashboard.getFullTableData("Location");

        if (data && data.length > 0) {
            // Find the index dynamically based on the header name in config
            const locIdx = locationConfig.columns.findIndex(c => c.header === "Location_ID");

            const list = data.map(row => {
                const rowCells = row.values[0]; 
                return rowCells[locIdx];
            });

            // Clean the list: remove nulls/duplicates and keep "TBD" at the top
            window.maeLocations = ["TBD", ...new Set(list.filter(i => i && i !== "TBD"))];
            console.log("MAE System: Location cache refreshed using Header Mapping.");
        }
    } catch (e) {
        console.warn("MAE System: Could not find 'Location_ID' column. Using default 'TBD'.", e);
        window.maeLocations = ["TBD"];
    }
}




//==== END WORKER FUNCTION======

// ======= FUNCTION verifySpreadSheetExists =============
async function verifySpreadsheetExists(){
    // Logic here to check if maeSystemConfig.spreadsheetName exists
    // If 404: Call a function to CREATE the workbook using the config
    // If 200: All good, ready to work.
    try {
        // 🌟 NEW CLEAN CALL
        const token = await window.getGraphToken();
        // Check if file exists in the root of OneDrive
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}`;
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

         if (response.status === 404) {
            console.warn("File not found. MAE System: Initializing new workbook...");
            // Pass the clean token directly to the next function
            await createInitialWorkbook(token);
        } else {
            console.log("MAE System: Workbook verified and ready.");
            // Pass the clean token directly to the next function
            await initializeSheetAndTable(token);
        }
    } catch (error) {
        console.error("Verification Error:", error);
    }
}

    
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
    console.log("MAE System: Running Health Check..."); 
    // Check the first table in config to see if the file is "healthy"
    const firstTableName = maeSystemConfig.worksheets[0].tableName;
    const checkUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${firstTableName}`;

    try {
        const response = await fetch(checkUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        if (response.ok) {
            console.log("MAE System: Health Check Passed. Loading Landing Page...");
            const homeBtn = document.querySelector('.home-btn');
            if (homeBtn) {
                document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
                homeBtn.classList.add('active');
            }
  
            loadTableData("Master_Dashboard");
            
        } else {
            // Error handling for the "Bowing Out" strategy
            UI.setHealthStatus(false, firstTableName);
        }
    } catch (error) {
        console.error("Health check error:", error);
        UI.showError("Health check failed.  Check Internet connection.");
    }
}

//========END FUNCTION initializeSheetAndTable===========

// ========== DATE CONVERSION HELPER ===============
/**
 * HELPER FUNCTION which Converts an Excel serial date number to a formatted MM/DD/YYYY string.
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
async function loadTableData(tableName, filterType = null) {
   window.currentTable = tableName;
   const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
   UI.showLoading(tableName);
   try {
    const token = await window.getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
    const response = await fetch(url, {
        headers: {'Authorization' : `Bearer ${token}`}
    });

    if (!response.ok) throw new Error(`Graph API error: ${response.status}`);
    const data = await response.json();

    let displayTitle = `View: ${sheetConfig.tabName}`;

    //==== DASHBOARD BRIDGE ===============
    if(tableName === "Master_Dashboard"){
        // Did Graph API return any rows?
        const hasData = data.value && data.value.length>0;

        //===first data row (under header) data.value[0] is index 0
        const summaryValues = (hasData && data.value[0].values[0]) ? data.value[0].values[0] : [0,0,0,0,0,0,0];

        // RUGGED SIDEBAR RESET: Ensure "Home" button is highlighted
        document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
        const homeBtn = document.querySelector('.home-btn');
        if (homeBtn) homeBtn.classList.add('active');


        UI.renderDashboard(summaryValues, sheetConfig);
        UI.renderCommandBar(tableName);
        return; // EXIT HERE: stops rest of function from running
    }

    // Process rows to format dates before rendering
     //  We update the values but keep the row object so UI.js 
    // can still find the row index and metadata.
    let formattedRows = data.value.map(rowObj => {
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

    //=Apply smart filters
    if (filterType){
        formattedRows = applyDashboardFilters(tableName, formattedRows,filterType);
    }
    //let displayTitle = null;
    if (filterType === 'resell-active'){
        displayTitle = "RESELL INVENTORY: Work In-Progress, Complete and For Sale";
    } else if (filterType === 'low-stock'){
        displayTitle = "Shop Consumables Low Stock";
    } else if (filterType === 'needs-repair'){
        // Map the technical table name to your specific "Operational Issues" title
        const repairTitles = {
            'Shop_Machinery': "Shop Machinery: Operational Issues",
            'Shop_Power_Tools': "Shop Power Tools: Operational Issues",
            'Shop_Hand_Tools': "Shop Hand Tools: Operational Issues"
        };

        const baseTitle = repairTitles[tableName] || "Equipment With Operational Issues";

        // Add the "Back to Dashboard" button logic also used in Overhead
        displayTitle = `
        <div style="display: flex; align-items: center; gap: 15px;">
            <button class="action-btn" 
                    style="padding: 5px 12px; font-size: 0.8rem; background: #7f8c8d;" 
                    onclick="loadTableData('Master_Dashboard')">
                ← Back to Dashboard
            </button>
            <span>${baseTitle}</span>
        </div>`;        
    }
    // Shop Overhead title logic
      else if (tableName === 'Shop_Overhead' && filterType) {
        const titleMap = {
            'due-7': "Bills Due In The Next 7 Days",
            'due-30': "Bills Due In The Next 30 Days",
            'due-90': "Bills Due In The Next 90 Days",
            'due-180': "Bills Due In The Next 180 Days"
        };
        // If the filterType exists in our map, use that title
        if (titleMap[filterType]) {
            // RUGGED NAVIGATION: Injects a compact back button next to the custom title
            displayTitle = `
                <div style="display: flex; align-items: center; gap: 15px;">
                    <button class="action-btn" 
                            style="padding: 5px 12px; font-size: 0.8rem; background: #7f8c8d;" 
                            onclick="loadTableData('Master_Dashboard')">
                        ← Back to All Bills
                    </button>
                    <span>${titleMap[filterType]}</span>
                </div>`;       
        }
    }
    //========== END Shop Overhead TItle logic ============

    //========== SHOP MAINTENANCE title logic ==============
    else if (tableName === 'Maintenance_Log' && filterType?.startsWith('maint-')) {
    const maintTitles = {
        'maint-7': "Maintenance Due In Next 7 Days",
        'maint-30': "Maintenance Due In Next 30 Days",
        'maint-90': "Maintenance Due In Next 90 Days",
        'maint-180': "Maintenance Due In Next 180 Days"
    };
    
    const selectedTitle = maintTitles[filterType] || "Upcoming Maintenance Tasks";

    displayTitle = `
        <div style="display: flex; align-items: center; gap: 15px;">
            <button class="action-btn" 
                    style="padding: 5px 12px; font-size: 0.8rem; background: #7f8c8d;" 
                    onclick="loadTableData('Master_Dashboard')">
                ← Back to Dashboard
            </button>
            <span>${selectedTitle}</span>
        </div>`;        
}
    //====== END Shop Maintenance title logic ==============

    


    // Hand off cleaned data to to UI module
    //Pass 1. The Rows, 2. The Table Name, 3. The Config Blueprint
    UI.renderTable(formattedRows, tableName, sheetConfig, displayTitle);

    // Draw command bar at the bottom
    UI.renderCommandBar(tableName);


   } catch (error) {
    console.error("MAE System: Error loading table data:", error);
    UI.showError("Error: Could not load data.  Ensure spreadsheet is closed in Excel");
   }
} 
//=========== END loadTableData ===================


//======= helper for filtering logic====
function applyDashboardFilters(tableName, rows, filterType) {
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    const now = new Date();

    // Safety check: if no config is found, return all rows to prevent a crash
    if (!sheetConfig) {
        console.warn(`MAE System: No config found for ${tableName}. Skipping filters.`);
        return rows;
    }
    
    return rows.filter(row => {
        // RUGGED: Extract the inner array [cleanValues]
        const values = row.values[0]; 
        
        switch (filterType) {
            case 'low-stock':
                const stockIdx = sheetConfig.columns.findIndex(c => c.header === "Current Stock");
                const reorderIdx = sheetConfig.columns.findIndex(c => c.header === "Reorder Point");
    
                // RUGGED: Extract the raw value from the cell
                const stockVal = values[stockIdx];
    
                // 1. Text-Based Trigger: If the owner feels there are "Few," it's a Low Stock alert.
                if (stockVal === "Few") return true;
    
                // 2. Text-Based Pass: If it's "Adequate" or "Many," it's not an alert.
                if (stockVal === "Adequate" || stockVal === "Many") return false;
    
                // 3. Number-Based Trigger: Standard logic for numeric entries
                const numericStock = parseFloat(stockVal);
                const numericReorder = parseFloat(values[reorderIdx]);
    
                return !isNaN(numericStock) && numericStock <= numericReorder;
            
            case 'needs-repair':
                const conditionIdx = sheetConfig.columns.findIndex(c => c.header === "Condition");
                const condition = (values[conditionIdx] || "").toLowerCase();
                // MISSION: Sum of Needs Repair, Repair In-Progress, and Unusable/Junk
                return ["needs repair", "repair in-progress", "unusable/junk"].includes(condition);

            case 'resell-active':
                const statusIdx = sheetConfig.columns.findIndex(c => c.header === "Current Status");
    
                // Graph API rows store cell data in row.values[0] 
                // row.values is a 2D array [["Item", "Status", ...]]
                const rowCells = Array.isArray(row.values[0]) ? row.values[0] : row.values;
                // 2. Clean the value to ensure NO hidden spaces break the match
                 const rawStatus = (rowCells[statusIdx] || "").toString().trim();
                // 3. Match against your exact dropdown options
                const activeStatuses = ["In-Progress", "Complete", "For Sale"];
                // DEBUG: This will show you exactly what the app is seeing in your browser console
                console.log(`MAE System: Row status is [${rawStatus}]`); 

                return activeStatuses.includes(rawStatus);

            case 'maint-30':
                // index 8 = Next Service Date
                const nextDate = new Date(values[8]);
                const thirtyDays = new Date();
                thirtyDays.setDate(now.getDate() + 30);
                return nextDate >= now && nextDate <= thirtyDays;
            // Filter logic for Shop Overhead Dashboard buttons
            case 'due-7':   return isWithinDays(row.values[0], sheetConfig, 7);
            case 'due-30':  return isWithinDays(row.values[0], sheetConfig, 30);
            case 'due-90':  return isWithinDays(row.values[0], sheetConfig, 90);
            case 'due-180': return isWithinDays(row.values[0], sheetConfig, 180);
            // Filter logic for Shop Maintenance Dashboard buttons
            case 'maint-7':   return isWithinDays(row.values[0], sheetConfig, 7, "Next Service Date");
            case 'maint-30':  return isWithinDays(row.values[0], sheetConfig, 30, "Next Service Date");
            case 'maint-90':  return isWithinDays(row.values[0], sheetConfig, 90, "Next Service Date");
            case 'maint-180': return isWithinDays(row.values[0], sheetConfig, 180, "Next Service Date");

            default: return true;
        }
    });
}
//=========== END helper for filter logic ===================


// WORKER function to check if a row's date is within a certain number of days
function isWithinDays(rowValues, sheetConfig, days, colName = "Due Date") {
    const dateIdx = sheetConfig.columns.findIndex(c => c.header === colName);
    
    // Safety check: if no "Due Date" column exists in this table, skip
    if (dateIdx === -1 || !rowValues[dateIdx]) return false;

    // Convert Excel serial to JS Date
    const dueDate = new Date(excelSerialToDate(rowValues[dateIdx]));
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const limitDate = new Date();
    limitDate.setDate(today.getDate() + days);
    limitDate.setHours(23, 59, 59, 999);

    return dueDate >= today && dueDate <= limitDate;
}
// ====== END WORKER FUNCTION ===============

//=========== GLOBAL CLICK LISTENER FUNCTION===========
// 1. GLOBAL CLICK LISTENER (Event Delegation)
// This stays active even when buttons are deleted/recreated
// Updated Global Click Listener
document.getElementById('action-bar-zone').addEventListener('click', (event) => {
    const btn = event.target.closest('button');
    if (!btn) return;

    const config = window.maeSystemConfig;
    const currentTable = window.currentTable;

    if (btn.id === 'btn-commit-sync') {
        btn.disabled = true;
        btn.innerText = "⌛ Syncing...";
         // Force all inputs to "blur" so they save their values before the scraper runs
        if (document.activeElement) document.activeElement.blur();
        // 1. Trigger the batch sync
        setTimeout(() => {
        processInPlaceTableUpdate(window.currentTable);
        }, 100)};
    } 
    else if (btn.id === 'btn-discard-edit') {
        if (confirm("Discard all unsaved changes?")){
        UI.exitEditMode();// Passes true to trigger a full refresh
        }
    }
    else if (btn.id === 'btn-add') {
        handleAddClick(currentTable); 
    } 
    else if (btn.id === 'btn-edit') {
        handleEditClick(currentTable);
    } 
    // CONSOLIDATED PRINT LOGIC: Handles Table, Manual Log, and TBD Audit
    else if (btn.id === 'btn-print' || btn.id === 'btn-manual-print' || btn.id === 'btn-print-audit') {
        const sheetConfig = config.worksheets.find(s => s.tableName === currentTable);
        const today = new Date().toLocaleDateString('en-US');
        
        // RUGGED CHECK: Determine if we are looking at an Audit View or a standard Map/Table
        const table = document.getElementById("main-data-table");
        const isAuditViewActive = table && table.innerHTML.includes("Assign New Location");

        let finalPrintTitle;

        if (isAuditViewActive || btn.id === 'btn-print-audit') {
            // Case 1: TBD Audit View (Two columns, grouped)
            finalPrintTitle = `Items with Location_ID: TBD (as of ${today})`;
        } else if (currentTable === "Location") {
            // Case 2: Standard Workshop Location Map (4 columns)
            finalPrintTitle = `Workshop Location Map (as of ${today})`;
        } else {
            // Case 3: Standard Inventory Tables (Consumables, Machinery, etc.)
            const titleElement = document.getElementById("current-view-title");
            const spanElement = titleElement.querySelector("span");
            const currentTitleText = spanElement ? spanElement.innerText : titleElement.innerText;
            finalPrintTitle = `${currentTitleText} (as of ${today})`;
        }
        
        // Branch to appropriate UI function
        // Note: For Audit view, we pass 'null' for sheetConfig so it prints exactly what's on screen
        if (btn.id === 'btn-print' || btn.id === 'btn-print-audit') {
            UI.printTable(currentTable, isAuditViewActive ? null : sheetConfig, finalPrintTitle);
        } else {
            UI.printManualLog(currentTable, sheetConfig, finalPrintTitle);
        }
    } 
    else if (btn.id === 'btn-inventory-update') {
        handleQuickUpdate(currentTable);
    }
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

//=========Helper to bridge Scan directly to an Add Form ====

async function handleAddClickWithId(scannedId) {
    const container = document.getElementById("table-container");
    
    // 1. Create a "Shop-Ready" selector
    container.innerHTML = `
        <div style="padding:20px; text-align:center;">
            <h3>Select Category for Tag: ${scannedId}</h3>
            <select id="mobile-table-selector" class="edit-dropdown" style="margin-bottom:20px; height:50px; font-size:1.2rem;">
                <option value="Shop_Machinery">Shop Machinery</option>
                <option value="Shop_Power_Tools">Shop Power Tools</option>
                <option value="Shop_Hand_Tools">Shop Hand Tools</option>
                <option value="Shop_Consumables">Shop Consumables</option>
                <option value="Resell_Inventory">Resell Inventory</option>
            </select>
            <button class="action-btn" onclick="confirmMobileAdd('${scannedId}')" style="width:100%; background:#27ae60;">
                Continue to Form
            </button>
        </div>
    `;
}

// Helper to bridge the selection
window.confirmMobileAdd = async (scannedId) => {
    const tableName = document.getElementById('mobile-table-selector').value;

    // 1. Clear the selector UI first so the Add Form has a clean stage
    document.getElementById("table-container").innerHTML = ""; 

    // Store the ID in the global "Mailbox"
    window.pendingScanValue = scannedId;

    // 2. Open the standard Add Form
    await handleAddClick(tableName);

    // 3. RUGGED INJECTION: Wait for the DOM to render the form, then inject the ID
    setTimeout(() => {
        // Your code generates IDs using 'field-' + header name
        const idInput = document.getElementById("field-mae_id");
        
        if (idInput) {
            idInput.value = scannedId;
            // Visual confirmation for the developer/user that injection happened
            idInput.style.backgroundColor = "#fffde7"; 
            console.log("MAE System: Scanned ID successfully injected into field-mae_id:", scannedId);
        } else {
            // If this logs, we know the form didn't render the mae_id field
            console.warn("MAE System: Could not find input field-mae_id. Check if it's rendered in ui.js");
        }
    }, 500); // 500ms is usually the "sweet spot" for DOM rendering
};
//======= END Helper to bridge Scan to Add Form ===============

// ============ GLOBAL "CLICK OFF" HANDLER for all Edit Modes ==========
async function globalClickOffHandler(e) {
    // 1. If we aren't in Edit Mode, this handler does nothing
    if (!window.isEditing) return;

    const table = document.getElementById("main-data-table");
    const container = document.getElementById("table-container");
    if (!table || !container) return;

    // SCROLLBAR DETECTION 
    const rect = container.getBoundingClientRect();
    const isScrollbarClick = 
        (e.clientX > rect.left + container.clientWidth) || 
        (e.clientY > rect.top + container.clientHeight);

    // 2. IDENTIFY SAFE ZONES
    const isInsideTable = table.contains(e.target);
    const isCommitBtn = e.target.closest('#btn-commit-sync');
    const isDiscardBtn = e.target.closest('#btn-discard-edit');
    const isEntryForm = e.target.closest('#entry-form');

    // 3. THE GUARDRAIL: If they click outside the "Safe Zones"
    if (!isInsideTable && !isCommitBtn && !isDiscardBtn && !isScrollbarClick && !isEntryForm) {
        
        // Use the rugged browser confirm dialog
        const confirmDiscard = confirm("MAE System: You have uncommitted changes. Discard all changes and exit Edit Mode?");
        
        if (confirmDiscard) {
            console.log("MAE System: User chose to discard changes.");
            
            // Immediately stop listening for clicks to prevent loops
            document.removeEventListener('mousedown', globalClickOffHandler);
            
            // Clean up the UI without saving to OneDrive
            UI.exitEditMode();
        } else {
            // User clicked 'Cancel', keep everything exactly as it is
            console.log("MAE System: Discard cancelled. Continuing edit session.");
        }
    }
}
//=======END GLOBAL "CLICK OFF" HANDLER ===========

//========== Handle EDIT CLICK function ================
function handleEditClick(tableName) {
    window.isEditing = true; // Turn on the flag
    UI.renderCommandBar(tableName); // Refresh the buttons immediately

    const table = document.getElementById("main-data-table");
    if (!table || table.classList.contains("is-editing")) return;

    window.currentTable = tableName; 
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    table.classList.add("is-editing");
    
    const cells = table.querySelectorAll(".editable-cell");

    cells.forEach(cell => {
        const colIdx = parseInt(cell.getAttribute('data-col-index'));
        const colDef = sheetConfig.columns[colIdx];

        // RUGGED: Ensure EVERY cell is interactive and "on top"
        cell.style.position = "relative";
        cell.style.zIndex = "100";
        cell.style.pointerEvents = "auto";
        cell.setAttribute('tabindex', '0');
        cell.onmousedown = (e) => e.stopPropagation();

        //===Location_ID: treat as "Dropdown READ-Only" to prevent accidental changes whe bulk editing
        if (colDef.header === "Location_ID") {
            cell.contentEditable = "false"; 
            cell.classList.add("dropdown-edit-zone");

            // RUGGED: Create a dropdown using the established Location Map cache
            const currentVal = cell.innerText.trim();
            let selectHtml = `<select class="edit-dropdown" style="width:100%; border:none; background:#fffde7;">`;
    
            // Always pull from the window.maeLocations cache we refreshed on startup
            window.maeLocations.forEach(loc => {
                selectHtml += `<option value="${loc}" ${loc === currentVal ? 'selected' : ''}>${loc}</option>`;
            });
            selectHtml += `</select>`;

            cell.innerHTML = selectHtml;
            const select = cell.querySelector('select');
    
            // Save the change when they pick a new one or click away
            const finishEdit = () => {
                cell.innerText = select.value;
                cell.classList.remove("dropdown-edit-zone");
            };

            select.onchange = finishEdit;
            select.onblur = finishEdit;
                    return; // Move to the next cell
        }

        // --- BRANCH: HYBRID INVENTORY (Consumables/Hand Tools) ---
       
        if (colDef.type === "hybrid-inventory") {
            cell.contentEditable = "false";
            const currentVal = cell.innerText.trim();
            const rowIndex = cell.closest('tr').getAttribute('data-row-index');
    
            // Use 'colIdx' instead of 'idx' to fix the ReferenceError
            const tempId = `edit-hybrid-${colIdx}-${rowIndex}`;
    
            // Call your new modular component
            cell.innerHTML = UI.createHybridInventoryHTML(tempId, currentVal);
    
            // Prevent the click-off sync from firing when clicking the dropdown
            cell.onmousedown = (e) => e.stopPropagation();
    
            return; // Move to the next cell
        }


        // --- BRANCH 1: NUMBERS (Integer & Currency) ---
        if (colDef.type === "number") {
            const isCurrency = colDef.format && colDef.format.includes("$");
            const currentVal = cell.innerText.replace(/[^0-9.-]+/g, "") || 0;

            cell.contentEditable = "false"; 
            cell.innerHTML = `<input type="number" class="edit-number-input" value="${currentVal}" step="${isCurrency ? '0.01' : '1'}" min="0">`;
    
            const input = cell.querySelector('input');
            input.onkeydown = (e) => {
                if (e.key.toLowerCase() === "e") e.preventDefault();
                if (!isCurrency && (e.key === "." || e.key === ",")) e.preventDefault();
            };

            input.onblur = () => {
                let val = parseFloat(input.value) || 0;
                cell.innerText = isCurrency ? val.toFixed(2) : Math.floor(val).toString();
            };
        } 
        // ---- BRANCH 2: BOOLEAN CHECKBOXES ---
        else if (colDef.type === "boolean") {
            cell.contentEditable = "false"; // Rugged: No typing allowed
            const isChecked = cell.innerText.trim().toUpperCase() === "TRUE" || 
                      (cell.querySelector('input') && cell.querySelector('input').checked);
    
            // Inject the checkbox
            cell.innerHTML = `<input type="checkbox" class="mae-checkbox" ${isChecked ? 'checked' : ''}>`;
    
            const checkbox = cell.querySelector('input');
    
            // Stop propagation: clicking the checkbox shouldn't trigger the global "Click-Off" sync immediately
            checkbox.onmousedown = (e) => e.stopPropagation();
    
            // Ensure clicking the cell toggles the checkbox
            cell.onclick = (e) => {
                if (e.target !== checkbox) {
                    checkbox.checked = !checkbox.checked;
                }
            };
        }

        // --- BRANCH 3: DROPDOWNS ---
        else if (colDef.type === "dropdown") {
            cell.contentEditable = "false"; 
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
                    const selectedText = select.options[select.selectedIndex].text;
                    cell.innerHTML = selectedText; 
                    cell.classList.remove("dropdown-edit-zone");  
                };
       
                select.onchange = finishEdit;
                select.onblur = finishEdit;
                select.onkeydown = (k) => { if(k.key === 'Enter') finishEdit(); };
            };

            cell.onclick = startDropdownEdit;
            cell.onkeydown = (k) => { if(k.key === 'Enter') startDropdownEdit(k); };
        } 
        // --- BRANCH 3: STANDARD TEXT ---
        else {
            cell.contentEditable = "true";
            cell.classList.add("text-edit-focus"); 
        }
    }); // This correctly closes the cells.forEach

    setTimeout(() => {
        document.addEventListener('mousedown', globalClickOffHandler);
    }, 150);
}

// ====== END handleEditClick function =================


//===========FUNCTION submitNewRow====to send data to Microsoft========

async function submitNewRow(tableName, sheetConfig) {
    const rowData = sheetConfig.columns.map(col => {
        const fieldId = `field-${col.header.replace(/\s+/g, '')}`;
        const input = document.getElementById(fieldId);

        // 1. Primary Key: mae_id
        if (col.header === "mae_id") {
            const scannedValue = input ? input.value.trim() : "";
            return (scannedValue !== "") ? scannedValue : `MAE-${Date.now()}`; 
        }

        // 2. Formulas: Always null (let Excel calculate)
        if (col.type === "formula") return null;

        // 3. Checkboxes: Boolean Logic
        if (col.type === "boolean") {
            return input ? input.checked : false; 
        }

        // 4. Numbers & Currency: Float logic
        if (col.type === "number" || (col.format && col.format.includes("$"))) {
            if (!input || input.value === "") return null;
            const num = parseFloat(input.value);
            return isNaN(num) ? null : num;
        }

        // 5. Dropdowns & Strings: Logic for "Control Tower" consistency
        if (!input || input.value === "") {
            // RUGGED: If Location_ID is empty, force it to the "Intake" bucket (TBD)
            if (col.header === "Location_ID") return "TBD";
            return (col.type === "date") ? null : "";
        }

        // Hybrid-Inventory
        if (col.type === "hybrid-inventory") {
            const select = document.getElementById(fieldId);
            const numInput = document.getElementById(`${fieldId}-num`);
    
            // Only return the number if the dropdown is actually set to "Number"
            if (select && select.value === "Number") {
                return (numInput && numInput.value !== "") ? parseInt(numInput.value) : 0;
            }
            return select ? select.value : "PEND";

        }

        // Location ID: ensure defaults to "TBD" if empty
        if (col.header === "Location_ID" && (!input || input.value === "")) {
            return "TBD"; // Rapid entry fallback
        }

        // Return trimmed string for clean Excel data
        return input.value.trim();
    });

    try {
        const token = await window.getGraphToken();
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows`;
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowData] }) // Values must be 2D array
        });

        if (response.ok) {
            console.log(`MAE System: Row successfully added to ${tableName}`);
            alert("Entry Saved Successfully!");
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
    const rows = table.querySelectorAll("tbody tr:not(.repair-group-header tr)");
    const title = document.getElementById("current-view-title");
    
    const updates = [];

    // 1. DATA HARVESTING (The Scraper)
    rows.forEach(tr => {
        const rowIndex = tr.getAttribute("data-row-index");
        const rowValues = [];

        sheetConfig.columns.forEach((col, index) => {
            if (col.type === "formula") { rowValues.push(null); return; }
            if (col.header === "mae_id") { rowValues.push(tr.getAttribute('data-mae-id') || ""); return; }

            const cell = tr.querySelector(`td[data-col-index="${index}"]`);
            if (!cell) { rowValues.push(col.type === "number" ? null : ""); return; }

            let val = "";

            // --- 1. HYBRID INVENTORY (High Priority Check) ---
            if (col.type === "hybrid-inventory") {
                const hSelect = cell.querySelector('select');
                const hInput = cell.querySelector('.edit-number-input');
                
                if (hSelect && hSelect.value === "Number") {
                    // If the box is visible, TAKE THE VALUE. 
                    // If it's null, we MUST take the innerText as a last resort.
                    const liveValue = hInput ? hInput.value : cell.innerText.trim();
                    val = (liveValue !== "") ? parseInt(liveValue) : 0;

                } else if (hSelect) {
                    val = hSelect.value; 
                } else {
                    val = cell.innerText.trim();
                }
            } 
            // --- 2. BOOLEAN CHECK (Specific lookup) ---
            else if (col.type === "boolean") {
                const cb = cell.querySelector('input[type="checkbox"]');
                val = cb ? cb.checked : (cell.innerText.trim().toUpperCase() === "TRUE");
            }
            // --- 3. GENERIC LOOKUPS (Else catch-all) ---
            else {
                const select = cell.querySelector('select');
                const input = cell.querySelector('input:not([type="checkbox"])');

                if (select) {
                    val = select.value;
                } else if (input) {
                    val = input.value;
                } else {
                    val = cell.innerText.replace(/[$,]/g, "").trim();
                }
            }

            // --- 4. TYPE ENFORCEMENT ---
            if (col.type === "number") {
                const isCurrency = col.format && col.format.includes("$");
                // If val is null, undefined, or an empty string, 
                // we send a literal empty string "". Excel interprets "" as "no change/blank" 
                // better than it handles null in a batch patch.
                if (val === "" || val === null || val === undefined) {
                    val = 0; 
                } else {
                    let cleanNum = parseFloat(val.toString().replace(/[^0-9.-]+/g, ""));
                    val = isNaN(cleanNum) ? 0 : (isCurrency ? parseFloat(cleanNum.toFixed(2)) : Math.floor(cleanNum));
                }
            }
            
            rowValues.push(val);
        });
        updates.push({ index: rowIndex, values: [rowValues] });
    });

    // 2. BATCHING & CHUNKING LOGIC
    try {
        const token = await window.getGraphToken();
        const chunkSize = 20; 
        const totalRows = updates.length;
        
        for (let i = 0; i < totalRows; i += chunkSize) {
            const chunk = updates.slice(i, i + chunkSize);
            const percent = Math.round((i / totalRows) * 100);
            if (title) title.innerText = `💾 Syncing to OneDrive: ${percent}%...`;

            const batchRequests = chunk.map((update, idx) => ({
                id: (i + idx).toString(),
                method: "PATCH",
                url: `/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${update.index})`,
                body: { values: update.values },
                headers: { "Content-Type": "application/json" }
            }));

            const response = await fetch("https://graph.microsoft.com/v1.0/$batch", {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ requests: batchRequests })
            });

            if (!response.ok) throw new Error("Batch request failed");

            const batchResult = await response.json();
            batchResult.responses.forEach(res => {
                if (res.status < 200 || res.status > 299) {
                    console.error(`MAE Sync Fail: Row ID ${res.id} failed with status ${res.status}`, res.body);
                }
            });

            await new Promise(r => setTimeout(r, 500));
        }

        if (title) title.innerText = "✅ Sync Complete";
        console.log("MAE System: Batch sync successful.");
        await new Promise(r => setTimeout(r, 600));

    } catch (err) {
        console.error("Batch Sync Error:", err);
        UI.showError("Failed to sync changes. Check connection.");
    } finally {
        UI.exitEditMode();
        
        const sheetConfigFinal = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        if (title && sheetConfigFinal) {
            title.innerText = `View: ${sheetConfigFinal.tabName}`;
        }

        console.log("MAE System: Triggering verification refresh...");
        setTimeout(() => {
            loadTableData(tableName);
        }, 2000); 
    }
}

// =====  END  function processInPlaceTableUpdate   ==========


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
        const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        const token = await window.getGraphToken();
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
        const response = await fetch(url, {
            method: 'DELETE',
            headers: { 'Authorization': `Bearer ${token}` }
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
    if (!table || table.classList.contains("is.quick-updating")) return;

    window.currentTable = tableName; 
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    
    table.classList.add("is-editing", "is-quick-updating");

    const cells = table.querySelectorAll("td");
    cells.forEach(cell => {
        const colIdxAttr = cell.getAttribute('data-col-index');
        if (colIdxAttr === null) return; 

        const colIdx = parseInt(colIdxAttr);
        const colDef = sheetConfig.columns[colIdx];
        if (!colDef) return; 

        // FIX 1: STRICT INVENTORY CHECK (Prevents arrows in Unit Cost/Price)
        const isQtyField = (colDef.header === "Quantity" || colDef.header === "Current Stock") && colDef.type === "number";
        
        if (isQtyField) {
            cell.classList.add("quick-edit-focus");
            const currentVal = parseInt(cell.innerText.replace(/[^0-9.-]+/g, "")) || 0;
            
            cell.contentEditable = "false"; 
            cell.innerHTML = `<input type="number" class="edit-number-input" value="${currentVal}" step="1" min="0">`;
            
            const input = cell.querySelector('input');
            
            // FIX 2: REMOVED the setTimeout focus from here to prevent the "Last Row Only" bug.
            // All rows will now be editable at once.

            input.onblur = () => {
                cell.innerText = input.value;
            };

            input.onkeydown = (e) => {
                if (e.key.toLowerCase() === "e") e.preventDefault();
                if (e.key === "." || e.key === ",") e.preventDefault();
            };
        } else {
            // Visual lock for non-inventory columns
            cell.contentEditable = "false";
            cell.style.backgroundColor = "#f4f4f4";
            cell.style.color = "#999";
        }
    });
    
    setTimeout(() => {
        document.addEventListener('mousedown', globalClickOffHandler);
    }, 150);
}

// ========END handleQuickUpdate ============

// =========== UPDATED UNIVERSAL SCANNER LOOKUP ===========
async function handleUniversalLookup(scannedId) {
    // RUGGED: Scanners often send hidden spaces or newline characters
    const cleanId = scannedId.toString().trim();
    console.log("MAE System: Searching Ledger for ID:", cleanId);
    
    UI.showLoading(`Searching Shop Records for: ${cleanId}...`);
    
    // Priority order for search
    const tables = ["Resell_Inventory", "Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables"];
    
    try {
        const token = await window.getGraphToken();

        for (const tableName of tables) {
            const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
            if (!sheetConfig) continue;
            // API path to the specific table
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows`;
            
            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (response.ok) {
                const data = await response.json();
                
                // FInd index of the Tag_ID column
                const tagIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
                // Finds the first unhidden column containing "Name" or "Tool"
                const nameIdx = sheetConfig.columns.findIndex(c => !c.hidden && (c.header.includes("Name") || c.header.includes("Tool")));
                if (tagIdx === -1) continue;
                
                // Find ALL rows where Tag_ID matches (allows for "Multiple" tags)
                const matchedRows = data.value.filter(row => {
                    const rowCells = row.values[0]; 
                    return String(rowCells[tagIdx]).trim() === cleanId;
                });
                
                if (matchedRows.length > 0) {
                    console.log(`MAE System: Found ${matchedRows.length} match(es) in ${tableName}`);
                    window.currentTable = tableName;

                    const uniqueTables = ["Resell_Inventory","Shop_Machinery", "Shop_Power_Tools"];
                    if (uniqueTables.includes(tableName)) {
                        console.log("MAE Integrity: Unique item detected. Blocking 'Add New'.");
                        // Only allow View/Edit, do not show the "Register New Item" prompt
                        UI.renderScanResultCard(matchedRows[0].values[0], tableName, sheetConfig, matchedRows[0].index);
                        return; 
                    }



                    // If it is just ONE item (Unique), use your existing Card view
                    if (matchedRows.length === 1) {
                        UI.renderScanResultCard(matchedRows[0].values[0], tableName, sheetConfig, matchedRows[0].index);
                    } else {
                        // 🌟 NEW CAPABILITY: If multiple items share this tag (like a drawer), 
                        // map them and send them to the Virtual Table Hub we will build next.
                        const auditData = matchedRows.map(row => {
                            const rowCells = row.values[0];
                            return {
                                category: sheetConfig.tabName,
                                itemName: rowCells[nameIdx] || "N/A",
                                mae_id: rowCells[0], // Column 0
                                tableName: tableName
                            };
                        });
                        UI.renderVirtualSearchHub(auditData);
                    }
                    return; 
                }
            }
        }       
        // NO MATCH FOUND: Industrial Prompt
        UI.showError(`
            <div style="text-align:center; padding: 20px;">
                <h3 style="color:var(--primary);">New Tag Detected</h3>
                <div style="font-size: 1.5rem; font-weight: 800; background: #fffde7; border: 2px dashed var(--accent); padding: 15px; margin: 10px 0;">
                    ID: ${cleanId}
                </div>
                <p>This ID is not recognized in the priority inventory lists.</p>
                <button class="action-btn" onclick="handleAddClickWithId('${cleanId}')" style="width:100%; margin-bottom:10px;">
                    ➕ Register New Item
                </button>
                <button class="action-btn cancel-btn" onclick="loadTableData('Master_Dashboard')" style="width:100%;">
                    Discard Scan
                </button>
            </div> 
        `);

    } catch (error) {
        console.error("MAE System: Lookup failed", error);
        UI.showError("Network error during lookup. Check workshop internet.");
    }
}
// =========== END UPDATED LOOKUP ===========

// ============ FUNCTION TO HANDLE SINGLE-ROW UPDATES===========
async function updateSingleRowFromForm(tableName, rowIndex, sheetConfig) {
    // 1. Gather data from the form fields
    const rowData = sheetConfig.columns.map(col => {
        const fieldId = `field-${col.header.replace(/\s+/g, '')}`;
        const input = document.getElementById(fieldId);
        
        if (col.type === "formula") return null; // Don't overwrite formulas
        if (col.type === "boolean") return input.checked ? "TRUE" : "FALSE";
        return input ? input.value : "";
    });

    try {
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        // 2. Target the SPECIFIC row index
        
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
        const response = await fetch(url, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowData] })
        });

        if (response.ok) {
            console.log("MAE System: Single row update successful.");
            return true;
        }
    } catch (err) {
        console.error("MAE System: Single row update failed", err);
    }
    return false;
}
//=========== END FUNCTION TO HANDLE SINGLE-ROW UPDATES ===========



//======= submitNewLocationToTable : writes data to Excel
async function submitNewLocationToTable(rowDataMap) {
    try {
        const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        
        // Map the object to the correct Excel column positions
        const newRow = sheetConfig.columns.map((col, idx) => {
            if (col.header === "mae_id") return `LOC-${Date.now()}`;
            if (rowDataMap.hasOwnProperty(col.header)) return rowDataMap[col.header];
            return ""; 
        });

        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: window.account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/Location/rows`;
      //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [newRow] })
        });

        return response.ok;
    } catch (err) {
        console.error("MAE System: Failed to establish location", err);
        return false;
    }
}

//===== END submitNewLocationToTable =============

//===== TBD location audit function =====
async function loadTbdAudit() {
    const tablesToSearch = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
    UI.showLoading("Auditing TBD Items...");
    
    let allTbdRows = [];
    
    try {
        for (const table of tablesToSearch) {
            const data = await Dashboard.getFullTableData(table);
            const config = maeSystemConfig.worksheets.find(s => s.tableName === table);
            const locIdx = config.columns.findIndex(c => c.header === "Location_ID");
            
            // RUGGED: Only process if data exists to prevent "Cannot read property of undefined"
            if (data && data.length > 0) {
                const tbdRows = data.filter(row => {
                    // Check row.values[0] exists before accessing the index
                    const rowCells = row.values[0];
                    return rowCells && rowCells[locIdx] === "TBD";
                });
                allTbdRows = allTbdRows.concat(tbdRows);
            }
        }

        // Render the results using the Location blueprint
        const locationConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        UI.renderTable(allTbdRows, "Location", locationConfig, "Audit: Items Awaiting Final Location Assignment");
        UI.renderCommandBar("Location");

    } catch (error) {
        console.error("MAE System: TBD Audit failed", error);
        UI.showError("Audit failed. Ensure all inventory tables are healthy.");
    }
}
//==== END TBD location audit function =====

//========= update location name =========
async function updateLocationRecord(rowIndex, rowDataMap) {
    try {
        const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        
        // DYNAMIC MAPPING: Build an array of nulls the same length as the config
        const rowValues = new Array(sheetConfig.columns.length).fill(null);

        // Fill the array only where the headers match the data we provided
        sheetConfig.columns.forEach((col, idx) => {
            if (rowDataMap.hasOwnProperty(col.header)) {
                rowValues[idx] = rowDataMap[col.header];
            }
        });

        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: window.account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/Location/rows/itemAt(index=${rowIndex})`;
      //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/Location/tables/Location/rows/itemAt(index=${rowIndex})`; 
        const response = await fetch(url, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowValues] })
        });

        return response.ok;
    } catch (err) {
        console.error("MAE System: Record update failed", err);
        return false;
    }
}

//====  END update location name ===========

//===========  scan all inventory tables to find "TBD" items====
async function runLocationAudit() {
    UI.showLoading("Performing Cross-Table Audit...");
    const results = await getTbdAuditData();
    
    // The list of tables we want to check for TBD locations
    const inventoryTables = ["Resell_Inventory", "Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables"];
    let auditResults = [];

    for (const tableName of inventoryTables) {
        const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        // Use your existing worker to get raw rows
        const data = await Dashboard.getFullTableData(tableName); 

        // HEADER DISCOVERY: Find the index of Location_ID and the Item Name
        const locIdx = sheetConfig.columns.findIndex(c => c.header === "Location_ID");
        // We look for the first column that isn't hidden/ID to find the "Item Name"
        const nameIdx = sheetConfig.columns.findIndex(c => !c.hidden && c.header.includes("Name") || c.header.includes("Tool"));

        data.forEach(row => {
            const cells = row.values[0];
            // Check if Location_ID is exactly "TBD"
            if (cells[locIdx] === "TBD") {
                auditResults.push({
                    category: sheetConfig.tabName,
                    itemName: cells[nameIdx],
                    mae_id: cells[0], // mae_id is always index 0 per your config
                    tableName: tableName
                });
            }
        });
    }
    // Hand off the aggregated results to the UI
    UI.renderAuditGrid(auditResults);
}
//=====  END scan all inventory tables to find "TBD items ======"

//=====  Update "Engine" for TBD =========
async function handleAuditUpdate(tableName, maeId, newLoc, rowHtmlId = null) {
    try {
        const success = await commitCellChange(tableName, maeId, "Location_ID", newLoc);
        
        if (success && rowHtmlId) {
            const row = document.getElementById(rowHtmlId);
            if (row) {
                row.style.opacity = "0";
                setTimeout(() => row.remove(), 500);
            }
        }
        return success;
    } catch (err) {
        console.error("MAE Sync Error:", err);
        return false;
    }
}

// ==== HELPER function for the "Update Engine"=====
async function commitCellChange(tableName, maeId, columnName, newValue) {
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    if (!sheetConfig) return false;
    
    const colIdx = sheetConfig.columns.findIndex(c => c.header === columnName);

    try {
        const token = await window.getGraphToken();

        // 1. Fetch fresh table data via the Dashboard module
        const data = await Dashboard.getFullTableData(tableName);
        
        // RUGGED LOOKUP: Graph API returns an array of row objects.
        // Inside each object, 'values' is a 2D array: [[col0, col1, col2...]]
        const rowIndex = data.findIndex(row => {
            // Dig into the inner array if it exists, otherwise use values as-is
            const cells = (row.values && Array.isArray(row.values[0])) ? row.values[0] : row.values;
            
            // Compare the primary key (mae_id) at Index 0
            return String(cells[0]).trim() === String(maeId).trim();
        });

        // INDUSTRIAL LOGGING: Essential for verifying the "Guardrail" found the item
        console.log(`MAE DEBUG: Looking for ID [${maeId}] in [${tableName}]. Match found at row index: ${rowIndex}`);
       
        if (rowIndex === -1) {
            console.warn(`MAE System: mae_id [${maeId}] not found in ${tableName}.`);
            return false;
        }

        // 2. Prepare Sparse Update Array
        // Fills an array with 'null' so Excel only updates the changed cell
        const rowValues = new Array(sheetConfig.columns.length).fill(null);
        rowValues[colIdx] = newValue;

        // 🌟 SAFETY FIX 1: Force the index to be a strict base-10 integer
        const strictIndex = parseInt(rowIndex, 10);

        // API endpoint targeting the specific Table Row Index directly
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(maeSystemConfig.spreadsheetName)}:/workbook/tables/${tableName}/rows/itemAt(index=${strictIndex})`;

        const response = await fetch(url, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowValues] })
        });

        if (response.ok) {
            console.log(`MAE System: Success. ${maeId} re-homed to ${newValue} in OneDrive.`);
            return true;
        } else {
            const errData = await response.json();
            console.error("MAE System: Graph API PATCH failed", errData);
            return false;
        }

    } catch (err) {
        console.error("MAE System: Quick Sync Exception", err);
        return false;
    }
}
//==== END HELPER function for the "Update Engine" =====

// ====== SCans all inventory tables for TBD items ======
async function getTbdAuditData() {
    const inventoryTables = ["Resell_Inventory", "Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables"];
    let auditResults = [];

    for (const tableName of inventoryTables) {
        const config = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        const data = await Dashboard.getFullTableData(tableName); 

        // HEADER DISCOVERY
        const locIdx = config.columns.findIndex(c => c.header === "Location_ID");
        const nameIdx = config.columns.findIndex(c => !c.hidden && (c.header.includes("Name") || c.header.includes("Tool")));

        data.forEach(row => {
            const cells = row.values;
            if (cells[locIdx] === "TBD") {
                auditResults.push({
                    category: config.tabName,
                    itemName: cells[nameIdx],
                    mae_id: cells[0], // Column 0
                    tableName: tableName
                });
            }
        });
    }
    return auditResults;
}

// ====== END   SCans all inventory tables for TBD items ======
//===== END  Update "Engine" for TBD =========

//========== Getting Location Dependencies to support Deletion of Location_ID
async function getLocationDependencies(locationId) {
    const inventoryTables = ["Resell_Inventory", "Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables"];
    let dependencies = [];

    for (const tableName of inventoryTables) {
        const config = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        const data = await Dashboard.getFullTableData(tableName);
        
        const locIdx = config.columns.findIndex(c => c.header === "Location_ID");
        const nameIdx = config.columns.findIndex(c => !c.hidden && (c.header.includes("Name") || c.header.includes("Tool")));

        data.forEach(row => {
            const cells = row.values[0]; // Graph API nested array
            if (cells[locIdx] === locationId) {
                dependencies.push({
                    tableName: tableName,
                    mae_id: cells[0],
                    itemName: cells[nameIdx],
                    currentLoc: cells[locIdx]
                });
            }
        });
    }
    return dependencies;
}

//==========  END Getting Location Dependencies to support Deletion of Location_ID

window.Dashboard = Dashboard;
window.UI = UI;
window.Labels = Labels;

window.handleEditClick = handleEditClick;
window.handleQuickUpdate = handleQuickUpdate;
window.handleAddClick = handleAddClick;
window.requestDelete = requestDelete;
window.loadTableData = loadTableData;
window.handleAddClickWithId = handleAddClickWithId;
window.handleUniversalLookup = handleUniversalLookup;
window.processInPlaceTableUpdate = processInPlaceTableUpdate;
window.updateSingleRowFromForm = updateSingleRowFromForm;
window.submitNewLocationToTable = submitNewLocationToTable;
window.refreshLocationCache = refreshLocationCache;
window.loadTbdAudit = loadTbdAudit;
window.updateLocationRecord = updateLocationRecord;
window.deleteExcelRow = deleteExcelRow;
window.runLocationAudit = runLocationAudit;
window.handleAuditUpdate = handleAuditUpdate;
window.getLocationDependencies = getLocationDependencies;
window.globalClickOffHandler = globalClickOffHandler;
window.getGraphToken = getGraphToken; 

startup();