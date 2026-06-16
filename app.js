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

            // 🌟 GLOBAL HUMAN SNAPSHOT OBSERVER AXIS 🌟
            // Tracks manual typing entries so the Shield can restore original strings during accidental bursts
            document.addEventListener('input', (event) => {
               const target = event.target;
               if (target && target.tagName === 'INPUT' && target.id !== 'field-Tag_ID') {
                   target.setAttribute("data-pre-scan-value", target.value || "");
               }
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

//===== WORKER FUNCTION: refresh location cache; specifically pulls Location List for dropdowns =====
async function refreshLocationCache() {
    try {
        const locationConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        const data = await Dashboard.getFullTableData("Location");

        if (data && data.length > 0) {
            // Find the index dynamically based on the header name in config
            const locIdx = locationConfig.columns.findIndex(c => c.header === "Location_ID");

            const list = data.map(row => {
                // 🌟 MAE ENGINE RUGGED FIXED APPARATUS 🌟
                // Safely handles both nested 2D arrays and flat single-dimensional fallback array formats
                const rowCells = (row.values && Array.isArray(row.values[0])) 
                    ? row.values[0] 
                    : (row.values && Array.isArray(row.values)) ? row.values : row;
                
                return rowCells ? rowCells[locIdx] : null;
            });

            // Clean the list: remove nulls/duplicates and keep "TBD" at the top
            window.maeLocations = ["TBD", ...new Set(list.filter(i => i && i !== "TBD"))];
            console.log("MAE System: Location dropdown cache successfully updated with raw ledger changes:", window.maeLocations);
        }
    } catch (e) {
        console.warn("MAE System Fail: Could not find 'Location_ID' column. Defaulting dropdown parameters to 'TBD'.", e);
        window.maeLocations = ["TBD"];
    }
}
//==== END WORKER FUNCTION ======




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
    
    // RUGGED: Reset UI state before fetching new data
    UI.exitEditMode(); 
    UI.showLoading(tableName);

    try {
        const token = await window.getGraphToken();
        
        // RUGGED: Explicit path including /drive/ to ensure Ledger connectivity
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
        const response = await fetch(url, {
            headers: { 'Authorization' : `Bearer ${token}` }
        });

        if (!response.ok) throw new Error(`Graph API error: ${response.status}`);
        const data = await response.json();

        let displayTitle = `View: ${sheetConfig.tabName}`;

        // ==========================================
        // 1. DASHBOARD BRIDGE
        // ==========================================
        if (tableName === "Master_Dashboard") {
            const hasData = data.value && data.value.length > 0;
            
            // Extract values from the first row [0] and nested array [0]
            const summaryValues = (hasData && data.value[0].values) 
                ? data.value[0].values[0] 
                : [0, 0, 0, 0, 0, 0, 0];

            // Update sidebar visual state
            document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
            const homeBtn = document.querySelector('.home-btn');
            if (homeBtn) homeBtn.classList.add('active');

            UI.renderDashboard(summaryValues, sheetConfig);
            UI.renderCommandBar(tableName);
            return; 
        }

        // ==========================================
        // 2. DATA FLATTENING BRIDGE
        // ==========================================
        let formattedRows = data.value.map(rowObj => {
            // Dig into Graph's nested array [[v1, v2]]
            const rawCells = (rowObj.values && Array.isArray(rowObj.values)) 
                ? rowObj.values[0] 
                : rowObj.values;
            
            const cleanValues = rawCells.map((cellValue, index) => {
                const colDef = sheetConfig.columns[index];
                
                // Convert Excel serial numbers to readable dates
                if (colDef && colDef.type === 'date') {
                    return excelSerialToDate(cellValue);
                }
                return cellValue;
            });

            return {
                ...rowObj,
                values: cleanValues 
            };
        });

        // ==========================================
        // 🌟 MAE ENGINE: CONTEXT-AWARE SORTING GATE  🌟
        // ==========================================
        
        // PATH A: Core Inventory Sheets & Location Map Axis
        const locationColumnIndex = sheetConfig.columns.findIndex(c => c.header === "Location_ID");
        
        if (locationColumnIndex !== -1) {
            formattedRows.sort((rowA, rowB) => {
                const locA = String(rowA.values[locationColumnIndex] || "").trim();
                const locB = String(rowB.values[locationColumnIndex] || "").trim();

                if (locA === locB) return 0;
                if (locA === "TBD") return -1; // Force unassigned "TBD" assets straight to the top
                if (locB === "TBD") return 1;

                return locA.localeCompare(locB, undefined, { numeric: true, sensitivity: 'base' });
            });
            console.log(`MAE Sorting Engine: [${tableName}] memory rows sorted by Location_ID axis.`);
        } 
        
        // PATH B: Maintenance Log (Sort Chronologically by Next Service Date)
        else if (tableName === "Maintenance_Log") {
            const targetDateHeader = "Next Service Date";
            const dateColumnIndex = sheetConfig.columns.findIndex(c => c.header === targetDateHeader);

            if (dateColumnIndex !== -1) {
                formattedRows.sort((rowA, rowB) => {
                    const rawValA = rowA.values[dateColumnIndex];
                    const rawValB = rowB.values[dateColumnIndex];

                    // RUGGED GUARDRAILS: Push items with blank or invalid dates to the bottom
                    if (!rawValA) return 1;
                    if (!rawValB) return -1;

                    const dateA = new Date(rawValA);
                    const dateB = new Date(rawValB);

                    // If a date fails to parse, protect the sort loop from breaking
                    if (isNaN(dateA)) return 1;
                    if (isNaN(dateB)) return -1;

                    // Chronological sort: Past/Overdue items first, moving into the future
                    return dateA - dateB;
                });
                console.log(`MAE Sorting Engine: [${tableName}] sorted chronologically by [${targetDateHeader}].`);
            }
        }

        // PATH C: Shop Overhead (Sort Chronologically by Due Date)
        else if (tableName === "Shop_Overhead") {
            const targetDateHeader = "Due Date";
            const dateColumnIndex = sheetConfig.columns.findIndex(c => c.header === targetDateHeader);

            if (dateColumnIndex !== -1) {
                formattedRows.sort((rowA, rowB) => {
                    const rawValA = rowA.values[dateColumnIndex];
                    const rawValB = rowB.values[dateColumnIndex];

                    if (!rawValA) return 1;
                    if (!rawValB) return -1;

                    const dateA = new Date(rawValA);
                    const dateB = new Date(rawValB);

                    if (isNaN(dateA)) return 1;
                    if (isNaN(dateB)) return -1;

                    return dateA - dateB;
                });
                console.log(`MAE Sorting Engine: [${tableName}] sorted chronologically by [${targetDateHeader}].`);
            }
        }

        // PATH D: Standalone Modules & Custom Extensions (e.g., Supplier Contacts)
        else if (tableName === "Supplier_Contacts") {
            const targetSortHeader = "Supplier Name";
            const sortColumnIndex = sheetConfig.columns.findIndex(c => c.header === targetSortHeader);

            if (sortColumnIndex !== -1) {
                formattedRows.sort((rowA, rowB) => {
                    const stringA = String(rowA.values[sortColumnIndex] || "").trim();
                    const stringB = String(rowB.values[sortColumnIndex] || "").trim();
                    
                    return stringA.localeCompare(stringB, undefined, { numeric: true, sensitivity: 'base' });
                });
                console.log(`MAE Sorting Engine: Standalone [${tableName}] rows sorted by [${targetSortHeader}].`);
            }
        }

        // ==========================================
        // 3. SMART FILTERS
        // ==========================================
        if (filterType) {
            formattedRows = applyDashboardFilters(tableName, formattedRows, filterType);
        }

        // ==========================================
        // 4. DYNAMIC TITLE LOGIC
        // ==========================================
        if (filterType === 'resell-active') {
            displayTitle = "RESELL INVENTORY: WIP, Complete and For Sale";
        } 
        else if (filterType === 'low-stock') {
            displayTitle = "Shop Consumables: Low Stock Alerts";
        } 
        else if (filterType === 'needs-repair') {
            const repairTitles = {
                'Shop_Machinery': "Shop Machinery: Operational Issues",
                'Shop_Power_Tools': "Shop Power Tools: Operational Issues",
                'Shop_Hand_Tools': "Shop Hand Tools: Operational Issues"
            };
            
            const baseTitle = repairTitles[tableName] || "Equipment Issues";
            
            displayTitle = `
                <div style="display: flex; align-items: center; gap: 15px;">
                    <button class="action-btn" 
                            style="padding: 5px 12px; font-size: 0.8rem; background: #7f8c8d;" 
                            onclick="loadTableData('Master_Dashboard')">
                        ← Back
                    </button>
                    <span>${baseTitle}</span>
                </div>`;        
        } 
        else if (tableName === 'Shop_Overhead' && filterType) {
            const titleMap = {
                'due-7': "Next 7 Days", 
                'due-30': "Next 30 Days", 
                'due-90': "Next 90 Days", 
                'due-180': "Next 180 Days"
            };
            
            displayTitle = `
                <div style="display: flex; align-items: center; gap: 15px;">
                    <button class="action-btn" 
                            style="padding: 5px 12px; font-size: 0.8rem; background: #7f8c8d;" 
                            onclick="loadTableData('Master_Dashboard')">
                        ← Back
                    </button>
                    <span>Bills Due: ${titleMap[filterType]}</span>
                </div>`;       
        }
        else if (tableName === 'Maintenance_Log' && filterType && filterType.startsWith('maint-')) {
            const maintTitles = {
                'maint-7': "Maintenance Due: Next 7 Days",
                'maint-30': "Maintenance Due: Next 30 Days",
                'maint-90': "Maintenance Due: Next 90 Days",
                'maint-180': "Maintenance Due: Next 180 Days"
            };
    
            const selectedTitle = maintTitles[filterType] || "Upcoming Maintenance Tasks";

            displayTitle = `
                <div style="display: flex; align-items: center; gap: 15px;">
                 <button class="action-btn" 
                            style="padding: 5px 12px; font-size: 0.8rem; background: #7f8c8d;" 
                            onclick="loadTableData('Master_Dashboard')">
                        ← Back
                    </button>
                    <span>${selectedTitle}</span>
                </div>`;        
        }

        // ==========================================
        // 5. UI HANDOFF
        // ==========================================
        UI.renderTable(formattedRows, tableName, sheetConfig, displayTitle);
        UI.renderCommandBar(tableName);

    } catch (error) {
        console.error("MAE System: Error loading table data:", error);
        UI.showError("Error: Could not load data. Ensure spreadsheet is closed in Excel.");
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
        const values = row.values; 
        
        switch (filterType) {
            case 'low-stock':
                const levelIdx = sheetConfig.columns.findIndex(c => c.header === "Stock_Level");
                const countIdx = sheetConfig.columns.findIndex(c => c.header === "Stock_Count");
                const reorderIdx = sheetConfig.columns.findIndex(c => c.header === "Reorder Point");

                const stockLevel = values[levelIdx];
                const stockCount = parseFloat(values[countIdx]);
                const reorderPoint = parseFloat(values[reorderIdx]);

                // RUGGED SILO LOGIC:
                // 1. Bulk Items: Only trigger alert if specifically marked "None" or "Few"
                if (["None", "Few"].includes(stockLevel)) return true;

                // 2. Counted Items: Only trigger alert if methodology is "Counted" AND it's below the reorder point
                if (stockLevel === "Counted") {
                    return !isNaN(stockCount) && !isNaN(reorderPoint) && stockCount <= reorderPoint;
                }

    // Otherwise (Adequate, Many, or unassigned), do not show in Low Stock list
    return false;

            case 'needs-repair':
                const conditionIdx = sheetConfig.columns.findIndex(c => c.header === "Condition");
                const condition = (values[conditionIdx] || "").toLowerCase();
                return ["needs repair", "repair in-progress", "unusable/junk"].includes(condition);

            case 'resell-active':
                const statusIdx = sheetConfig.columns.findIndex(c => c.header === "Current Status");
                const rawStatus = (values[statusIdx] || "").toString().trim();
                return ["In-Progress", "Complete", "For Sale"].includes(rawStatus);

            // Consolidated Date Filters: These all now use the updated 'isWithinDays'
            // which includes Past Due items.
            case 'due-7':   return isWithinDays(values, sheetConfig, 7);
            case 'due-30':  return isWithinDays(values, sheetConfig, 30);
            case 'due-90':  return isWithinDays(values, sheetConfig, 90);
            case 'due-180': return isWithinDays(values, sheetConfig, 180);

            case 'maint-7':   return isWithinDays(values, sheetConfig, 7, "Next Service Date");
            case 'maint-30':  return isWithinDays(values, sheetConfig, 30, "Next Service Date");
            case 'maint-90':  return isWithinDays(values, sheetConfig, 90, "Next Service Date");
            case 'maint-180': return isWithinDays(values, sheetConfig, 180, "Next Service Date");

            default: return true;
        }
    });
}
//=========== END helper for filter logic ===================


// WORKER function to check if a row's date is within a certain number of days
function isWithinDays(rowValues, sheetConfig, days, colName = "Due Date") {
    const dateIdx = sheetConfig.columns.findIndex(c => c.header === colName);
    const completeIdx = sheetConfig.columns.findIndex(c => c.header === "Complete");
    
    const rawDateVal = rowValues[dateIdx];
    if (dateIdx === -1 || !rawDateVal) return false;

    // RUGGED: If the item is marked Complete (checkbox/boolean), exclude it from "Upcoming" views
    if (completeIdx !== -1) {
        const isDone = rowValues[completeIdx];
        if (isDone === true || String(isDone).toUpperCase() === "TRUE") return false;
    }

    const dueDate = new Date(excelSerialToDate(rawDateVal));
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const limitDate = new Date();
    limitDate.setDate(today.getDate() + days);
    limitDate.setHours(23, 59, 59, 999);

    /**
     * RUGGED LOGIC:
     * We removed "dueDate >= today".
     * Now, any date from the past (Overdue) up to the future limit (Upcoming) 
     * will return true, keeping the shop owner focused on all pending tasks.
     */
    return dueDate <= limitDate;
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

    if (btn.id === "btn-commit-sync") {
        try {
            // 1. Force the active input to commit its value (triggers the onblur)
            if (document.activeElement) document.activeElement.blur();

            // 2. IMMEDIATE HARVEST: This now includes the "Sold with $0" confirm dialog
            const capturedUpdates = harvestTableData(window.currentTable);

            // 3. Visual Feedback: Only happens if harvesting succeeded
            btn.disabled = true;
            btn.innerText = "⌛ Syncing...";

            // 4. Start Sync
            processInPlaceTableUpdate(window.currentTable, capturedUpdates);

     } catch (err) {
            // RUGGED RECOVERY: If the user clicked "Cancel" on the warning, 
            // the harvest throws an error. we catch it here to reset the UI.
            console.warn("MAE System:", err.message);
        
            // Reset the button so the user can fix the price and try again
            btn.disabled = false;
            btn.innerText = "💾 Commit Changes";
        
            // No alert needed here because harvestTableData already showed the confirm/error
        }
    }
    else if (btn.id === 'btn-discard-edit') {
        if (confirm("Discard all unsaved changes?")) {
            // FIX: Pass 'true' so ui.js knows to force a refresh from OneDrive
            UI.exitEditMode(true); 
        }
    } 
    else if (btn.id === 'btn-add') {
        handleAddClick(currentTable); 
    } 
    else if (btn.id === 'btn-edit') {
        handleEditClick(currentTable);
    } 
    else if (btn.id === 'btn-print' || btn.id === 'btn-manual-print' || btn.id === 'btn-print-audit') {
        const sheetConfig = config.worksheets.find(s => s.tableName === currentTable);
        const today = new Date().toLocaleDateString('en-US');
        const table = document.getElementById("main-data-table");
        const isAuditViewActive = table && table.innerHTML.includes("Assign New Location");

        let finalPrintTitle;
        if (isAuditViewActive || btn.id === 'btn-print-audit') {
            finalPrintTitle = `Items with Location_ID: TBD (as of ${today})`;
        } else if (currentTable === "Location") {
            finalPrintTitle = `Workshop Location Map (as of ${today})`;
        } else {
            const titleElement = document.getElementById("current-view-title");
            const spanElement = titleElement.querySelector("span");
            const currentTitleText = spanElement ? spanElement.innerText : titleElement.innerText;
            finalPrintTitle = `${currentTitleText} (as of ${today})`;
        }
        
        if (btn.id === 'btn-print' || btn.id === 'btn-print-audit') {
            UI.printTable(currentTable, isAuditViewActive ? null : sheetConfig, finalPrintTitle);
        } else {
            UI.printManualLog(currentTable, sheetConfig, finalPrintTitle);
        }
    } 
    else if (btn.id === 'btn-inventory-update') {
        handleQuickUpdate(currentTable);
    }
    else if (btn.id === 'btn-resell-status-pivot') {
        if (typeof UI.renderStatusPivotControls === "function") {
            UI.renderStatusPivotControls();
        }
    }
});

//=========== END GLOBAL CLICK LISTENER ==============

//=========== Harvest FUnction /helper 
function harvestTableData(tableName) {
    const table = document.getElementById("main-data-table");
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    const rows = table.querySelectorAll("tbody tr:not(.repair-group-header tr)");
    const updates = [];

    rows.forEach(tr => {
        const rowIndex = tr.getAttribute("data-row-index");
        const rowValues = [];

        sheetConfig.columns.forEach((col, index) => {
            // 1. PROTECTIVE GUARDRAILS: Skip formulas and IDs
            if (col.type === "formula") { rowValues.push(null); return; }
            if (col.header === "mae_id") { rowValues.push(tr.getAttribute('data-mae-id') || ""); return; }

            const cell = tr.querySelector(`td[data-col-index="${index}"]`);
            if (!cell) { 
                rowValues.push(col.type === "number" ? null : ""); 
                return; 
            }

            // --- STEP 7: VALUATION GOVERNANCE GUARD ---
            // If the cell is locked by the methodology silo, force NULL to wipe 'Zombie Data'
            if (cell.classList.contains('silo-locked')) {
             rowValues.push(null);
                return;
            }

            // 2. UNIVERSAL HARVEST: Grab current value from Input, Select, or Text
            const input = cell.querySelector('input');
            const select = cell.querySelector('select');
    
            let val;
            if (select) {
                val = select.value;
            } else if (input) {
                val = (input.type === 'checkbox') ? input.checked : input.value;
            } else {
                val = cell.innerText.replace(/[$,]/g, "").trim();
            }

            // 3. RUGGED TYPE ENFORCEMENT & NULL PROTECTION
            if (col.type === "number") {
                // RUGGED: If the field is empty, send null (not 0) to keep Excel formulas healthy
                if (val === "" || val === null) {
                    rowValues.push(null);
                    return;
                }

                const isCurrency = col.format && col.format.includes("$");
                let cleanNum = parseFloat(val.toString().replace(/[^0-9.-]+/g, ""));
        
                if (isNaN(cleanNum)) {
                    val = null; // Send null for invalid entries to prevent formula corruption
                } else {
                    val = isCurrency ? parseFloat(cleanNum.toFixed(2)) : Math.floor(cleanNum);
                }
            } else if (col.type === "boolean") {
                val = (val === true || val.toString().toUpperCase() === "TRUE");
            } else {
                val = val.toString().trim();
                if (val === "") val = null; // Standardize empty strings as null
            }

            rowValues.push(val);
        });

        // --- NEW: VALIDATION CHECK FOR INCOMPLETE SALES ---
        // This stops the sync if a 'Sold' item has no price, unless the user confirms.
        if (tableName === "Resell_Inventory") {
            const statusIdx = sheetConfig.columns.findIndex(c => c.header === "Current Status");
            const priceIdx = sheetConfig.columns.findIndex(c => c.header === "Actual Sale Price");
            
            const currentStatus = (rowValues[statusIdx] || "").toString().trim();
            const currentPrice = rowValues[priceIdx] || 0;

            if (currentStatus === "Sold" && currentPrice <= 0) {
                const itemNameIdx = sheetConfig.columns.findIndex(c => c.header === "Item Name");
                const itemName = rowValues[itemNameIdx] || "this item";
                
                const proceed = confirm(`MAE System: "${itemName}" is marked SOLD but has a $0.00 price.\n\nSync anyway and leave the orange highlight as a reminder?`);
                if (!proceed) {
                    throw new Error("Sync cancelled to correct price.");
                }
            }
        }

        updates.push({ index: rowIndex, values: [rowValues] });
    });

    console.log(`MAE System: Harvested ${updates.length} rows from ${tableName}. Ready for Sync.`);
    return updates;
}

// ===== END Harvest FUnction /helper 

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
    window.isEditing = true; 
    UI.renderCommandBar(tableName); 

    const table = document.getElementById("main-data-table");
    if (!table || table.classList.contains("is-editing")) return;

    window.currentTable = tableName; 
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    table.classList.add("is-editing");
    
    const cells = table.querySelectorAll(".editable-cell");

    cells.forEach(cell => {
        const colIdx = parseInt(cell.getAttribute('data-col-index'));
        const colDef = sheetConfig.columns[colIdx];

        cell.style.position = "relative";
        cell.style.zIndex = "100";
        cell.style.pointerEvents = "auto";
        cell.setAttribute('tabindex', '0');
        cell.onmousedown = (e) => e.stopPropagation();

        // --- BRANCH 1A: LOCATION_ID ---
        if (colDef.header === "Location_ID") {
            cell.contentEditable = "false"; 
            cell.classList.add("dropdown-edit-zone");
            const currentVal = cell.innerText.trim();
            let selectHtml = `<select class="edit-dropdown" style="width:100%; border:none; background:#fffde7;">`;
            window.maeLocations.forEach(loc => {
                selectHtml += `<option value="${loc}" ${loc === currentVal ? 'selected' : ''}>${loc}</option>`;
            });
            selectHtml += `</select>`;
            cell.innerHTML = selectHtml;
            const select = cell.querySelector('select');
            const finishEdit = () => {
                cell.innerText = select.value;
                cell.classList.remove("dropdown-edit-zone");
            };
            select.onchange = finishEdit;
            select.onblur = finishEdit;
            return; 
        }

        // BRANCH 1B - DYNAMIC TAG_TYPE RE-ORIENTATION AXIS 🌟
        if (colDef.header === "Tag_Type") {
            cell.contentEditable = "false";
            cell.classList.add("dropdown-edit-zone");
            const currentVal = cell.innerText.trim() || "UNIQUE"; // Grab existing or fallback to default
            
            let selectHtml = `<select class="edit-dropdown" style="width:100%; border:none; background:#fffde7; font-weight:bold; color:var(--primary);">`;
            selectHtml += `<option value="UNIQUE" ${currentVal === "UNIQUE" ? 'selected' : ''}>UNIQUE</option>`;
            selectHtml += `<option value="MULTIPLE" ${currentVal === "MULTIPLE" ? 'selected' : ''}>MULTIPLE</option>`;
            selectHtml += `</select>`;
            
            cell.innerHTML = selectHtml;
            const select = cell.querySelector('select');
            
            const finishEdit = () => {
                cell.innerText = select.value;
                cell.classList.remove("dropdown-edit-zone");
            };
            
            select.onchange = finishEdit;
            select.onblur = finishEdit;
            return;
        }

        // BRANCH 1C - TAG_ID DISCIPLINARY LOCKDOWN GUARD 🌟
        if (colDef.header === "Tag_ID") {
            cell.contentEditable = "false";
            // Strip any accidental editing borders and turn it into a standard locked element
            cell.classList.remove("text-edit-focus");
            cell.classList.add("locked-cell");
            cell.style.backgroundColor = "#eeeeee"; // High-visibility locked gray tint
            cell.style.color = "#888888";
            cell.style.pointerEvents = "none";      // Shield mouse clicks completely
            return; // Escape cell processing loop early to prevent standard text initialization
        }


        // --- BRANCH 2: NUMBERS & RUGGED VALIDATION ---
        if (colDef.type === "number") {
            const isCurrency = colDef.format && colDef.format.includes("$");
            // RUGGED: Extract raw numeric value, preserving decimals only for currency
            const currentVal = cell.innerText.replace(/[^0-9.-]+/g, "") || 0;
            
            cell.contentEditable = "false"; 
            cell.innerHTML = `<input type="number" class="edit-number-input" value="${currentVal}" step="${isCurrency ? '0.01' : '1'}" min="0">`;
            
            const input = cell.querySelector('input');

            // 1. REAL-TIME VALIDATION (Fat-Finger Guardrail)
            input.oninput = () => {
                // Strips non-numeric characters immediately
                const regex = isCurrency ? /[^0-9.]/g : /[^0-9]/g;
                if (regex.test(input.value)) {
                    input.value = input.value.replace(regex, "");
                }

                // Visual feedback: Red for empty/invalid, Green for valid
                if (input.value === "" || isNaN(parseFloat(input.value))) {
                    input.style.border = "2px solid #e74c3c";
                    input.style.backgroundColor = "#fadbd8";
                } else {
                    input.style.border = "2px solid #27ae60";
                    input.style.backgroundColor = "#e8f8f5";
                }
            };

            // 2. KEYBOARD PROTECTION
            input.onkeydown = (e) => {
                // Prevent 'e', '+', and '-' which are valid in scientific notation but break inventory
                if (["e", "E", "+", "-"].includes(e.key)) e.preventDefault();
                // Prevent decimals if this is a whole-number count field
                if (!isCurrency && (e.key === "." || e.key === ",")) e.preventDefault();
            };

            // 3. BLUR HANDLER (Commit to UI)
            input.onblur = () => {
                if (cell.contains(input)) {
                    let val = parseFloat(input.value) || 0;
                    // Format correctly for display, but keep the cell ready for re-edit
                    cell.innerText = isCurrency ? UI.formatCurrency(val) : Math.floor(val).toString();
                    cell.style.zIndex = ""; 
                    
                    // RUGGED: If this is part of a methodology silo, ensure the UI state is refreshed
                    if (tableName === "Shop_Consumables") {
                        UI.refreshSiloLocks(cell.closest('tr'), tableName, sheetConfig);
                    }
                }
            };
        }
        
        // ---- BRANCH 3: BOOLEAN ---
        else if (colDef.type === "boolean") {
            cell.contentEditable = "false"; 
            const isChecked = cell.innerText.trim().toUpperCase() === "TRUE" || 
                      (cell.querySelector('input') && cell.querySelector('input').checked);
            cell.innerHTML = `<input type="checkbox" class="mae-checkbox" ${isChecked ? 'checked' : ''}>`;
            const checkbox = cell.querySelector('input');
            checkbox.onmousedown = (e) => e.stopPropagation();
            cell.onclick = (e) => { if (e.target !== checkbox) checkbox.checked = !checkbox.checked; };
        }

        // --- BRANCH 4: DROPDOWNS (METHODOLOGY OBSERVER) ---
        else if (colDef.type === "dropdown") {
            cell.contentEditable = "false"; 
            cell.classList.add("dropdown-edit-zone");

            const startDropdownEdit = (e) => {
                e.stopPropagation();
                if (cell.querySelector('select')) return;

                const currentVal = cell.innerText.trim();
                cell.setAttribute('data-old-value', currentVal);

                let selectHtml = `<select class="edit-dropdown" style="width:100%; height:100%; border:none; background:#fffde7; font:inherit; cursor:pointer;">`;
                const options = colDef.options || [];
                options.forEach(opt => {
                    selectHtml += `<option value="${opt}" ${opt === currentVal ? 'selected' : ''}>${opt}</option>`;
                });
                selectHtml += `</select>`;

                cell.innerHTML = selectHtml;
                const select = cell.querySelector('select');
                select.focus();

                const finishEdit = () => {
                    cell.innerHTML = select.value; 
                    cell.classList.remove("dropdown-edit-zone");  
                };

                select.onchange = () => {
                    const newVal = select.value;
                    const oldVal = cell.getAttribute('data-old-value');
                    const row = cell.closest('tr');

                    if (tableName === "Shop_Consumables" && colDef.header === "Stock_Level") {

                        // 1. RUGGED RESET: Clear all existing locks in this row
                        const siloHeaders = ["Unit Cost", "Stock_Count", "Bulk_Value", "Reorder Point"];
                        siloHeaders.forEach(h => {
                            const idx = sheetConfig.columns.findIndex(c => c.header === h);
                            const target = row.querySelector(`td[data-col-index="${idx}"]`);
                            if (target) {
                                target.classList.remove('silo-locked');
                                target.style.opacity = "1";
                                target.style.pointerEvents = "auto";
                            }
                        });

                        const isCounted = newVal === "Counted";
                        const isBulk = ["None", "Few", "Adequate", "Many"].includes(newVal);
                        const wasCounted = oldVal === "Counted";
                        const wasBulk = ["None", "Few", "Adequate", "Many"].includes(oldVal);

                        // 2. NUCLEAR WIPE GUARD: Logic for pivoting between Counted and Bulk
                        if ((wasCounted && isBulk) || (wasBulk && isCounted)) {
                            const proceed = confirm(`MAE SYSTEM ALERT: Methodology Switch!\n\nThis will WIPE the abandoned silo data to ensure valuation integrity. Proceed?`);

                            if (!proceed) {
                                select.value = oldVal; 
                                finishEdit();
                                return;
                            }

                            const unitIndices = ["Unit Cost", "Stock_Count", "Reorder Point"].map(h => sheetConfig.columns.findIndex(c => c.header === h));
                            const bulkIndices = ["Bulk_Value"].map(h => sheetConfig.columns.findIndex(c => c.header === h));
                                          
                            const activeIndices = isCounted ? unitIndices : bulkIndices;
                            const inactiveIndices = isCounted ? bulkIndices : unitIndices;

                            // A. LOCK AND WIPE the inactive side
                            inactiveIndices.forEach(idx => {
                                const targetCell = row.querySelector(`td[data-col-index="${idx}"]`);
                                if (targetCell) {
                                    targetCell.innerHTML = ""; // Clear data for OneDrive sync
                                    targetCell.classList.add('silo-locked');
                                    targetCell.style.pointerEvents = "none";
                                    targetCell.style.opacity = "0.5";
                                }
                            });

                            // B. RE-ACTIVATE the active side with fresh Input boxes
                            activeIndices.forEach(idx => {
                                const targetCell = row.querySelector(`td[data-col-index="${idx}"]`);
                                if (targetCell) {
                                    const colDef = sheetConfig.columns[idx];
                                    const isCurrency = colDef.format && colDef.format.includes("$");

                                 targetCell.classList.remove('silo-locked');
                                    targetCell.style.pointerEvents = "auto";
        
                                    targetCell.innerHTML = `
                                        <input type="number" 
                                            class="edit-number-input" 
                                            value="0" 
                                            step="${isCurrency ? '0.01' : '1'}" 
                                            min="0">`;
                                }
                            });
                        }

                        // 3. FAT-FINGER RE-ENTRY: Maintain interactivity for the correct silo
                        // Now includes Reorder Point in the active unit silo
                        const activeHeaders = isCounted ? ["Unit Cost", "Stock_Count", "Reorder Point"] : ["Bulk_Value"];
                        const inactiveHeaders = isCounted ? ["Bulk_Value"] : ["Unit Cost", "Stock_Count", "Reorder Point"];

                        activeHeaders.forEach(h => {
                            const idx = sheetConfig.columns.findIndex(c => c.header === h);
                            const target = row.querySelector(`td[data-col-index="${idx}"]`);
                            if (target) {
                                target.style.pointerEvents = "auto";
                                target.style.opacity = "1";
                                target.classList.remove('silo-locked');
                                if (sheetConfig.columns[idx].type !== "dropdown") {
                                    target.contentEditable = "true";
                                }
                            }
                        });

                        inactiveHeaders.forEach(h => {
                            const idx = sheetConfig.columns.findIndex(c => c.header === h);
                            const target = row.querySelector(`td[data-col-index="${idx}"]`);
                            if (target) {
                                target.style.pointerEvents = "none";
                                target.style.opacity = "0.5";
                                target.classList.add('silo-locked');
                                target.contentEditable = "false";
                            }
                        });
                    }

                    finishEdit();

                    // 4. SYNC LOCKS
                    if (typeof UI.refreshSiloLocks === "function") {
                        UI.refreshSiloLocks(row, tableName, sheetConfig);
                    }
                };

                select.onkeydown = (k) => { if(k.key === 'Enter') finishEdit(); };
            };

            cell.onclick = startDropdownEdit;
        }
        // --- BRANCH 5: STANDARD TEXT ---
        else {
            cell.contentEditable = "true";
            cell.classList.add("text-edit-focus"); 
        }
    }); 

    setTimeout(() => {
        document.addEventListener('mousedown', globalClickOffHandler);
    }, 150);
}

// ====== END handleEditClick function =================

//========Helper function to trigger "Nuclear Wipe" warning before clearing data
function confirmMethodologySwitch(oldMethod, newMethod) {
    const from = oldMethod === "Counted" ? "UNIT-BASED (Count)" : "BULK-BASED (Subjective)";
    const to = newMethod === "Counted" ? "UNIT-BASED (Count)" : "BULK-BASED (Subjective)";
    
    return confirm(
        `MAE SYSTEM ALERT: Methodology Switch Detected!\n\n` +
        `Changing from ${from} to ${to}.\n\n` +
        `This will permanently WIPE the data in the abandoned silo to ensure system integrity. Proceed?`
    );
}
//====== END Helper function to trigger "Nuclear Wipe" warning before clearing data


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

        // Location ID: ensure defaults to "TBD" if empty
        if (col.header === "Location_ID" && (!input || input.value === "")) {
            return "TBD"; // Rapid entry fallback
        }

        // Return trimmed string for clean Excel data
        return input.value.trim();
    });

    // 🌟 MAE ENGINE RUGGED FIXED APPARATUS: ENFORCE TAG_ID DISCIPLINE 🌟
    const tagColumnIndex = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
    if (tagColumnIndex !== -1) {
        const rawTagVal = rowData[tagColumnIndex];
        // If the Tag_ID is empty, null, or whitespace, force-inject the absolute string token "UNTAGGED"
        if (rawTagVal === "" || rawTagVal === null || (typeof rawTagVal === 'string' && rawTagVal.trim() === "")) {
            rowData[tagColumnIndex] = "UNTAGGED";
            console.log(`MAE Intake Engine: Blank tag detected. Enforcing fallback default token: [UNTAGGED]`);
        }
    }

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
async function processInPlaceTableUpdate(tableName, preCapturedUpdates = null) {
    const table = document.getElementById("main-data-table");
    if (!table) return;
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    const title = document.getElementById("current-view-title");

    // RUGGED: Use the pre-captured snapshot. If none, run a fresh harvest.
    const updates = preCapturedUpdates || harvestTableData(tableName);

    // 1. SAFETY CHECK: If no rows were changed or found, just exit
    if (!updates || updates.length === 0) {
        console.log("MAE System: No data found to sync.");
        UI.exitEditMode();
        return;
    }

    // 2. BATCHING & CHUNKING LOGIC
    try {
        const token = await window.getGraphToken();
        const chunkSize = 20; 
        const totalRows = updates.length;
        
        for (let i = 0; i < totalRows; i += chunkSize) {
            const chunk = updates.slice(i, i + chunkSize);
            
            // UI Progress Update
            const percent = Math.round((i / totalRows) * 100);
            if (title) title.innerText = `💾 Syncing to OneDrive: ${percent}%...`;

            // Prepare the Batch Request
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

            // Wait 500ms between chunks to respect OneDrive's locking
            await new Promise(r => setTimeout(r, 500));
        }

        if (title) title.innerText = "✅ Sync Complete";
        console.log("MAE System: Batch sync successful.");
        
        // Final settle delay before UI cleanup
        await new Promise(r => setTimeout(r, 600));

    } catch (err) {
        console.error("Batch Sync Error:", err);
        UI.showError("Failed to sync changes. Check connection.");
    } finally {
        // Clean the UI inputs
        UI.exitEditMode();

        // Restore original title
        const sheetConfigFinal = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        if (title && sheetConfigFinal) {
            title.innerText = `View: ${sheetConfigFinal.tabName}`;
        }

        // Verification Refresh from Ground Truth (OneDrive)
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
    if (!table || table.classList.contains("is-quick-updating")) return;

    window.currentTable = tableName; 
    window.isEditing = true; 
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    
    table.classList.add("is-editing", "is-quick-updating");

    // 1. THE RUGGED WHITELIST (Matches Print Manual Log exactly)
    const keepHeaders = ["location_id", "machine name", "tool name", "item name", "stock_count", "quantity"];
    
    const headers = table.querySelectorAll("thead th");
    const visibleIndices = [];

    // 2. HEADER GOVERNANCE: Determine what to show
    headers.forEach((th, idx) => {
        const headerText = th.innerText.trim().toLowerCase();
        const isAction = th.classList.contains("edit-only-cell");
        const isKeep = keepHeaders.some(h => headerText.includes(h)) || isAction;

        if (!isKeep) {
            th.style.display = "none";
        } else {
            visibleIndices.push(idx);
        }
    });

    // 3. ROW GOVERNANCE: Hide non-whitelisted cells and enable Stock_Count input
    const rows = table.querySelectorAll("tbody tr");
    rows.forEach(row => {
        const cells = row.querySelectorAll("td");
        cells.forEach((cell, idx) => {
            if (!visibleIndices.includes(idx)) {
                cell.style.display = "none";
                return;
            }

            const colIdxAttr = cell.getAttribute('data-col-index');
            if (colIdxAttr === null) return; 

            const colDef = sheetConfig.columns[parseInt(colIdxAttr)];
            const isStockField = (colDef.header === "Stock_Count" || colDef.header === "Quantity");

            if (isStockField) {
                cell.classList.add("quick-edit-focus");
                const currentVal = parseInt(cell.innerText.replace(/[^0-9.-]+/g, "")) || 0;
                cell.innerHTML = `<input type="number" class="edit-number-input" value="${currentVal}" step="1" min="0">`;
                
                const input = cell.querySelector('input');
                input.onblur = () => { cell.innerText = input.value; };
                input.onkeydown = (e) => {
                    if (["e", ".", ","].includes(e.key.toLowerCase())) e.preventDefault();
                };
            } else {
                // Read-only reference fields (Name, Location)
                cell.style.backgroundColor = "#f4f4f4";
                cell.style.color = "#666";
            }
        });
    });

    setTimeout(() => {
        document.addEventListener('mousedown', globalClickOffHandler);
    }, 150);

    UI.renderCommandBar(tableName);
}

// ========END handleQuickUpdate ============

// =========================================================================
// =========== UNIFIED UNIVERSAL SCANNER LOOKUP (Item_Description AXIS) ===========
// =========================================================================
async function handleUniversalLookup(scannedId) {
    // 🌟 TRANSACTION CIRCUIT BREAKER REGISTRATION 🌟
    // Generate an authoritative token for this specific asynchronous network thread pass
    const currentTransactionId = Date.now();
    window.activeScanTransactionId = currentTransactionId;
    
    
    // 1. Enforce strict token discipline by cleaning the hardware payload string
    const cleanId = window.Labels.extractCleanId(scannedId).toString().trim().toUpperCase();
    if (!cleanId || cleanId === "" || cleanId === "UNTAGGED") {
        console.warn("MAE Security Guard: Aborting universal evaluation loop for null or unassigned tokens.");
        return;
    }
    
    console.log("MAE Scanner System: Intercepted hardware string payload:", cleanId);

    // 🌟 MAE ENGINE RUGGED FIXED APPARATUS: SHIELD RESTORATION GATE (WITH LEAKAGE SCRUBBER) 🌟
    // Locate any input elements that were frozen by labels.js during the capture phase
    const frozenInput = document.querySelector('input[disabled], select[disabled]');
    if (frozenInput && frozenInput.id !== "field-Tag_ID" && frozenInput.id !== "mae-bulk-container-input") {
        console.warn(`MAE Focus Protection: Reclaiming frozen field [${frozenInput.id}] following scan intercept.`);
        
        // 1. Re-enable the form field so it is fully interactive again
        frozenInput.disabled = false;
        
        // 2. 🌟 THE CLEANSE: Restore the pristine human data snapshot, wiping out leaked "htt" characters
        const originalValue = frozenInput.getAttribute("data-pre-scan-value") || "";
        frozenInput.value = originalValue;
        
        // 3. Alert the user of the focus layout conflict
        alert(`MAE INPUT FIELD PROTECTION:\n\nYour cursor is sitting in the wrong field line item box.\n\nTo prevent data corruption, this scan has been aborted. Please click directly into the "Tag_ID" field to scan hardware labels.`);
        
        // 4. Safely return focus so the worker can continue their data entry seamlessly
        setTimeout(() => { if (frozenInput) frozenInput.focus(); }, 50);
        return; // Terminate background lookup immediately to save ledger values from pollution
    }

    // --- STANDARD CROSS-TABLE SEARCH PIPELINE PROCEEDS (Cursor is safe or field is un-focused) ---
    const priorityTables = ["Resell_Inventory", "Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables"];
    let matchedAssetRowsList = [];
    let matchingTableName = "";
    let sheetConfigMatch = null;
    let foundTagType = "";

    try {
        for (const table of priorityTables) {
            // 🌟 CIRCUIT BREAKER CHECK (Inside Loop) 🌟
            // If the user reset the form while fetching, kill this thread immediately
            if (window.activeScanTransactionId !== currentTransactionId) {
                console.warn("MAE Circuit Breaker: Form reset detected mid-fetch. Aborting loop.");
                return;
            }

            const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === table);
            if (!sheetConfig) continue;

            const rowsData = await window.Dashboard.getFullTableData(table);
            if (!rowsData || rowsData.length === 0) continue;

            // 🌟 MAE ENGINE RUGGED FIXED APPARATUS: PARSE VIA NAME-BASED ORM OBJECT TRANSLATOR 🌟
            const normalizedRecords = window.Dashboard.transformRowsToObjects(rowsData, sheetConfig);
            
            // Filter all database records matching our sanitized barcode identifier string
            const internalMatches = normalizedRecords.filter(record => {
                const tagValue = record.data["Tag_ID"];
                return tagValue && String(tagValue).trim().toUpperCase() === cleanId;
            });

            if (internalMatches.length > 0) {
                // Map the structured database objects back into Graph-style row format for downstream UI compliance
                matchedAssetRowsList = internalMatches.map(rec => {
                    const rawValuesArray = sheetConfig.columns.map(col => rec.data[col.header] ?? null);
                    return { index: rec.index, id: rec.id, values: [rawValuesArray] };
                });
                
                matchingTableName = table;
                sheetConfigMatch = sheetConfig;
                
                // Read the tag type using the explicit name property key dictionary
                foundTagType = String(internalMatches[0].data["Tag_Type"] || "UNIQUE").trim().toUpperCase();
                break; // Escape search cycle loop early
            }
        }
        // 🌟 CIRCUIT BREAKER CHECK (Post-Loop) 🌟
        // Verify system state alignment before modifying any UI layouts or panels
        if (window.activeScanTransactionId !== currentTransactionId) {
            console.warn("MAE Circuit Breaker: Scan aborted by user reset. Blocking UI redirection.");
            return;
        }


        // --- SUB-ROUTINE A: THE INTAKE GUARDRAIL (Form Panel Mode Active) ---    
        const formPanelActive = document.getElementById("entry-form") !== null;
        const isRegistrationMode = window.currentTable === "inventory_registration" || window.currentTable === "inventory_search" || formPanelActive;
        
        if (isRegistrationMode) {
            console.log("MAE Safety Guard: Form active. Forwarding to Intake Modal Handler.");
            
            const tableSelect = document.getElementById("mae-central-table-selector");
            const targetTable = tableSelect ? tableSelect.value : window.currentTable;
            const tagInputFieldBox = document.getElementById("field-Tag_ID");
            
            // If we are actively inside the Registration Wizard layout, update the input field box
            if (tagInputFieldBox && window.currentTable === "inventory_registration") {
                tagInputFieldBox.value = cleanId;
            }

            if (matchedAssetRowsList.length === 0) {
                // 🌟 FRESH UNIQUE TAG: Forward straight to the custom modal categories picker window
                if (tableSelect && tableSelect.value !== "") {
                    window.UI.renderTagTypeWizardModal(targetTable, cleanId, currentTransactionId);
                } else {
                    alert(`Fresh Tag Intercepted: ${cleanId}\n\nPlease select a Target Table category dropdown row first.`);
                    if (tagInputFieldBox) tagInputFieldBox.value = "";
                }
                return;
            }

            if (foundTagType === "UNIQUE") {
                if (tagInputFieldBox) {
                    tagInputFieldBox.value = "";
                    tagInputFieldBox.style.borderColor = "#e74c3c";
                    tagInputFieldBox.style.backgroundColor = "#fadbd8";
                }
                const nameKey = sheetConfigMatch.columns.some(c => c.header === "Item_Description") ? "Item_Description" : "Item Name";
                const matchedItemLabelName = matchedAssetRowsList[0].values[0][sheetConfigMatch.columns.findIndex(c => c.header === nameKey)] || "Unknown Asset";
                
                alert(`CRITICAL CONFLICT ERROR:\n\nThis barcode is ALREADY permanently mapped to an individual unique asset:\n\n• Table Category: ${sheetConfigMatch.tabName}\n• Existing Item: "${matchedItemLabelName}"\n\nTo preserve database ledger integrity, you CANNOT reuse this tag here.`);
                return;
            }

            if (foundTagType === "MULTIPLE") {
                // 🌟 FORWARD CONTAINER DETECTIONS CLEANLY THROUGH THE CUSTOM WINDOW
                window.UI.renderTagTypeWizardModal(targetTable, cleanId, currentTransactionId);
                return;
            }
        }

        // --- SUB-ROUTINE B: STANDARD NAVIGATION LOOKUP (No active panels on screen) ---
        window.UI.showLoading(`Sweeping Shop Records for: ${cleanId}...`);

        if (matchedAssetRowsList.length === 0) {
            window.UI.showError(`
                <div style="text-align:center; padding: 20px;">
                    <h3 style="color:var(--primary);">New Tag Detected</h3>
                    <div style="font-size: 1.5rem; font-weight: 800; background: #fffde7; border: 2px dashed var(--accent); padding: 15px; margin: 10px 0;">
                        ID: ${cleanId}
                    </div>
                    <p>This ID is not recognized in the active shop inventory lists.</p>
                    <button class="action-btn" onclick="window.handleAddClickWithId('${cleanId}')" style="width:100%; margin-bottom:10px;">
                        ➕ Register New Item
                    </button>
                    <button class="action-btn cancel-btn" onclick="window.loadTableData('Master_Dashboard')" style="width:100%;">
                        Discard Scan
                    </button>
                </div>
            `);
            return;
        }

        // ROUTE B1: REDIRECT SINGLE UNIQUE ASSETS TO CARD VIEW
        if (foundTagType === "UNIQUE") {
            window.currentTable = matchingTableName;
            const singleFlattenedCells = matchedAssetRowsList.values;
            window.UI.renderScanResultCard(singleFlattenedCells, matchingTableName, sheetConfigMatch, matchedAssetRowsList.index);
            return;
        }

        // ROUTE B2: REDIRECT MULTIPLE CONTAINER ENTRIES TO SEARCH HUB GRID VIEW
        if (foundTagType === "MULTIPLE") {
            window.currentTable = matchingTableName;
            const nameIdx = sheetConfigMatch.columns.findIndex(c => c.header === "Item_Description");
            const locationColumnIdx = sheetConfigMatch.columns.findIndex(c => c.header === "Location_ID");
            
            const containerBundleResults = matchedAssetRowsList.map(rowObj => {
                const cells = rowObj.values;
                const physicalLocation = cells[locationColumnIdx] || "TBD";
                return {
                    category: `${sheetConfigMatch.tabName} [Storage: ${physicalLocation}]`,
                    itemName: cells[nameIdx] || "N/A",
                    tableName: matchingTableName,
                    rowIndex: rowObj.index
                };
            });
            window.UI.renderVirtualSearchHub(containerBundleResults, cleanId, matchingTableName);
            return;
        }

    } catch (error) {
        console.error("MAE Security Module Ingest Failure:", error);
        window.UI.showError("Verification routine failed. Check network connection.");
    }
}
// =========== END UNIVERSAL SCANNER LOOKUP ===========






// ============ FUNCTION TO HANDLE SINGLE-ROW UPDATES===========
async function updateSingleRowFromForm(tableName, rowIndex, sheetConfig) {
    // 1. Gather data from the form fields
    const rowData = sheetConfig.columns.map((col, index) => {
        const fieldId = `field-${col.header.replace(/\s+/g, '')}`;
        const input = document.getElementById(fieldId);
        
        if (col.type === "formula") return null; // Don't overwrite formulas
        if (col.type === "boolean") return input ? (input.checked ? "TRUE" : "FALSE") : "FALSE";
        
        // MAE INTEGRITY FIX: If the field is an input, explicitly capture its data value 
        // even if it has a 'disabled' parameter active on screen
        if (input) {
            if (col.type === "number") {
                return input.value === "" ? null : parseFloat(input.value);
            }
            return input.value.trim();
        }
        return "";
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

//==========  END Getting Location Dependencies to support Deletion of Location_ID===

// =========================================================================
//  GLOBAL FOCUSED-FIELD SEARCH SCANNER ENGINE 
// =========================================================================
async function executeFocusedAssetSearch(searchQuery) {
    const cleanQuery = searchQuery.toString().trim().toLowerCase();
    console.log("MAE Search Engine: Sweeping memory cache records for lookup:", cleanQuery);

    UI.showLoading(`Scanning All Shop Records for: "${searchQuery}"...`);

    // Prioritized operational asset tables requiring focus field evaluation checks
    const tablesToSearch = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
    let consolidatedMatches = [];

    try {
        for (const tableName of tablesToSearch) {
            const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
            if (!sheetConfig) continue;

            // Fetch partition dataset array cleanly from your dashboard module
            const dataRows = await Dashboard.getFullTableData(tableName);
            if (!dataRows || dataRows.length === 0) continue;

            // Extract index references only for human-readable label rows (Enforces Focus-Field constraint)
            const targetedHeaderIndices = [];
            sheetConfig.columns.forEach((col, idx) => {
                const headerName = col.header.toLowerCase();
                if (headerName.includes("name") || 
                    headerName.includes("tool") || 
                    headerName.includes("machine") || 
                    headerName.includes("description") || 
                    headerName.includes("model") ||
                    headerName.includes("remarks") ||
                    headerName.includes("brand")) {
                    targetedHeaderIndices.push(idx);
                }
            });

            // Target column index vectors mapping your identifier labels and physical storage spot metrics
            const firstVisibleNameIdx = sheetConfig.columns.findIndex(c => !c.hidden && (c.header.includes("Name") || c.header.includes("Tool") || c.header.includes("Description")));
            const locationColumnIdx = sheetConfig.columns.findIndex(c => c.header === "Location_ID");

            dataRows.forEach(rowObj => {
                // Flatten structural array blocks cleanly matching your standard Graph API row mappings
                const cells = (rowObj.values && Array.isArray(rowObj.values)) 
                    ? rowObj.values[0] 
                    : (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values : rowObj;
                    
                if (!cells) return;

                let recordMatchesQuery = false;
                for (const colIdx of targetedHeaderIndices) {
                    const cellContents = String(cells[colIdx] || "").toLowerCase();
                    if (cellContents.includes(cleanQuery)) {
                        recordMatchesQuery = true;
                        break; // Substring verified, escape column check loop early for speed
                    }
                }

                if (recordMatchesQuery) {
                    // Extract the item's physical storage spot coordinate vector
                    const itemPhysicalLocation = cells[locationColumnIdx] || "TBD";
                    
                    consolidatedMatches.push({
                        category: `${sheetConfig.tabName} [Physical Storage Spot: ${itemPhysicalLocation}]`,
                        itemName: cells[firstVisibleNameIdx] || "N/A",
                        tableName: tableName
                    });
                }
            });
        }

        console.log(`MAE Search Engine: Sweeping complete. Discovered ${consolidatedMatches.length} matching rows.`);
        
        if (consolidatedMatches.length === 0) {
            UI.showError(`
                <div style="text-align:center; padding: 25px;">
                    <h3 style="color:var(--primary);">No Matches Located</h3>
                    <p>The system string search for <b>"${searchQuery}"</b> did not find any inventory records matching your identifier fields.</p>
                    <button class="action-btn" onclick="renderSearchControls()" style="width:100%; max-width:300px; margin-top:15px;">Run New Search</button>
                </div>
            `);
            return;
        }

        // Forward matching collections down to your Multiple Items visual Search Hub viewport
        UI.renderVirtualSearchHub(consolidatedMatches);

    } catch (err) {
        console.error("MAE Search Engine: Processor routine loop crashed.", err);
        UI.showError("Search matrix lookup processing failed. Check network link indicators.");
    }
}

//=======  END  GLOBAL FOCUSED-FIELD SEARCH SCANNER ENGINE ==========

// =========================================================================
// UNTAGGED ASSETS AUDIT AND REGISTRATION SUB-SYSTEM 
// =========================================================================
async function runUntaggedAudit() {
    UI.showLoading("Performing cross-table compliance sweep of UNTAGGED assets...");
    
    const inventoryTables = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
    let untaggedResultsList = [];

    try {
        for (const tableName of inventoryTables) {
            const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
            if (!sheetConfig) continue;

            const dataRows = await Dashboard.getFullTableData(tableName); 
            if (!dataRows || dataRows.length === 0) continue;

            const tagIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
            // MAE REPAIR: Clean lookup pointing directly to your new standardized text field axis
            const nameIdx = sheetConfig.columns.findIndex(c => c.header === "Item_Description");
            const idIdx = sheetConfig.columns.findIndex(c => c.header === "mae_id");

            if (tagIdx === -1 || nameIdx === -1) continue;

            dataRows.forEach(rowObj => {
                // 🌟 MAE ENGINE RUGGED APPARATUS FIX: Dig down inside Graph's 2D nested array cell container safely 🌟
                const cells = (rowObj.values && Array.isArray(rowObj.values[0])) 
                    ? rowObj.values[0] 
                    : (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values : null;

                if (cells && cells[tagIdx] !== undefined && cells[tagIdx] !== null) {
                    const currentTagTextValue = String(cells[tagIdx]).trim().toUpperCase();
                    
                    // Match rows containing your strict default fallback code string "UNTAGGED"
                    if (currentTagTextValue === "UNTAGGED" || currentTagTextValue === "") {
                        // Extract unique row identification key safely
                        const rowUniqueMaeId = idIdx !== -1 ? cells[idIdx] : `ROW-${rowObj.index}`;

                        untaggedResultsList.push({
                            category: sheetConfig.tabName,
                            itemName: cells[nameIdx] || "Unnamed Asset Row",
                            mae_id: rowUniqueMaeId, // Extracted cell primary key value
                            tableName: tableName,
                            rowIndex: rowObj.index // Persistent row tracker index anchor
                        });
                    }
                }
            });
        }

        console.log(`MAE Audit Engine: Complete. Located ${untaggedResultsList.length} compliance gaps.`);
        
        // Forward matched arrays straight to your dedicated virtual table layout renderer inside ui.js
        UI.renderUntaggedAuditGrid(untaggedResultsList);

    } catch (error) {
        console.error("MAE Engine Untagged Cross-Table Audit Crashed:", error);
        UI.showError("Audit loop processing was interrupted. Ensure local connection is healthy.");
    }
}

//====== END UNTAGGED ASSETS AUDIT AND REGISTRATION SUB-SYSTEM ==========

// =========================================================================
//  MAE REGISTERED TAG DISCIPLINE: DECOMMISSION AND REPLACEMENT SUB-SYSTEM
// =========================================================================
async function initiateTagReplacementWorkflow(tableName, rowIndex, oldTagId) {
    const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    const tagTypeIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_Type");
    const tagIdIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
    
    // Fetch row directly to read its type state fact reliably
    const rowsData = await Dashboard.getFullTableData(tableName);
    const targetRow = rowsData.find(r => parseInt(r.index, 10) === parseInt(rowIndex, 10));
    const targetRowCells = (targetRow.values && Array.isArray(targetRow.values)) ? targetRow.values : targetRow.values;
    
    const isMultipleContainer = targetRowCells && tagTypeIdx !== -1 && String(targetRowCells[tagTypeIdx]).trim().toUpperCase() === "MULTIPLE";

    let warningPrompt = `MAE MAINTENANCE SERVICE NOTICE:\n\nYou are altering the physical tag attachment for Tag_ID: [${oldTagId}].\n\n`;
    if (isMultipleContainer) {
        warningPrompt += `⚠️ ATTENTION: This label represents a MULTIPLE item container space. Changing or removing it will apply to ALL items bundled within this spot in the spreadsheet ledger.\n\n`;
    }
    warningPrompt += `Select an action:\n\n• Click OK to mark this asset record as "UNTAGGED" and route it straight to the compliance audit workspace queue.\n• Click CANCEL to abort operation completely.`;

    const proceedToUntagged = confirm(warningPrompt);
    if (!proceedToUntagged) {
        console.log("MAE System: Tag replacement aborted by user choice.");
        return;
    }

    UI.showLoading("Decommissioning active sticker identifier from ledger...");

    try {
        // Prepare Sparse Update Row Array
        const rowValues = new Array(sheetConfig.columns.length).fill(null);
        
        // Inject absolute tracking tokens
        rowValues[tagIdIdx] = "UNTAGGED";
        
        // Standardize tag type back to safe UNIQUE initial default state since the group is broken
        if (tagTypeIdx !== -1) rowValues[tagTypeIdx] = "UNIQUE";

        const token = await window.getGraphToken();
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(maeSystemConfig.spreadsheetName)}:/workbook/tables/${tableName}/rows/itemAt(index=${parseInt(rowIndex, 10)})`;
        const response = await fetch(url, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowValues] })
        });

        if (response.ok) {
            console.log(`MAE Ledger Update Success: Item row index ${rowIndex} reset to UNTAGGED status.`);
            alert("System Alignment Complete!\n\nThis asset has been stripped of its previous ID tracking data. Please proceed to 'Audit Untagged Items' on the dashboard to register your replacement label sticker roll.");
            
            // Redirect smoothly straight back to dashboard workspace panels
            window.loadTableData("Master_Dashboard");
        } else {
            const errData = await response.json();
            throw new Error(errData.error.message || "OneDrive network timeout.");
        }
    } catch (err) {
        console.error("MAE Tag Decommission Routine Crash:", err);
        UI.showError("Failed to detach tracking parameters. Verify cloud asset links.");
    }
}

//===== END  MAE REGISTERED TAG DISCIPLINE: DECOMMISSION AND REPLACEMENT SUB-SYSTEM ==========

//  ====Registered tag discipline: Batch Container Group Update Engine
async function executeBulkContainerGroupingTransition(rawScannerInput) {
    if (!rawScannerInput || rawScannerInput.trim() === "") return;

    // 1. URL SANITIZATION & SCRUBBING
    // Processes the raw string to strip away ToteScan web parameters, keeping just the clean token
    const cleanTagId = window.Labels.extractCleanId(rawScannerInput).toUpperCase();
    console.log("MAE Bulk Engine: Initializing container build for Tag:", cleanTagId);

    const bulkInputField = document.getElementById("mae-bulk-container-input")
    if (bulkInputField) bulkInputField.value = cleanTagId;

    // 2. ANTI-COLLISION SAFETY BLOCK
    // Check if the scanned sticker is already used as a unique tool or a different container
    const isTagAlreadyInUse = await verifyTagUniquenessCrossTable(cleanTagId);
    if (isTagAlreadyInUse) {
        alert(`CRITICAL ASSIGNMENT CONFLICT BLOCKED:\n\nYou cannot build this group. Barcode [${cleanTagId}] is already actively assigned to an asset record inside your database ledger.\n\nGrab a completely fresh, unused sticker.`);
        if (bulkInputField) {
            bulkInputField.value = "";
            bulkInputField.focus();
        }
        return;
    }

    // 3. DATA HARVESTING VIA THE CHECKBOXES
    const checkedBoxes = document.querySelectorAll('.mae-audit-bulk-checkbox:checked');
    if (checkedBoxes.length === 0) {
        alert("MAE System Exception: Zero rows are currently selected for grouping.");
        return;
    }

    const confirmation = confirm(`MAE COMPLIANCE REGULATION CONFIRMATION:\n\nYou are about to group ${checkedBoxes.length} loose items onto a shared container tag:\n\n• New Shared Tag ID: [${cleanTagId}]\n• Tag Type Assignment: MULTIPLE\n\nThis will link them in the ledger while leaving all individual descriptions, costs, and remarks completely intact. Proceed?`);
    if (!confirmation) {
        if (bulkInputField) bulkInputField.value = "";
        return;
    }

    UI.showLoading(`LOCKING SYSTEM: Packaging ${checkedBoxes.length} items into container group...`);

    try {
        const token = await window.getGraphToken();
        
        // 4. TRANSACTION EXECUTION LOOP (THROTTLED PER ROW)
        for (const box of checkedBoxes) {
            const tableName = box.getAttribute('data-table');
            const rowIndex = parseInt(box.getAttribute('data-row'), 10);
            const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);

            if (!sheetConfig) continue;

            const tagIdIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
            const tagTypeIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_Type");

            // Sparse update structure: Null values guarantee existing item rows remain completely untouched
            const rowValues = new Array(sheetConfig.columns.length).fill(null);
            rowValues[tagIdIdx] = cleanTagId;
            
            if (tagTypeIdx !== -1) {
                rowValues[tagTypeIdx] = "MULTIPLE"; // Enforce the shared container constraint fact
            }

            console.log(`MAE Bulk Sync: Updating row index ${rowIndex} inside ${tableName} to Container Tag [${cleanTagId}]`);

            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(window.maeSystemConfig.spreadsheetName)}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;

            const response = await fetch(url, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: [rowValues] })
            });

            if (!response.ok) {
                console.error(`MAE Fault Intercept: Network update failed on row index ${rowIndex} inside ${tableName}`);
            }

            // 400ms Throttling protection delay to safeguard Microsoft Graph workbook concurrency locks
            await new Promise(r => setTimeout(r, 400));
        }

        alert(`System Integrity Verified!\n\nSuccessfully bundled ${checkedBoxes.length} items onto Container Tag: [${cleanTagId}]. All unique descriptions have been preserved.`);
        
        UI.showLoading("Synchronizing local ledger partitions... please wait.");
        
        // 🌟 INDUSTRIAL SETTLE DELAY: Allows OneDrive file properties to completely settle on Microsoft's servers
        await new Promise(r => setTimeout(r, 1200));

        UI.clearBulkAuditSelection();
        await window.runUntaggedAudit();
        
        // Clear loading overlay and bounce view back to master layout cleanly
        window.loadTableData("Master_Dashboard");

    } catch (err) {
        console.error("MAE Bulk Grouping Sub-System Crash:", err);
        UI.showError("Failed to cleanly assemble container group. Verify cloud asset links.");
    }
}
//==== END Registered tag discipline: Batch Container Group Update Engine ========

// ==== MAE REGISTERED TAG DISCIPLINE: UTILITY COMPLIANCE CROSS-TABLE SCANNER
// == Name-Based Utility Compliance Cross-Table Scanner
// =========================================================================
async function verifyTagUniquenessCrossTable(tagIdToCheck) {
    if (!tagIdToCheck || tagIdToCheck.toString().trim() === "" || tagIdToCheck.toString().toUpperCase() === "UNTAGGED") {
        return false; // The absolute fallback string 'UNTAGGED' is skipped to allow grouping audits to process cleanly
    }
    const targetSearchToken = tagIdToCheck.toString().trim().toUpperCase();
    const priorityTables = ["Resell_Inventory", "Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables"];

    for (const table of priorityTables) {
        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === table);
        if (!sheetConfig) continue;

        const rawRowsData = await window.Dashboard.getFullTableData(table);
        if (!rawRowsData || rawRowsData.length === 0) continue;

        // 🌟 MAE ENGINE RUGGED FIXED APPARATUS: PARSE RAW MATRIX TO KEYED DATABASE RECORDS 🌟
        const normalizedRecords = window.Dashboard.transformRowsToObjects(rawRowsData, sheetConfig);

        // Scan across your database records checking tracking attributes via explicit text name dictionary values
        const matchFound = normalizedRecords.find(record => {
            const recordTagValue = record.data["Tag_ID"];
            return recordTagValue && String(recordTagValue).trim().toUpperCase() === targetSearchToken;
        });

        if (matchFound) {
            console.warn(`MAE Security Matrix Intercept: Label collision detected inside sheet classification [${sheetConfig.tabName}] at row index ${matchFound.index}`);
            return true; // Collision verified! This tag is already in use by an active asset record row.
        }
    }
    console.log(`MAE Security Matrix Verification: Label token [${targetSearchToken}] confirmed safe and unique.`);
    return false; // Safe and clean! The tag is verified unique and ready for assignment.
}

// ===== END MAE REGISTERED TAG DISCIPLINE: UTILITY COMPLIANCE CROSS-TABLE SCANNER ========



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
window.excelSerialToDate = excelSerialToDate;
window.renderLocationInspectorControls = UI.renderLocationInspectorControls;
window.executeLocationInspection = UI.executeLocationInspection;
window.printInspectedLocationTable = UI.printInspectedLocationTable;
window.executeFocusedAssetSearch = executeFocusedAssetSearch;
window.renderSearchControls = UI.renderSearchControls;
window.triggerAssetSearch = UI.triggerAssetSearch;
window.runUntaggedAudit = runUntaggedAudit;
window.renderCentralRegistrationWizard = UI.renderCentralRegistrationWizard;
window.renderUntaggedAuditGrid = UI.renderUntaggedAuditGrid;
window.submitNewRow = submitNewRow;
window.initiateTagReplacementWorkflow = initiateTagReplacementWorkflow;
window.renderTagMaintenanceWizard = UI.renderTagMaintenanceWizard;
window.executeMaintenanceSearch = UI.executeMaintenanceSearch;
window.executeDirectTagWipe = UI.executeDirectTagWipe;
window.evaluateAuditCheckboxStateChanges = UI.evaluateAuditCheckboxStateChanges;
window.clearBulkAuditSelection = UI.clearBulkAuditSelection;
window.executeBulkContainerGroupingTransition = executeBulkContainerGroupingTransition;
window.verifyTagUniquenessCrossTable = verifyTagUniquenessCrossTable;

startup();