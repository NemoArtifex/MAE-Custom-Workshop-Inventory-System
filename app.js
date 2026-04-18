window.maeLocations =["TBD"]; // Global cache default for intake workflow
import { maeSystemConfig } from './config.js'
import { UI} from './ui.js';
import { Labels } from './labels.js';
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

export let myMSALObj;
let account = null;
async function startup() {
    try {
        // Initialize the PublicClientApplication
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
            // SCENARIO 1: USER IS LOGGED IN
            updateUIForLoggedInUser(account); 

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

    loadTableData("Master_Dashboard");
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

    if (btn.id === 'btn-add') {
        handleAddClick(currentTable); 
    } 
    else if (btn.id === 'btn-edit') {
        handleEditClick(currentTable);
    } 
    
    //  Button scan Logic
    else if (btn.id === 'btn-scan') {
        // Call your new Labels module
        Labels.startScanner((cleanId) => {
            // Once scanned, run the lookup
            handleUniversalLookup(cleanId);
        });
    }

    // CONSOLIDATED PRINT LOGIC: Handles both Table and Manual Log
    else if (btn.id === 'btn-print' || btn.id === 'btn-manual-print') {
        const sheetConfig = config.worksheets.find(s => s.tableName === currentTable);
        const titleElement = document.getElementById("current-view-title");
        
        // RUGGED: Extract only the title text (ignores the "Back" button)
        const spanElement = titleElement.querySelector("span");
        const currentTitleText = spanElement ? spanElement.innerText : titleElement.innerText;

        // DATE GENERATION: MM/DD/YYYY
        const today = new Date().toLocaleDateString('en-US');

        // Logic-Based Title Override
        let finalPrintTitle = `${currentTitleText} (as of ${today})`;

        // Branch to appropriate UI function
        if (btn.id === 'btn-print') {
            UI.printTable(currentTable, sheetConfig, finalPrintTitle);
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
// Centralized "Click-Off" handler for all edit modes
async function globalClickOffHandler(e) {
    const table = document.getElementById("main-data-table");
    const container = document.getElementById("table-container");// scrollbar area
    const title = document.getElementById("current-view-title");
    if (!table || !title) return;
 
    // SCROLLBAR DETECTION 
    // If the click is inside the container but the clientX is in the scrollbar gutter
    const rect = container.getBoundingClientRect();
    const isScrollbarClick = 
        (e.clientX > rect.left + container.clientWidth) || 
        (e.clientY > rect.top + container.clientHeight);


    // 1. IDENTIFY SAFE ZONES
    const isInsideTable = table.contains(e.target);
    const isCommandBtn = e.target.closest('.action-btn');
    const isDeleteBtn = e.target.closest('.delete-row-btn');
    const isEntryForm = e.target.closest('#entry-form');

    // 2. TRIGGER SYNC ONLY ON BACKGROUND CLICK
    if (!isInsideTable && !isCommandBtn && !isDeleteBtn &&!isScrollbarClick && !isEntryForm) {
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
    
            // If "Number" is chosen, send as Integer. Else, send the string label.
            return (select.value === "Number") ? (parseInt(numInput.value) || 0) : select.value;
        }

        // Return trimmed string for clean Excel data
        return input.value.trim();
    });



    try {
        // 2. AUTH: Get fresh token
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        // 3. API CALL: Corrected URL path for Table Rows
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows`;
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
    const rows = table.querySelectorAll("tbody tr:not(.repair-group-header tr)"); // Skip sub-headers
    
    const updates = [];

    rows.forEach(tr => {
        const rowIndex = tr.getAttribute("data-row-index");
        const rowValues = [];

        sheetConfig.columns.forEach((col, index) => {
            // 1. RUGGED PROTECTION: Identify Formulas
            if (col.type === "formula") {
                rowValues.push(null);
                return;
            }

            // 2. PRIMARY KEY INTEGRITY: Pull mae_id from attribute
            if (col.header === "mae_id") {
                const anchoredId = tr.getAttribute('data-mae-id');
                rowValues.push(anchoredId || ""); 
                return;
            }

            // Find the specific cell for this column
            const cell = tr.querySelector(`td[data-col-index="${index}"]`);
            if (!cell) {
                const isNumeric = col.type === "number" || col.type === "date";
                rowValues.push(isNumeric ? null : "");
                return;
            }

            // 3. DATA EXTRACTION: Identify the UI element and grab the value
            let val = "";
            const select = cell.querySelector('select');
            const input = cell.querySelector('input[type="number"], input[type="text"]');
            const checkbox = cell.querySelector('input[type="checkbox"]');

            if (checkbox) {
                // Returns true or false as boolean values
                val = checkbox.checked;
            } else if (select) {
                val = select.value;
            } else if (input) {
                val = input.value;
            } else {
                // Fallback to plain text, stripping currency symbols for processing
                val = cell.innerText.replace(/[$,]/g, "").trim();
            }

            // 4. TYPE ENFORCEMENT
            if (col.type === "number") {
                const isCurrency = col.format && col.format.includes("$");
                let cleanNum = parseFloat(val.toString().replace(/[^0-9.-]+/g, ""));
                
                if (isNaN(cleanNum)) {
                    val = 0;
                } else {
                    val = isCurrency ? parseFloat(cleanNum.toFixed(2)) : Math.floor(cleanNum);
                }
            }
            rowValues.push(val);
        });
        updates.push({ index: rowIndex, values: [rowValues] });
    });

    // 5. SYNC TO MICROSOFT GRAPH
    try {
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        for (const update of updates) {
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${update.index})`;
                       
            const response = await fetch(url, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${tokenResponse.accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: update.values })
            });

            if (!response.ok) {
                const errorBody = await response.json();
                console.error("MAE System: Microsoft Graph Error during sync:", errorBody);
            }
        }
        console.log("MAE System: Primary Key Integrity maintained. All changes synced.");
    } catch (err) {
        console.error("Sync Error:", err);
        UI.showError("Failed to sync changes. Check internet connection.");
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
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
       
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
        // Get fresh token once before the loop to save on resources
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account // Uses the account variable defined in app.js
        });

        for (const tableName of tables) {
            const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
            
            // API path to the specific table
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows`;
            
            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
            });

            if (response.ok) {
                const data = await response.json();
                
                // Find row where Column 0 (mae_id) matches
                const matchedRow = data.value.find(row => {
                    const rowId = row.values[0][0]; 
                    return String(rowId).trim() === cleanId;
                });

                if (matchedRow) {
                    console.log(`MAE System: Match found in ${tableName} at index ${matchedRow.index}`);
                    window.currentTable = tableName;
                    // RUGGED TABLET UI: This replaces the 'Mobile' version
                    UI.renderScanResultCard(matchedRow.values[0], tableName, sheetConfig, matchedRow.index);
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

//========= ADD refreshLocationCache Location_ID SYNC===========
async function refreshLocationCache() {
    try {
        const data = await Dashboard.getFullTableData("Location");
        const locConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        const locIdx = locConfig.columns.findIndex(c => c.header === "Location_ID");

        if (data) {
            // Extract the IDs, filter nulls, and ensure "TBD" is the first option
            const list = data.map(row => row.values[0][locIdx]);
            window.maeLocations = ["TBD", ...new Set(list.filter(i => i && i !== "TBD"))];
            console.log("MAE System: Location Control Tower Synced.");
        }
    } catch (e) {
        console.warn("Location sync failed, using last known cache.");
    }
}

//====== END refreshLocationCache  Location_ID sync============

//======= submitNewLocationToTable : writes data to Excel
async function submitNewLocationToTable(locationId) {
    try {
        const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        
        // Build the row based on the Header names in config.js
        const newRow = sheetConfig.columns.map(col => {
            if (col.header === "mae_id") return `LOC-${Date.now()}`;
            if (col.header === "Location_ID") return locationId;
            if (col.header === "Description") return "New Location - Update in Location Tab";
            return ""; // For Type or Parent_Location
        });

        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
        //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
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
        console.error("MAE System: Failed to establish new location", err);
        return false;
    }
}

//===== END submitNewLocationToTable =============




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