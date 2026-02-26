import { maeSystemConfig } from './config.js'
const fileName = maeSystemConfig.spreadsheetName;
// =============CONFIGURATION: The "Blueprint"  ======================
// Defines the configuration object for the Microsoft Authentication Libray (MSAL)
// Used to integrate Microsoft's identity and sign-in features into web apps
const msalConfig = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://login.microsoftonline.com/common",
        //redirectUri: "http://localhost:5500" ,
        redirectUri: "https://nemoartifex.github.io/MAE-Custom-Workshop-Inventory-System/",
        // This helps the popup find the parent window in production
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
// the signIn function calls updateUIForLoggedInUser() if successful 'login'
// changes text on button and triggers loadDynamicMenu() function  
function updateUIForLoggedInUser(userAccount) {
    const authButton = document.getElementById("auth-btn");
    authButton.onclick = null;  //CLEAR: wipe any old "onclick" or "inline" handlers firs  
    console.log("Enabling the Connect button now...");
    authButton.disabled = false;
    authButton.innerText = `Sign Out: ${userAccount.username}`;
    authButton.style.background = "#c0392b"; // Change to red for "Sign Out"
    authButton.style.color = "white";
    authButton.removeEventListener("click", signIn); // Remove sign-in listener to prevent multiple logins
    authButton.addEventListener("click", signOut); // Add sign-out functionality for better UX
    console.log("Loading dynamic menu for user:", userAccount.username);
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
    const authButton = document.getElementById("auth-btn");
    authButton.innerText = "Connect to Microsoft";
    authButton.style.background = ""; // Resets to original CSS color
    authButton.style.color = ""; // Resets to original CSS color
    authButton.onclick = null; 
    // Clear event listeners correctly  
    authButton.removeEventListener("click", signOut); // Remove sign-out listener
    authButton.addEventListener("click", signIn);
    
    // Clear actual Data/UI elements; clear innerHTML so list items literally disappear  
    const menu = document.getElementById("menu");
    if (menu) menu.innerHTML = "";

    const container = document.getElementById("table-container");
    if (container) container.innerHTML = "";

    const title = document.getElementById("current-view-title");
    if (title) title.innerText = "Please connect to view inventory data.";

    console.log("UI Reset: Inventory Data Cleared.");   


}
//========END FUNCTION TO RESET UI AFTER SIGN-OUT =============

//======= FUNCTION Load Dynamic Menu ================
async function loadDynamicMenu() {
   const menu = document.getElementById("menu");
    menu.innerHTML = ""; // Clear any existing menu items
    
    console.log("Building dynamic menu from config...");
    // We iterate through the CONFIG, not the Excel file. 
    // This ensures the App stays "Locked" to the business agreement.
    maeSystemConfig.worksheets.forEach(sheet => {
        const btn = document.createElement("button");
        btn.innerText = sheet.tabName;
        btn.className = "menu-btn";
        btn.onclick = () => loadTableData(sheet.tableName);

        const listItem = document.createElement("li");
        listItem.appendChild(btn);
        menu.appendChild(listItem);
    });

    // Check if the actual Excel file exists on OneDrive
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

    //const fileName = maeSystemConfig.spreadsheetName;
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
        }
    } catch (error) {
        console.error("Verification Error:", error);
    }
}
//======END FUNCTION verifySpreadSheetExists =============

//=========FUNCTION createInitialWorkbook =============
/**
 * Step 2: Create the .xlsx file and build the Tables/Headers
 * from the maeSystemConfig.
 * 
 *  Prep work to create an empty .xlsx file. The Graph API requires a binary upload, 
 * so we create a minimal Excel file in-memory.
 * 
 */

// This generates a valid, minimal blank Excel file directly from bytes.
// No atob() or base64 string required.
function getBlankExcelBuffer() {
    const bytes = new Uint8Array([
        0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x5B, 0x43,
        0x6F, 0x6E, 0x74, 0x65, 0x6E, 0x74, 0x5F, 0x54, 0x79, 0x70, 0x65, 0x73, 0x5D, 0x2E, 0x78, 0x6D,
        0x6C, 0xAD, 0x4D, 0xCB, 0x0E, 0xC2, 0x30, 0x10, 0xBC, 0xF7, 0x22, 0xBE, 0x31, 0xAD, 0x2D, 0x2D,
        0x20, 0xD2, 0x02, 0x11, 0x91, 0x82, 0x14, 0x52, 0x41, 0x4C, 0x23, 0x31, 0x70, 0x19, 0xC1, 0xFE,
        0x7B, 0xD1, 0x36, 0x93, 0x92, 0x2B, 0xAE, 0x2F, 0xDF, 0x7C, 0x1F, 0x1F, 0x0B, 0xAE, 0xB2, 0x12,
        0x52, 0x01, 0x35, 0x5D, 0x1B, 0x45, 0x0E, 0x52, 0x79, 0x02, 0xD4, 0x71, 0x35, 0xB5, 0x22, 0xB2,
        0x01, 0x21, 0x83, 0x2B, 0xD4, 0x54, 0xD6, 0x72, 0x51, 0xD6, 0x7E, 0x32, 0xE5, 0x2A, 0x23, 0x1C,
        0x3B, 0x0F, 0x5C, 0x4C, 0x63, 0x0F, 0x00, 0x00, 0xFF, 0xFF, 0x50, 0x4B, 0x05, 0x06, 0x00, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x36, 0x00, 0x00, 0x00, 0x9D, 0x00, 0x00, 0x00, 0x00, 0x00
    ]);
    return bytes.buffer;
}

async function createInitialWorkbook(accessToken) {
    //const fileName = maeSystemConfig.spreadsheetName;
    const baseUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/content?@microsoft.graph.conflictBehavior=fail`;
    
    // 1. Create the empty Excel file
    const createRes = await fetch(baseUrl, {
        method: 'PUT',
        headers: { 
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  
        },
        body: getBlankExcelBuffer() // Directly use the ArrayBuffer for the blank Excel file    
    });

    if (!createRes.ok) throw new Error("Failed to create file");
    
    // 2. Loop through config to add Worksheets and Tables
    // Note: Excel creates "Sheet1" by default, so we use that for the first config item
    for (let i = 0; i < maeSystemConfig.worksheets.length; i++) {
        const sheet = maeSystemConfig.worksheets[i];
        await initializeSheetAndTable(accessToken, fileName, sheet, i === 0);
    }
    
    alert("Workshop System Initialized Successfully!");
}

//=========END FUNCTION createInitialWorkbook =============

//=========FUNCTION initializeSheetAndTable =============
/**
 * Step 3: Helper to add a sheet, add a table, and set headers.
 */
async function initializeSheetAndTable(accessToken, fileName, sheetConfig, isFirstSheet) {
    //const fileName = maeSystemConfig.spreadsheetName;
    const workbookUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook`;
    
    // A. Add (Skip if it's the first sheet "Sheet1")
    if (!isFirstSheet) {
        await fetch(`${workbookUrl}/worksheets`, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: sheetConfig.tabName })
        });
    } else {
        // Rename default Sheet1 to our first tabName
        await fetch(`${workbookUrl}/worksheets/Sheet1`, {
            method: 'PATCH',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: sheetConfig.tabName })
        });
    }

    // B. Create the Table
    // We assume a standard starting range (A1 to [Column Count]1)
    const lastColLetter = String.fromCharCode(64 + sheetConfig.columns.length);
    const tableRange = `${sheetConfig.tabName}!A1:${lastColLetter}1`;

    const tableRes = await fetch(`${workbookUrl}/worksheets/${sheetConfig.tabName}/tables/add`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ address: tableRange, hasHeaders: true })
    });
    
    const tableData = await tableRes.json();
    const tableId = tableData.id;

    // C. Rename Table to our tableName (The "Rugged" ID)
    await fetch(`${workbookUrl}/tables/${tableId}`, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: sheetConfig.tableName })
    });

    // D. Set Header Names
    const headers = sheetConfig.columns.map(col => col.header);
    await fetch(`${workbookUrl}/tables/${sheetConfig.tableName}/headerRowRange`, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: [headers] })
    });
}

//=========END FUNCTION initializeSheetAndTable =============

//========== Placeholder for loadTableData function =============
async function loadTableData(tableName) {
    console.log(`Loading data for table: ${tableName}`);
    const container = document.getElementById("table-container");
    container.innerHTML = `<p>Loading data for ${tableName}...</p>`;
    
    // Logic to fetch rows from Graph API will go here later
}
//========== End of Placeholder for loadTableData function =============







