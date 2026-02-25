import { maeSystemConfig } from './config.js'
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
    },
    system: {
        // increases reliability for popup communication  
        allowRedirectInIframe: true,
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
        myMSALObj = new msal.PublicClientApplication(msalConfig);
        const response = await myMSALObj.handleRedirectPromise();
    
        if (response) {
        account = response.account;
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
        scopes: ["User.Read", "Files.ReadWrite"]
    };

    try {
        // v2 supports the simple loginPopup method
        const loginResponse = await myMSALObj.loginPopup(loginRequest);
        console.log("Login Successful:", loginResponse);
        account = loginResponse.account;
        updateUIForLoggedInUser(account);
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
    console.log("Starting sign-out process...");
     // Safeguard: if the account object is missing, just reset the UI locally
    if (!account) return resetUI();

    try {
        const logoutRequest = {
            account: myMSALObj.getAccountByUsername(account.username),
            // Where the popup should go after it finishes
            postLogoutRedirectUri: window.location.origin 
        };
     // triggers popup   
        await myMSALObj.logoutPopup(logoutRequest);
        
    } catch (error) {
        // If an interaction is already happening, MSAL is "locked." 
        // We catch that error and force a local cleanup anyway.
        if (error.errorMessage && error.errorMessage.includes("interaction_in_progress")) {
            console.warn("Interaction locked. Forcing local logout.");
        } else { 
        console.error("Sign-out failed:", error);
    }
 } finally {
    // This 'finally' block ensures your UI resets NO MATTER WHAT
        account = null;
        sessionStorage.clear(); // Rugged move: wipe the cache manually
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

async function verifySpreadsheetExists(){
    // Logic here to check if maeSystemConfig.spreadsheetName exists
    // If 404: Call a function to CREATE the workbook using the config
    // If 200: All good, ready to work.
    const tokenResponse = await myMSALObj.acquireTokenSilent({
        scopes: ["Files.ReadWrite"],
        account: account
    });

    const fileName = maeSystemConfig.spreadsheetName;
    // Check if file exists in the root of OneDrive
    const url = `https://graph.microsoft.com{fileName}`;

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

/**
 * Step 2: Create the .xlsx file and build the Tables/Headers
 * from the maeSystemConfig.
 */
async function createInitialWorkbook(accessToken) {
    const fileName = maeSystemConfig.spreadsheetName;
    const baseUrl = `https://graph.microsoft.com`;

    // 1. Create the empty Excel file
    const createRes = await fetch(baseUrl, {
        method: 'POST',
        headers: { 
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json' 
        },
        body: JSON.stringify({
            "name": fileName,
            "file": {},
            "@microsoft.graph.conflictBehavior": "fail"
        })
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

/**
 * Step 3: Helper to add a sheet, add a table, and set headers.
 */
async function initializeSheetAndTable(accessToken, fileName, sheetConfig, isFirstSheet) {
    const workbookUrl = `https://graph.microsoft.com{fileName}:/workbook`;
    
    // A. Add Worksheet (Skip if it's the first sheet "Sheet1")
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

//==========END FUNCTION Load Dynamic Menu ================





