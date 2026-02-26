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
        
        myMSALObj = new window.msal.PublicClientApplication(msalConfig);

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

    const fileName = maeSystemConfig.spreadsheetName;
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
// A minimal, valid Base64 string for an empty .xlsx file
const base64BlankExcel = "UEsDBBQAAAAAAL95WlYAAAAAAAAAAAAAAAAHAAAAX3JlbHMvLnJlbHOtkt1KAzEQRe+D7xDmbmZ7IaKybS9K6Uf0ASp9gEnbaZpMZpIofXvHrqAtpYIgepEhc+acc7Inm4vXatYpE3uXUFTVChid3XatSvi6vX/6AWIu6Iyd8yR8mYOnfXu7P9pInid0GZunmPsh9SPhSogY8uYpEq6mGAsG/yY04GvVzVzOnwZitFp7m8I504XQ01n12vX84/iYQ6fD8iGfDqO2KCOY8K0hNq9995I5xT/lM4p+FhVn2m2k9KAn8Yv6V1LzNn+Nivz4eD0E/G8+7uSvyf73fQ0UEsDBBQAAAAAAL95WlYAAAAAAAAAAAAAAAAHAAAAdXJsLnJlbHNz909SCS4tL0nNVXDOz9MvS80r0U/OSSwuVvBKTU7OyczPY/BNSfXxd/EP8Q3x9fX0DQpS8PP39A0KUtB3yc9TMNQvKC1KT83Lz0u15VBLAwQUAAAAAAC/eVpWAAAAAAAAAAAAAAAACAAAAHhsL3dvcmtib29rLnhtbG9Xy0rDQAzeB98hzN3MdhCisO0uSmn/oA9Q6QOsst00SZZZkvTtuHVV9CJ6kcGZfMmc7E5m07pU67SMvevFmG7WwOisNulS8XF9f/cLxFzQGTvmSfgmD572ze3u6KJ5ntBnbJ5i7ofUj4QrIWK4mKdImE4xFgz+SWjA5mXbeZk/DcRolfYuhXOnC6GnseqVzbnT38WHzp0Pyod8OozSko9gwoeG2Dx77yVzin/Kp7T9f1X9An0UfGg+Kof0uK0L6p9RkD8fr66A/5uPOfkrsv+HjwFUEsDBBQAAAAAAL95WlYAAAAAAAAAAAAAAAAJAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNz909SCS4tL0nNVXDOz9MvS80r0U/OSSwuVvBKTU7OyczPY/BNSfXxd/EP8Q3x9fX0DQpS8PP39A0KUtB3yc9TMNQvKC1KT83Lz0u15VBLAwQUAAAAAAC/eVpWAAAAAAAAAAAAAAAAEQAAAHhsL3NoZWV0cy9zaGVldDEueG1sbVfLToNAEN0bvwOZO8y29iNEVNo9Kam90AfY9AFG2U6TyU6S6I93LNoYjE/0IsM5XGaGm5k83Vz1U6tG56VvT6E2S6CRVr1Iu1C/v7w8fAKfS7YRR9/YUKvTCH+v7+9eP2XPI7YatitY74M0TISbICVvKqESpqscC9qfRErcrGzfXP66iNE6092VsVq7EGrV61G7hT39W7/p3vmobKgnw9A6pBUM7fGgO6e+e5Q9yT9lpWp/FpVmWs3U/tYp+Wf8m6/i6eP1I+N/9mFDfyf1t/81UEsBAhQAFAAAAAAAv3laVgAAAAAAAAAAAAAAAAcAAAAAAAAAAAAAAAAAAAAAAF9yZWxzLy5yZWxzUEsBAhQAFAAAAAAAv3laVgAAAAAAAAAAAAAAAAcAAAAAAAAAAAAAAAAALAAAAHVybC5yZWxzUEsBAhQAFAAAAAAAv3laVgAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAYQAAAHhsL3dvcmtib29rLnhtbFBLAQIUABQAAAAAAL95WlYAAAAAAAAAAAAAAAAJAAAAAAAAAAAAAAAAAL0AAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQIUABQAAAAAAL95WlYAAAAAAAAAAAAAAAARAAAAAAAAAAAAAAAAAPEAAAB4bC9zaGVldHMvc2hlZXQxLnhtbFBLBQYAAAAAAQABAFoAAAA1AQAAAAA=";

// 1. Convert Base64 to a Binary ArrayBuffer
function base64ToBuffer(base64) {
    const binaryString = atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}
async function createInitialWorkbook(accessToken) {
    const excelBinaryData = base64ToBuffer(base64BlankExcel);
    const fileName = maeSystemConfig.spreadsheetName;
    const baseUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/content?@microsoft.graph.conflictBehavior=fail`;
    
    // 1. Create the empty Excel file
    const createRes = await fetch(baseUrl, {
        method: 'PUT',
        headers: { 
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  
        },
        body: excelBinaryData
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

//=========END FUNCTION initializeSheetAndTable =============







