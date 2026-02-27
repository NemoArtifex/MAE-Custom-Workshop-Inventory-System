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
 */

async function createInitialWorkbook(accessToken) {
   
// This is a Base64 string for a 100% valid, blank Excel workbook.  
    const base64 = "UEsDBBQAAAAAAM6QKVYAAAAAAAAAAAAAAAAGAAAAX3JlbHMvUEsDBBQAAAAAAM6QKVYAAAAAAAAAAAAAAAALAAAAX3JlbHMvLnJlbHOEzwEOwiAMBNC7E99B9m4GNozYm3AByU0S6vXvYmIDV+v6m5S2TofpLpM7v9ArNidI2YpUq88LpL2H8u2G9XyD0U6qGZ9Iq9WRE9pBy0W0fLpB9XmE0uIsX3C+pInoB9eXWvL6B62T3y6vUEsDBBQAAAAAAM6QKVYAAAAAAAAAAAAAAAALAAAAeGwvcmVscy8ucmVsc1BLAwQUAAAAAADOkClWAAAAAAAAAAAAAAAAEAAAAHhsL3dvcmtib29rLnhtbFBLAwQUAAAAAADOkClWAAAAAAAAAAAAAAAAFAAAAHhsL3NoZWV0cy9zaGVldDEueG1sUEsBAhQAFAAAAAAAzpApVvLpW9MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfcmVscy9QSwECAhQAFAAAAAAAzpApVvLpW9MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfcmVscy8ucmVsc1BLAQICFAAUAAAAAADOkClW8ulb0wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHhsL3JlbHMvLnJlbHNQSwECAhQAFAAAAAAAzpApVvLpW9MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4bC93b3JrYm9vay54bWxQSwECAhQAFAAAAAAAzpApVvLpW9MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4bC9zaGVldHMvc2hlZXQxLnhtbFBLBQYAAAAABQAFAAsBAAB6AAAAAAA=";

    try {
        const binaryString = window.atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/content`;
    
        const createRes = await fetch(url, {
            method: 'PUT',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/octet-stream'
            }
        });

    // IMPORTANT: Wait for OneDrive to index the new file before adding tables
        await new Promise(resolve => setTimeout(resolve, 3000));
           
    // 2. Loop through config to add Worksheets and Tables
    // Note: Excel creates "Sheet1" by default, so we use that for the first config item
        for (let i = 0; i < maeSystemConfig.worksheets.length; i++) {
            const sheet = maeSystemConfig.worksheets[i];
            await initializeSheetAndTable(accessToken, fileName, sheet, i === 0);
       }

    } catch (error) {
        console.error("Error creating initial workbook:", error);
        alert("Failed to initialize system. check console for details.");
    }
}

//=========END FUNCTION createInitialWorkbook =============

//=========FUNCTION initializeSheetAndTable =============
/**
 * Step 3: Helper to add a sheet, add a table, and set headers.
 */
async function initializeSheetAndTable(accessToken, fileName, sheetConfig, isFirstSheet) {

    const workbookUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook`;
    const authHeader = {
        'Authorization': 'Bearer ${accessToken}',
        'Content-Type': 'application/json'
    }
    // A. Handle worksheet (Rename index 0)
    if (isFirstSheet){
        await fetch(`${workbookUrl}/worksheets/itemAt(index=0)`, {
            method: 'PATCH',
            headers: authHeader,
            body: JSON.stringify({ name: sheetConfig.tabName })
        });
    } else {
        await fetch(`${workbookUrl}/worksheets/add`, {
            method: 'POST',
            headers: authHeader,
            body: JSON.stringify({ name: sheetConfig.tabName })
        });
    }

    // B. Create the Table
    // We assume a standard starting range (A1 to [Column Count]1)
    const lastColLetter = String.fromCharCode(64 + sheetConfig.columns.length);
    const tableRange = `${sheetConfig.tabName}!A1:${lastColLetter}1`;

    console.log(`Creating table: ${sheetConfig.tableName} at ${tableRange}`);
    //const tableRes = await fetch(`${workbookUrl}/worksheets/${sheetConfig.tabName}/tables/add`, {
    const tableRes = await fetch(`${workbookUrl}/worksheets/${encodeURIComponent(sheetConfig.tabName)}/tables/add`, {
        method: 'POST',
        method: 'POST',
        headers: authHeader,
        body: JSON.stringify({
            address: tableRange,
            hasHeaders: true 
        })
   });

   if (!tableRes.ok){
    const errorDetails = await tableRes.json();
    console.error(`Table creation failed at ${sheetConfig.tabName}:`, errorDetails);
    return; //STOOP if this step fails to prevent cascade errors
   }
    
    const tableData = await tableRes.json();
    const tableId = tableData.id;

    // C. Rename Table to our tableName (The "Rugged" ID)
    await fetch(`${workbookUrl}/tables/${tableId}`, {
        method: 'PATCH',
        headers: authHeader,
        body: JSON.stringify({ name: sheetConfig.tableName })
    });

    // D. Set Header Names
    const headers = sheetConfig.columns.map(col => col.header);
    await fetch(`${workbookUrl}/tables/${tableId}/headerRowRange`, {
        method: 'PATCH',
        headers: authHeader,
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







