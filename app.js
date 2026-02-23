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
const fileName = "MAE_Master_Inventory_Template.xlsx";

async function startup() {
    console.log("Checking for msal...", window.msal);

    if (typeof msal === 'undefined') {
        console.error("MSAL library not found. Check if msal-browser.min.js is in your project folder.");
        return;
    }

    console.log("MSAL started locally:", msal);

    try {
        //Intialize the PublicClientApplication
        //  MSAL V2 uses 'msal.PublicClientApplication'
        myMSALObj = new msal.PublicClientApplication(msalConfig);

        // This is the "Net" that catches the login result after the page reloads
    const response = await myMSALObj.handleRedirectPromise();
    
    if (response) {
        console.log("Caught user login!", response.account);
        account = response.account;
        updateUIForLoggedInUser(account);
    } 

        const authButton = document.getElementById("auth-btn");
        authButton.disabled = false; 

        // Check for existing session
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0) {
            account = accounts[0];
            console.log("Found existing account:", account.username);
            updateUIForLoggedInUser(account);
        } else {
            console.log("No existing session found. Waiting for user to click Connect.");
        }
//Endpoint of this function is to make the authButton active allow sign-in when clicked
        authButton.addEventListener("click", signIn);
    } catch (error) {
        console.error("Error during MSAL startup:", error);
    }
}
//========END STARTUP LOGIC ===========


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
    try {
        // Silent token acquisition is standard for rugged apps to avoid constant popups
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${fileName}:/workbook/tables`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
        });
        
        const data = await response.json();

        if (data.error) {
            console.error("Graph API Error:", data.error.message);
            return;
        }

        const menu = document.getElementById('menu');
        menu.innerHTML = ""; 

        data.value.forEach(table => {
            const li = document.createElement('li');
            li.style.cursor = "pointer";
            li.style.padding = "10px";
            
            const displayName = table.name.replace(/_/g, ' ').replace('Table', '');
            li.innerText = displayName;

            li.onclick = () => {
                document.getElementById('current-view-title').innerText = displayName;
                fetchTableData(table.name);
            };
            menu.appendChild(li);
        });
    } catch (error) {
        console.error("Error loading menu:", error);
        // If silent fails, user might need to click sign-in again
    }
}

//======= FUNCTION TO FETCH TABLE DATA ==============
async function fetchTableData(tableName) {
    try {
        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });

        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${fileName}:/workbook/tables/${tableName}/rows`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
        });
        const data = await response.json();
        
        console.log(`Data for ${tableName}:`, data.value);
        
        // Next step: render this data into the table-container
        const container = document.getElementById('table-container');
        container.innerHTML = `<pre>${JSON.stringify(data.value, null, 2)}</pre>`;
        
    } catch (error) {
        console.error("Error fetching table data:", error);
    }
}

//===TRIGGER that starts the whole engine ==========
console.log("App.js execution reaching the end...triggering startup()");

startup();
