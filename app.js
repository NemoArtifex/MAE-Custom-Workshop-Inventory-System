// =============CONFIGURATION: The "Blueprint"  ======================
const msalConfig = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "http://localhost:5500" ,
        redirectUri: "https://nemoartifex.github.io/MAE-Custom-Workshop-Inventory-System/"
    },
    cache: {
        cacheLocation: "sessionStorage", // Simple and effective for workshop environments
        storeAuthStateInCookie: false,
    }
};
// ===========END CONFIGURATION =============

// =========== STARTUP LOGIC ============
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

        authButton.addEventListener("click", signIn);
    } catch (error) {
        console.error("Error during MSAL startup:", error);
    }
}



//===========SIGN-IN FUNCTION ==========
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

// ======== FUNCTION TO UPDATE UI BASED ON LOGIN STATUS ========
function updateUIForLoggedInUser(userAccount) {
    const authButton = document.getElementById("auth-btn");
    console.log("Enabling the Connect button now...");
    authButton.disabled = false;
    authButton.innerText = `Connected: ${userAccount.username}`;
    authButton.style.background = "#27ae60"; 
    authButton.style.color = "white";

    console.log("Loading dynamic menu for user:", userAccount.username);
    loadDynamicMenu();
}

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
