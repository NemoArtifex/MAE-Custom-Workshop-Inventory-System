const config = {
    auth: {
        clientId: "YOUR_CLIENT_ID",
        authority: "https://login.microsoftonline.com",
        redirectUri: "https://yourusername.github.io" // Critical for GitHub Pages
    }
};

const msalInstance = new msal.PublicClientApplication(config);
let activeSheet = "Dashboard";

async function switchSheet(sheetName) {
    activeSheet = sheetName;
    document.getElementById('current-view-title').innerText = sheetName;
    
    // Map internal names to your actual Excel Worksheet names
    const sheetMap = {
        'ResellInventory': 'Resell_Inventory_Table',
        'Tools': 'Shop_Tools_Table',
        'Suppliers': 'Supplier_Contact_List'
    };
    
    fetchTableData(sheetMap[sheetName]);
}

async function fetchTableData(tableName) {
    const url = `https://graph.microsoft.com{tableName}/rows`;
    // ... same fetch logic as previous example ...
}
