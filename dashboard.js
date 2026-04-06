/**
 * dashboard.js - MAE Custom Digital Solutions
 * Purpose: Data parsing and independent fetch logic for the Master Dashboard.
 */
import { myMSALObj } from './app.js';

export const Dashboard = {
    // 1. Helper to map the raw array from Excel into a readable object
    parseSummary: (row, config) => {
        const data = {};
        config.columns.forEach((col, index) => {
            data[col.header] = row[index];
        });
        return data;
    },

    // 2. SELF-SUSTAINING: Fetches top 3 items for a card preview
    async getPreviewItems(tableName, filterType) {
        try {
            // AUTH: Request token internally
            const accounts = myMSALObj.getAllAccounts();
            if (accounts.length === 0) throw new Error("No active account.");

            const tokenResponse = await myMSALObj.acquireTokenSilent({
                scopes: ["Files.ReadWrite"],
                account: accounts[0]
            });

            const fileName = window.maeSystemConfig.spreadsheetName;
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;

            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
            });

            if (!response.ok) throw new Error(`Fetch failed for ${tableName}`);
            const data = await response.json();

            // RUGGED: Return first 3 rows. Filtering logic can be applied here or in UI.
            return data.value.slice(0, 3);

        } catch (error) {
            console.error("MAE System: Preview fetch error:", error);
            return []; // Return empty array so UI doesn't break
        }
    },

    // 3. SELF-SUSTAINING: Fetch all rows for charts or subdivided lists
    async getFullTableData(tableName) {
        // AUTH: Request token internally
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length === 0) throw new Error("No active account.");

        const tokenResponse = await myMSALObj.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: accounts[0]
        });

        const fileName = window.maeSystemConfig.spreadsheetName;
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
        });

        if (!response.ok) throw new Error(`MAE System: Failed to fetch table ${tableName}`);
        const data = await response.json();

        // Return all row objects (contains .values and .index)
        return data.value;
    }
};