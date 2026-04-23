/**
 * dashboard.js - MAE Custom Digital Solutions
 * Purpose: Data parsing and independent fetch logic for the Master Dashboard.
 */
import { myMSALObj } from './auth.js';

export const Dashboard = {
    // 1. Helper to map the raw array from Excel into a readable object
    parseSummary: (row, config) => {
        const data = {};
        config.columns.forEach((col, index) => {
            data[col.header] = row[index];
        });
        return data;
    },

    

    //  SELF-SUSTAINING: Fetch all rows for charts or subdivided lists
    async getFullTableData(tableName) {
        // USE GLOBAL HELPER in app.js
        const token = await window.getGraphToken();
        const fileName = window.maeSystemConfig.spreadsheetName;
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) throw new Error(`MAE System: Failed to fetch table ${tableName}`);
        const data = await response.json();

        // Return all row objects (contains .values and .index)
        return data.value;
    }
};