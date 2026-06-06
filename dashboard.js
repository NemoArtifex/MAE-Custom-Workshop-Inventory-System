/**
 * dashboard.js - MAE Custom Digital Solutions
 * Purpose: Data parsing and independent fetch logic for the Master Dashboard.
 * Philosophy: Practical, Functional, Simple, Rugged.
 */
import { myMSALObj } from './auth.js';

export const Dashboard = {
    // 1. Helper to map the raw array from Excel into a readable object (Used for Single-Row Dashboard Metric Block)
    parseSummary: (row, config) => {
        const data = {};
        config.columns.forEach((col, index) => {
            data[col.header] = row[index];
        });
        return data;
    },

    // 🌟 MAE ENGINE RUGGED FIXED APPARATUS: SYSTEMIC DATABASE OBJECT NORMALIZER 🌟
    // Purpose: Unboxes double-nested Graph cell arrays and maps values permanently 
    // to their column blueprint header names, mirroring a relational database record row.
    transformRowsToObjects: function(graphRowsList, sheetConfig) {
        if (!graphRowsList || graphRowsList.length === 0) return [];

        return graphRowsList.map(rowObj => {
            // A. Safely unwrap Graph API's 2D double-nested array cell matrix container safely [[...]]
            const rawCells = (rowObj.values && Array.isArray(rowObj.values[0])) ? rowObj.values[0] : (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values : rowObj;
            const mappedDataMap = {};

            // B. Loop through your sheet blueprint configuration and bind cell values by string header name
            if (sheetConfig && sheetConfig.columns) {
                sheetConfig.columns.forEach((colDef, positionIndex) => {
                    if (rawCells && rawCells[positionIndex] !== undefined) {
                        mappedDataMap[colDef.header] = rawCells[positionIndex];
                    }
                });
            }

            // C. Yield a database-structured layout entity holding named metrics alongside row index tracking pointers
            return {
                index: parseInt(rowObj.index, 10), // Strict base-10 integer pointer for Microsoft Graph PATCH updates
                id: rowObj.id,                    // Original Graph cloud tracker row identification key
                data: mappedDataMap               // The clean, name-parsed data dictionary (The Database Record)
            };
        });
    },

    // 2. SELF-SUSTAINING LEDGER EXTRACTION FETCH CORE
    async getFullTableData(tableName) {
        const token = await window.getGraphToken();
        const fileName = window.maeSystemConfig.spreadsheetName;
        
        // Appending a dynamic timestamp parameter forces Microsoft's Graph servers
        // to bypass cloud caches and pull the exact ground-truth values from the Excel ledger.
   
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows?ts=${Date.now()}`;
        const response = await fetch(url, {
            headers: { 
                'Authorization': `Bearer ${token}`, 
                'Cache-Control': 'no-cache, no-store, must-revalidate', 
                'Pragma': 'no-cache' 
            }
        });
        
        if (!response.ok) throw new Error(`MAE System: Failed to fetch table ${tableName}`);
        const data = await response.json();
        
        // Return all raw row objects containing original Graph response variables (.values and .index)
        return data.value;
    }
};

window.Dashboard = Dashboard;

//===============
 //const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows?ts=${Date.now()}`;