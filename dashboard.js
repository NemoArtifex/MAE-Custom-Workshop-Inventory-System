// TODO code for implementing the Master Dashboard set of features
export const Dashboard = {
    // Helper to map the raw array from Excel into a readable object
    parseSummary: (row, config) => {
        const data = {};
        config.columns.forEach((col,index) => {
            data[col.header] = row[index];
        });

        return data;
    },
    async getPreviewItems(accessToken, tableName, filterType) {
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/workbook/tables/${tableName}/rows`;
        const response = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        const data = await response.json();
        
        // Use the same filter logic from app.js to get relevant rows
        const filtered = data.value.filter(/* your filter logic */);
        return filtered.slice(0, 3); // Return only top 3
    }
};