// TODO code for implementing the Master Dashboard set of features
export const Dashboard = {
    // Helper to map the raw array from Excel into a readable object
    parseSummary: (row, config) => {
        const data = {};
        config.columns.forEach((col,index) => {
            data[col.header] = row[index];
        });

        return data;
    }
};