/**
 * MAE Custom Digital Solutions - Master System Manifest
 * Philosophy: Practical, Functional, Simple, Rugged.
 */

export const maeSystemConfig = {
    spreadsheetName: "MAE_Workshop_Inventory_MASTER_TEMPLATE.xlsx",
    version: "1.2.0",
    
    worksheets: [
        {
            tabName: "Master Dashboard",
            tableName: "Master_Dashboard",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Total Inventory Value", type: "number", format: "$#,##0.00", locked: true },
                { header: "Low Stock Alerts", type: "string", locked: true },
                { header: "Upcoming Maintenance", type: "string", locked: true },
                { header: "Monthly Overhead Total", type: "number", format: "$#,##0.00", locked: true },
                { header: "Supplier Performance", type: "string", locked: true }
            ]
        },
        {
            tabName: "Resell Inventory",
            tableName: "Resell_Inventory",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset ID", type: "string", locked: true },
                { header: "Item Name", type: "string", locked: true },
                { header: "Category", type: "string", locked: true },
                { header: "Acquisition Date", type: "date", format: "mm/dd/yyyy", locked: true },
                { header: "Purchase Price", type: "number", format: "$#,##0.00", locked: true },
                { header: "Restoration Cost", type: "number", format: "$#,##0.00", locked: true },
                { 
                    header: "Total Investment", 
                    type: "formula", 
                    formula: "=[[#This Row],[Purchase Price]]+[[#This Row],[Restoration Cost]]",
                    format: "$#,##0.00",
                    locked: true 
                },
                { header: "Current Status", type: "string", locked: true },
                { header: "Target Sale Price", type: "number", format: "$#,##0.00", locked: true },
                { header: "Actual Sale Price", type: "number", format: "$#,##0.00", locked: true },
                { header: "Location", type: "string", locked: true }
            ]
        },
        {
            tabName: "Shop Machinery",
            tableName: "Shop_Machinery",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset ID", type: "string", locked: true },
                { header: "Machine Name/Model", type: "string", locked: true },
                { header: "Manufacturer/Brand", type: "string", locked: true },
                { header: "Serial Number", type: "string", locked: true },
                { header: "Purchase Date", type: "date", format: "mm/dd/yyyy", locked: true },
                { header: "Purchase Cost", type: "number", format: "$#,##0.00", locked: true },
                { header: "Location", type: "string", locked: true },
                { header: "Status", type: "string", locked: true },
                { header: "Manual Link", type: "string", locked: true }
            ]
        },
        {
            tabName: "Maintenance Log",
            tableName: "Maintenance_Log",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Log ID", type: "string", locked: true },
                { header: "Asset ID", type: "string", locked: true },
                { header: "Service Date", type: "date", format: "mm/dd/yyyy", locked: true },
                { header: "Service Type", type: "string", locked: true },
                { header: "Performed By", type: "string", locked: true },
                { header: "Cost", type: "number", format: "$#,##0.00", locked: true },
                { header: "Next Service Date", type: "date", format: "mm/dd/yyyy", locked: true }
            ]
        },
        {
            tabName: "Shop Power Tools",
            tableName: "Shop_Power_Tools",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tool ID", type: "string", locked: true },
                { header: "Tool Name/Model", type: "string", locked: true },
                { header: "Category", type: "string", locked: true },
                { header: "Condition", type: "string", locked: true }
            ]
        },
        {
            tabName: "Shop Hand Tools",
            tableName: "Shop_Hand_Tools",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tool ID", type: "string", locked: true },
                { header: "Tool Name/Model", type: "string", locked: true },
                { header: "Category", type: "string", locked: true },
                { header: "Quantity", type: "number", format: "0", locked: true }
            ]
        },
        {
            tabName: "Shop Consumables",
            tableName: "Shop_Consumables",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Item Name", type: "string", locked: true },
                { header: "SKU/Item ID", type: "string", locked: true },
                { header: "Unit of Measure", type: "string", locked: true },
                { header: "Current Stock", type: "number", format: "0", locked: true },
                { header: "Reorder Point", type: "number", format: "0", locked: true },
                { header: "Unit Cost", type: "number", format: "$#,##0.00", locked: true },
                { header: "Preferred Supplier", type: "string", locked: true }
            ]
        },
        {
            tabName: "Shop Overhead",
            tableName: "Shop_Overhead",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Expense Category", type: "string", locked: true },
                { header: "Description", type: "string", locked: true },
                { header: "Payment Frequency", type: "string", locked: true },
                { header: "Due Date", type: "date", format: "mm/dd/yyyy", locked: true },
                { header: "Amount", type: "number", format: "$#,##0.00", locked: true },
                { header: "Auto-Pay?", type: "string", locked: true }
            ]
        },
        {
            tabName: "Supplier Contacts",
            tableName: "Supplier_Contacts",
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Supplier Contact Name", type: "string", locked: true },
                { header: "Category", type: "string", locked: true },
                { header: "Account Number", type: "string", locked: true },
                { header: "Primary Contact", type: "string", locked: true },
                { header: "Email", type: "string", locked: true },
                { header: "Phone", type: "string", locked: true },
                { header: "Lead Time", type: "string", locked: true },
                { header: "Website Link", type: "string", locked: true }
            ]
        }
    ]
};
