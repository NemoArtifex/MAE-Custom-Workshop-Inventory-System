/**
 * MAE Custom Digital Solutions - Master System Manifest
 * Philosophy: Practical, Functional, Simple, Rugged.
 * 
 * Target: Small Workshop Inventory Management
 * Storage: Customer-owned OneDrive via MSAL/Graph API
 */

export const maeSystemConfig = {
    spreadsheetName: "MAE_Workshop_Inventory.xlsx",
    version: "1.1.0",
    
    worksheets: [
        {
            tabName: "Master Dashboard",
            tableName: "Master_Dashboard",
            columns: [
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
                { header: "Asset ID", type: "string" },
                { header: "Item Name", type: "string" },
                { header: "Category", type: "string" },
                { header: "Acquisition Date", type: "date", format: "mm/dd/yyyy" },
                { header: "Purchase Price", type: "number", format: "$#,##0.00" },
                { header: "Restoration Cost", type: "number", format: "$#,##0.00" },
                { 
                    header: "Total Investment", 
                    type: "formula", 
                    formula: "=[[#This Row],[Purchase Price]]+[[#This Row],[Restoration Cost]]",
                    format: "$#,##0.00",
                    locked: true 
                },
                { header: "Current Status", type: "string" },
                { header: "Target Sale Price", type: "number", format: "$#,##0.00" },
                { header: "Actual Sale Price", type: "number", format: "$#,##0.00" },
                { header: "Location", type: "string" }
            ]
        },
        {
            tabName: "Shop Machinery",
            tableName: "Shop_Machinery",
            columns: [
                { header: "Asset ID", type: "string" },
                { header: "Machine Name/Model", type: "string" },
                { header: "Manufacturer/Brand", type: "string" },
                { header: "Serial Number", type: "string" },
                { header: "Purchase Date", type: "date", format: "mm/dd/yyyy" },
                { header: "Purchase Cost", type: "number", format: "$#,##0.00" },
                { header: "Location", type: "string" },
                { header: "Status", type: "string" },
                { header: "Manual Link", type: "string" }
            ]
        },
        {
            tabName: "Maintenance Log",
            tableName: "Maintenance_Log",
            columns: [
                { header: "Log ID", type: "string" },
                { header: "Asset ID", type: "string" },
                { header: "Service Date", type: "date", format: "mm/dd/yyyy" },
                { header: "Service Type", type: "string" },
                { header: "Performed By", type: "string" },
                { header: "Cost", type: "number", format: "$#,##0.00" },
                { header: "Next Service Date", type: "date", format: "mm/dd/yyyy" }
            ]
        },
        {
            tabName: "Shop Power Tools",
            tableName: "Shop_Power_Tools",
            columns: [
                { header: "Tool ID", type: "string" },
                { header: "Tool Name/Model", type: "string" },
                { header: "Category", type: "string" },
                { header: "Condition", type: "string" }
            ]
        },
        {
            tabName: "Shop Hand Tools",
            tableName: "Shop_Hand_Tools",
            columns: [
                { header: "Tool ID", type: "string" },
                { header: "Tool Name/Model", type: "string" },
                { header: "Category", type: "string" },
                { header: "Quantity", type: "number", format: "0" }
            ]
        },
        {
            tabName: "Shop Consumables",
            tableName: "Shop_Consumables",
            columns: [
                { header: "Item Name", type: "string" },
                { header: "SKU/Item ID", type: "string" },
                { header: "Unit of Measure", type: "string" },
                { header: "Current Stock", type: "number", format: "0" },
                { header: "Reorder Point", type: "number", format: "0" },
                { header: "Unit Cost", type: "number", format: "$#,##0.00" },
                { header: "Preferred Supplier", type: "string" }
            ]
        },
        {
            tabName: "Shop Overhead",
            tableName: "Shop_Overhead",
            columns: [
                { header: "Expense Category", type: "string" },
                { header: "Description", type: "string" },
                { header: "Payment Frequency", type: "string" },
                { header: "Due Date", type: "string" },
                { header: "Amount", type: "number", format: "$#,##0.00" },
                { header: "Auto-Pay?", type: "string" }
            ]
        },
        {
            tabName: "Supplier Contacts",
            tableName: "Supplier_Contacts",
            columns: [
                { header: "Supplier Contact Name", type: "string" },
                { header: "Category", type: "string" },
                { header: "Account Number", type: "string" },
                { header: "Primary Contact", type: "string" },
                { header: "Email", type: "string" },
                { header: "Phone", type: "string" },
                { header: "Lead Time", type: "string" },
                { header: "Website Link", type: "string" }
            ]
        }
    ]
};

