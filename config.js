/**
 * MAE Custom Digital Solutions - Master System Manifest
 * Philosophy: Practical, Functional, Simple, Rugged.
 * Version 1.2.1: added active:true to json file
 * Version 1.2.2: added two TEST worksheets in the beginning
 * Version 1.2.3: changed "locked: true" to "locked: false" to allow editing from app 
 *      [except for mae_id, formulas, Master Dashboard (should be read-only) ]
 * Version xxxxxxx
 */

const maeSystemConfig = {
    spreadsheetName: "MAE_Workshop_Inventory_MASTER_TEMPLATE.xlsx",
    version: "1.2.3",
    
    worksheets: [
        {
            tabName: "TEST Inventory",
            tableName: "TEST_Inventory",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "TEST String", type: "string", locked: false },
                { header: "TEST Integer", type: "number", format: "0", locked: false },
                { header: "TEST Currency", type: "number", format: "$#,##0.00", locked: false },
                {
                    header: "TEST Formula", 
                    type: "formula",
                    formula: "=[@[TEST Integer]]*[@[TEST Currency]]",
                    format: "$#,##0.00",
                    locked: true
                },
                {
                    header: "TEST Dropdown",
                    type: "dropdown",
                    options: ["Red", "White", "Blue"],
                    locked: false
                }
            ]
        },
        {
            tabName: "TEST Dashboard",
            tableName: "TEST_Dashboard",
            active: true,
            columns: [
                {
                    header: "TEST calc from other table",
                    type: "formula",
                    formula: "=SUM(TEST_Inventory[TEST Currency])",
                    format: "$#,##0.00",
                    locked: true
                },
                {
                    header: "TEST Number Calc from other table",
                    type: "formula",
                    formula: "=SUM(TEST_Inventory[TEST Integer])",
                    format: "0",
                    locked: true
                }
            ]
        },
        {
            tabName: "Master Dashboard",
            tableName: "Master_Dashboard",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { 
                    header: "Total Inventory Value", 
                    type: "formula", 
                    formula: "=SUM(Resell_Inventory[Total Investment])",
                    format: "$#,##0.00", 
                    locked: true 
                },
                { 
                    header: "Low Stock Alerts", 
                    type: "formula",
                    formula: "=COUNTIF(Shop_Consumables[Current Stock], \"<\"&Shop_Consumables[Reorder Point])",
                    locked: true 
                },
                { header: "Upcoming Maintenance", type: "string", locked: true },
                { 
                    header: "Monthly Overhead Total", 
                    type: "formula", 
                    formula: "=SUM(Shop_Overhead[Amount])",
                    format: "$#,##0.00", 
                    locked: true 
                },
                { header: "Supplier Performance", type: "string", locked: true }
            ]
        },
        {
            tabName: "Resell Inventory",
            tableName: "Resell_Inventory",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset ID", type: "string", locked: false },
                { header: "Item Name", type: "string", locked: false },
                { header: "Category", type: "string", locked: false },
                { header: "Acquisition Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Purchase Price", type: "number", format: "$#,##0.00", locked: false },
                { header: "Restoration Cost", type: "number", format: "$#,##0.00", locked: false },
                { 
                    header: "Total Investment", 
                    type: "formula", 
                    formula: "=[@[Purchase Price]]+[@[Restoration Cost]]",
                    format: "$#,##0.00",
                    locked: true 
                },
                { header: "Current Status", type: "string", locked: false },
                { header: "Target Sale Price", type: "number", format: "$#,##0.00", locked: false },
                { header: "Actual Sale Price", type: "number", format: "$#,##0.00", locked: false },
                { header: "Location", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Machinery",
            tableName: "Shop_Machinery",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset ID", type: "string", locked: false },
                { header: "Machine Name/Model", type: "string", locked: false },
                { header: "Manufacturer/Brand", type: "string", locked: false },
                { header: "Serial Number", type: "string", locked: false },
                { header: "Purchase Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Purchase Cost", type: "number", format: "$#,##0.00", locked: false },
                { header: "Location", type: "string", locked: false },
                { header: "Status", type: "string", locked: false },
                { header: "Manual Link", type: "string", locked: false}
            ]
        },
        {
            tabName: "Maintenance Log",
            tableName: "Maintenance_Log",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Log ID", type: "string", locked: false },
                { header: "Asset ID", type: "string", locked: true },
                { header: "Service Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Service Type", type: "string", locked: false },
                { header: "Performed By", type: "string", locked: false },
                { header: "Cost", type: "number", format: "$#,##0.00", locked: false },
                { header: "Next Service Date", type: "date", format: "mm/dd/yyyy", locked: false }
            ]
        },
        {
            tabName: "Shop Power Tools",
            tableName: "Shop_Power_Tools",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tool ID", type: "string", locked: false },
                { header: "Tool Name/Model", type: "string", locked: false },
                { header: "Category", type: "string", locked: false},
                { header: "Condition", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Hand Tools",
            tableName: "Shop_Hand_Tools",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tool ID", type: "string", locked: false },
                { header: "Tool Name/Model", type: "string", locked: false },
                { header: "Category", type: "string", locked: false },
                { header: "Quantity", type: "number", format: "0", locked: false }
            ]
        },
        {
            tabName: "Shop Consumables",
            tableName: "Shop_Consumables",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Item Name", type: "string", locked: false },
                { header: "SKU/Item ID", type: "string", locked: false },
                { header: "Unit of Measure", type: "string", locked: false },
                { header: "Current Stock", type: "number", format: "0", locked: false },
                { header: "Reorder Point", type: "number", format: "0", locked: false },
                { header: "Unit Cost", type: "number", format: "$#,##0.00", locked: false },
                { header: "Preferred Supplier", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Overhead",
            tableName: "Shop_Overhead",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Expense Category", type: "string", locked: false },
                { header: "Description", type: "string", locked: false },
                { header: "Payment Frequency", type: "string", locked: false },
                { header: "Due Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Amount", type: "number", format: "$#,##0.00", locked: false },
                { header: "Auto-Pay?", type: "string", locked: false }
            ]
        },
        {
            tabName: "Supplier Contacts",
            tableName: "Supplier_Contacts",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Supplier Name", type: "string", locked: false },
                { header: "Contact Person", type: "string", locked: false },
                { header: "Phone", type: "string", locked: false },
                { header: "Email", type: "string", locked: false },
                { header: "Website", type: "string", locked: false}
            ]
        }
    ]
};


window.maeSystemConfig = maeSystemConfig; 

