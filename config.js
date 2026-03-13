/**
 * MAE Custom Digital Solutions - Master System Manifest
 * Philosophy: Practical, Functional, Simple, Rugged.
 * THE MASTER SPREADSHEET MUST HAVE FORMULAS IN THE SPREADSHEET TO MATCH THIS
 * Version 1.2.1: added active:true to json file
 * Version 1.2.2: added two TEST worksheets in the beginning
 * Version 1.2.3: changed "locked: true" to "locked: false" to allow editing from app 
 *      [except for mae_id, formulas, Master Dashboard (should be read-only) ]
 * Version 1.2.4: modified several worksheets, added columns/dropdowns/added formulas
 *      to spreadsheet itself to match this file
 * Version xxxxxxx- 
 */

export const maeSystemConfig = {
    spreadsheetName: "MAE_Workshop_Inventory_MASTER_TEMPLATE.xlsx",
    version: "1.2.4",
    
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
                { 
                    header: "Category",
                    type: "dropdown",
                    options:["Machinery","Furniture","Electronics","Crafts","Auto Related","other"],
                    locked: false 
                },
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
                { 
                    header: "Current Status",
                    type: "dropdown",
                    options: ["Not Started","In-Progress","Complete","For Sale","Sold"],
                    locked: false 
                },
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
                { 
                    header: "Status",
                    type: "dropdown",
                    options: ["Operational","Needs Repair","Repair In-Progress","Unusable/Junk"],
                    locked: false 
                },
                { header: "Manual Link/Other Info", type: "string", locked: false}
            ]
        },
        {
            tabName: "Maintenance Log",
            tableName: "Maintenance_Log",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Log ID", type: "string", locked: false },
                { header: "Asset ID", type: "string", locked: false },
                { header: "Asset Description", type: "string", locked: false},
                { header: "Service Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { 
                    header: "Service Type",
                    type: "dropdown",
                    options: ["Preventive","Repair"],
                    locked: false 
                },
                { 
                    header: "Performed By",
                    type: "dropdown",
                    options: ["Self in Shop","Contractor in Shop","Outside Facility"],
                    locked: false 
                    },
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
                { 
                    header: "Functional Category",
                    type: "dropdown",
                    options: ["Driling","Cutting","Grinding","Sanding","Fastening","Shaping/Routing","Other"],
                    locked: false
                },
                {
                    header: "Operational Category",
                    type: "dropdown",
                    options: ["Portable/Handheld","Stationary/Bench Top","Outdoor Power Equipment"],
                    locked: false
                },
                {
                    header: "Power Source",
                    type: "dropdown",
                    options: ["Corded(A/C)","Cordless/Battery","Pneumatic/Air","Fuel-Powered"],
                    locked: false
                },
                { 
                    header: "Condition",
                    type: "dropdown",
                    options: ["Operational","Needs Repair","Repair in Progress","Unusable/Junk"],
                    locked: false 
                }
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
                { 
                    header: "Category",
                    type: "dropdown",
                    options: ["Fastening/Turning","Measuring/Layout","Striking/Hammering","Gripping/Holding","Cutting/Shaping","Other"],
                    locked: false 
                },
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
                { 
                  header: "Category",
                  type: "dropdown",
                  options: [
                             "Fasteners/Hardware","Abrasives/Cutting","Fluids/Lubricants/Chemicals","Primer/Paint",
                             "Safety/PPE","Shop/Janitorial","Welding","Electrical","Other"
                            ],
                  locked: false 
                },
                { header: "Unit of Measure", type: "string", locked: false },
                { header: "Current Stock", type: "number", format: "0", locked: false },
                { header: "Reorder Point", type: "number", format: "0", locked: false },
                { header: "Unit Cost", type: "number", format: "$#,##0.00", locked: false },
                {
                  header: "Current Inventory Value",
                  type: "formula",
                  formula: "=[@[Current Stock]]*[@[Unit Cost]]",
                  format: "$#,##0.00",
                  locked: false
                },
                { header: "Preferred Supplier", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Overhead",
            tableName: "Shop_Overhead",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { 
                    header: "Expense Category",
                    type: "dropdown",
                    options: [
                               "Facility Costs/Rent/Mortgage","Insurance","Debt/Leases on Equipment/Business Loans",
                               "Subscriptions","Salaries","Utilities","Maintenance/Repair","Marketing/Advertising",
                              "Professional Fees","Travel/Vehicles","Depreciation"
                            ],
                    locked: false 
                },
                { header: "Description", type: "string", locked: false },
                { 
                    header: "Payment Frequency",
                    type: "dropdown",
                    options: ["Upon Receipt","Monthly","Weekly","Quarterly","Semi-Annually","Yearly","Other"],
                    locked: false 
                },
                { header: "Due Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Amount", type: "number", format: "$#,##0.00", locked: false },
                { 
                    header: "Auto-Pay?",
                    type: "dropdown",
                    options: ["Yes","No"],
                    locked: false 
                }
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
                { header: "Website", type: "string", locked: false},
                { header: "Notes/Other Info", type: "string", locked: false}
            ]
        }
    ]
};


window.maeSystemConfig = maeSystemConfig; 

