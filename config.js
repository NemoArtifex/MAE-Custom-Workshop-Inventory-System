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
 * Version 1.2.5: added Location_ID to Inv Sheets (to support future scannable tags), matched
 *      dropdowns with excel, added "Location" Table/Worksheet for future scannable tag
 * Version 1.3: FINALIZED Master Dashboard content and updated
 * Version 1.3.1: Updated Master Dashboard by adding 4 columns for calculations to support chart.js
 * Version 1.3.2: Updated Master Dashboard "Overhead Snapshot" to "Total Amount Due Next 30 Days" with new formula
 * Version 1.3.3: Added new worksheet: OVerhead Summary to support Master Dashboard Overhead calculations and source for graph
 * Version 1.3.4: Modified header in Master Dashboard table "Equipment With Operational Issues" and changed formula
 * Version 1.3.5: Added header in Master Dashboard for Maintenance Items card
 * Version 1.4:  Added features: enableScanning: true to support label scanning
 * Version 1.5: changed Asset_ID and Log_ID to "hidden:true", added checkboxes, modified some names
 * Version: 1.5.1: Modified formula for Maint Items Due in 30 days to include past due and not complete
 * Version: 1.5.2: updated Table: Location: changed "Name" to "Description; Added Tag_ID columns to 
 *                 Inventory related worksheets, adjusted column placement; made Location_ID type: "dropdown"
 * Version 1.5.3: changed quantity/current stock in hand tools and consumables to "hybrid inventory" type to 
 *                support both dropdown and number input
 * Version xxxxxx
 */

export const maeSystemConfig = {
    spreadsheetName: "MAE_Workshop_Inventory_MASTER_TEMPLATE.xlsx",
    version: "1.5.3",

    features: {
        enableScanning: true
    },
    
    worksheets: [
        {
            tabName: "TEST Inventory",
            tableName: "TEST_Inventory",
            active: false,
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
            active: false,
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
            active: false,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
        // Snapshot A: Resell Inventory Summary
                { 
                    header: "Total Resell Investment", 
                    type: "formula", 
                    formula: "=SUM(Resell_Inventory[Total Investment])",
                    format: "$#,##0.00", 
                    locked: true 
                },
                { 
                    header: "Total Actual Sales", 
                    type: "formula", 
                    formula: "=SUM(Resell_Inventory[Actual Sale Price])",
                    format: "$#,##0.00", 
                    locked: true 
                },
        // Snapshot B: Total Asset Value (Cross-Table Sum)
                {
                    header: "Total Machinery Value",
                    type: "formula",
                    formula: "=SUM(Shop_Machinery[Purchase Cost])",
                    format: "$#,#00.00",
                    locked: true,
                    hidden: true
                },
                {
                    header: "Total Power Tool Value",
                    type: "formula",
                    formula: "=SUM(Shop_Power_Tools[Purchase Price])",
                    format: "$#,#00.00",
                    locked: true,
                    hidden: true
                },
                {
                    header: "Total Hand Tool Value",
                    type: "formula",
                    formula: "=SUM(Shop_Hand_Tools[Purchase Price])",
                    format: "$#,#00.00",
                    locked: true,
                    hidden: true
                },
                {
                    header: "Total Consumables Value",
                    type: "formula",
                    formula: "=SUM(Shop_Consumables[Current Inventory Value])",
                    format: "$#,#00.00",
                    locked: true,
                    hidden: true
                },
                { 
                    header: "Total Shop Asset Value", 
                    type: "formula", 
                    formula: "=[@[Total Machinery Value]]+[@[Total Power Tool Value]]+[@[Total Hand Tool Value]]+[@[Total Consumables Value]]",
                    format: "$#,##0.00", 
                    locked: true 
                },
        // Snapshot C: Low Stock Alerts
                { 
                    header: "Low Stock Items Count", 
                    type: "formula",
                    formula: "=SUMPRODUCT(--(Shop_Consumables[Current Stock]<=Shop_Consumables[Reorder Point]))",
                    format: "0",       
                    locked: true 
                },
        // Snapshot E: Overhead Snapshot
                { 
                    header: "Total Amount Due Next 30 Days", 
                    type: "formula", 
                    formula: "=SUMIFS(Shop_Overhead[Amount], Shop_Overhead[Due Date], "<="&TODAY()+30, Shop_Overhead[Due Date], ">="&TODAY())",
                    format: "$#,##0.00", 
                    locked: true 
                },
        // Snapshot F: Condition Alerts (Count of items needing repair)
                { 
                    header: "Equipment With Operational Issues", 
                    type: "formula", 
                    // Backticks allow multi-line strings for better readability
                    formula: `
                        =SUM(
                            COUNTIFS(Shop_Machinery[Condition], {"Needs Repair","Repair In-Progress","Unusable/Junk"}),
                            COUNTIFS(Shop_Power_Tools[Condition], {"Needs Repair","Repair In-Progress","Unusable/Junk"}),
                            COUNTIFS(Shop_Hand_Tools[Condition], {"Needs Repair","Repair In-Progress","Unusable/Junk"})
                        )
                    `.trim(), // .trim() removes the extra line breaks at the start/end
                    format: "0",
                    locked: true 
                },
                {
                    header: "Maintenance Items Due in Next 30 Days",
                    type: "formula",
                    formula: `=COUNTIFS(Maintenance_Log[Next Service Date], "<="&TODAY()+30, Maintenance_Log[Complete], FALSE)`.trim(),
                    format: "0",
                    locked: true
                }
            ]
        },
        {
            tabName: "Overhead Summary",
            tableName: "Overhead_Summary",
            active: false,
            columns: [
                {
                    header: "Expense Category",
                    type: "string",
                    locked: true
                },
                {
                    header: "Annual Total",
                    type: "formula",
                    formula: "=SUMIF(Shop_Overhead[Expense Category], [@[Expense Category]], Shop_Overhead[Amount])",
                    format: "$#,##0.00",
                    locked: true
                }
            ]
        },
        {
            tabName: "Resell Inventory",
            tableName: "Resell_Inventory",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset ID", type: "string", hidden:true, locked: false },
                { header: "Tag_ID", type: "string", locked: false },
                { header: "Location_ID", type: "dropdown", locked: false},
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
                { header: "Location", type: "string", hidden:true,locked: false },
                { header: "Sold", type: "boolean", locked: false }
            ]
        },
        {
            tabName: "Shop Machinery",
            tableName: "Shop_Machinery",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset ID", type: "string", hidden: true, locked: false },
                { header: "Tag_ID", type: "string", locked: false },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Machine Name/Brand/Model", type: "string", locked: false },
                { header: "Serial Number", type: "string", locked: false },
                { header: "Purchase Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Purchase Cost", type: "number", format: "$#,##0.00", locked: false },
                { header: "Location", type: "string", hidden: true, locked: false },
                { 
                    header: "Condition",
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
                { header: "Log ID", type: "string", hidden: true, locked: false },
                { header: "Asset ID", type: "string", hidden: true, locked: false },
                { header: "Asset and Service Description", type: "string", locked: false},
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
                { header: "Next Service Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Complete", type: "boolean", locked: false },
                { header: "Remarks", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Power Tools",
            tableName: "Shop_Power_Tools",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset_ID", type: "string", hidden: true, locked: false },
                { header: "Tag_ID", type: "string", locked: false },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Tool Name/Brand/Model", type: "string", locked: false },
                { header: "Purchase Price", type: "number", format: "$#,##0.00", locked: false},
                { 
                    header: "Functional Category",
                    type: "dropdown",
                    options: ["Drilling","Cutting","Grinding","Sanding","Fastening","Shaping/Routing","Other"],
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
                    options: ["Operational","Needs Repair","Repair In-Progress","Unusable/Junk"],
                    locked: false 
                },
                { header: "Remarks", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Hand Tools",
            tableName: "Shop_Hand_Tools",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset_ID", type: "string", hidden: true, locked: false },
                { header: "Tag_ID", type: "string", locked: false },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Tool Name/Brand/Model/Description", type: "string", locked: false },
                { header: "Purchase Price", type: "number", format: "$#,##0.00", locked: false},
                { 
                    header: "Category",
                    type: "dropdown",
                    options: ["Fastening/Turning","Measuring/Layout","Striking/Hammering","Gripping/Holding","Cutting/Shaping","Other"],
                    locked: false 
                },
                {
                    header: "Condition",
                    type: "dropdown",
                    options: ["Operational","Needs Repair","Repair In-Progress","Unusable/Junk"],
                    locked: false
                },
                { header: "Quantity", type: "hybrid-inventory", options: ["Few", "Adequate", "Many", "Number"], locked: false },
                { header: "Remarks", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Consumables",
            tableName: "Shop_Consumables",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset_ID", type: "string", hidden: true, locked: false },
                { header: "Tag_ID", type: "string", locked: false },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Item Name", type: "string", locked: false },
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
                { header: "Current Stock", type: "hybrid-inventory", options: ["Few", "Adequate", "Many", "Number"], locked: false },
                { header: "Reorder Point", type: "number", format: "0", locked: false },
                { header: "Unit Cost", type: "number", format: "$#,##0.00", locked: false },
                {
                  header: "Current Inventory Value",
                  type: "formula",
                  formula: `
                            =IF(ISNUMBER([@[Current Stock]]), 
                            [@[Current Stock]] * [@[Unit Cost]], 
                            IFS([@[Current Stock]]="Many", 1.0, [@[Current Stock]]="Adequate", 0.5, [@[Current Stock]]="Few", 0.1, TRUE, 0) * [@[Unit Cost]]
                            )
                            `.trim(),
                  format: "$#,##0.00",
                  locked: true
                },
                { header: "Supplier/Remarks", type: "string", locked: false }
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
                              "Professional Fees","Travel/Vehicles","Depreciation","Other/Miscellaneous"
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
                },
                { header: "Remarks", type: "string", locked: false }
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
        },
        {
            tabName: "Location",
            tableName: "Location",
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false },
                { header: "Description", type: "string", locked: false },
                { header: "Type", type: "string", locked: false },
                { header: "Parent_Location", type: "string", locked: false}
            ]
        }
    ]
};


window.maeSystemConfig = maeSystemConfig; 

