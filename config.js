/**
 * MAE Custom Digital Solutions - Master System Manifest
 * Philosophy: Practical, Functional, Simple, Rugged.
 * THE MASTER SPREADSHEET MUST HAVE FORMULAS IN THE SPREADSHEET TO MATCH THIS
 *
 * Version 3.0:  Finalized design.  Old Version Saved in another Word Document. 
 *                New Baseline: Master Dashboard/Overhead Summary (Hidden files) first; 
 *                then Location; then Inventory Worksheets; then Maintenance Log 
 *                then Non-Inventory worksheets;
 *                For future Customization (per customer)  will avoid changes impacting
 *                Master Dashboard and Inventory worksheets as much as possible
 *                to maintain integrity of system and ease of updates.
 *                Inventory Sheets: Tag_ID and Tag_Type available have hidden: true(false)
 *                BASE OPTION (NO Scannable label option): hidden: true;
 *                ADVANCED OPTION (Scannable label option): hidden: false; 
 * Version: xxxx:
 */

export const maeSystemConfig = {
    spreadsheetName: "MAE_Workshop_Inventory_MASTER_TEMPLATE.xlsx",
    version: "1.6.1",

    features: {
        enableScanning: true
    },
    
    worksheets: [
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
                    formula: "=SUMIFS(Resell_Inventory[Actual Sale Price], Resell_Inventory[Current Status], Sold)",
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
                    formula: `
                        =SUMPRODUCT(
                            ((Shop_Consumables[Stock_Level]="Counted") * (ISNUMBER(Shop_Consumables[Stock_Count])) * (Shop_Consumables[Stock_Count]<=Shop_Consumables[Reorder Point])) +
                            ((ISNUMBER(MATCH(Shop_Consumables[Stock_Level], {"Few","None"}, 0))))
                        )
                    `.trim(),
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
                    formula: `=COUNTIFS(Maintenance_Log[Scheduled Service Date], "<="&TODAY()+30, Maintenance_Log[Complete], FALSE)`.trim(),
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
            tabName: "Location",
            tableName: "Location",
            isInventory: false,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false },
                { header: "Description", type: "string", locked: false },     
            ]
        },
        {
            tabName: "Resell Inventory",
            tableName: "Resell_Inventory",
            isInventory: true,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tag_ID", type: "string", hidden: false,  locked: false },
                { header: "Tag_Type", type: "dropdown", options: ["UNIQUE", "MULTIPLE"], hidden: false, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Item_Description", type: "string", locked: false },
                { 
                    header: "Category",
                    type: "dropdown",
                    options:["Machinery","Furniture","Electronics","Crafts","Auto Related","other"],
                    hidden: true,
                    locked: false 
                },
                { header: "Date Acquired", type: "date", format: "mm/dd/yyyy", locked: false },
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
                { header: "Date Sold", type: "date", format: "mm/dd/yyyy", locked: false }
            ]
        },
        {
            tabName: "Shop Machinery",
            tableName: "Shop_Machinery",
            isInventory: true,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tag_ID", type: "string", hidden: false, locked: false },
                { header: "Tag_Type", type: "dropdown", options: ["UNIQUE", "MULTIPLE"], hidden: false, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Item_Description", type: "string", locked: false },
                { header: "Purchase Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Purchase Cost", type: "number", format: "$#,##0.00", locked: false },
                { 
                    header: "Condition",
                    type: "dropdown",
                    options: ["Operational","Needs Repair","Repair In-Progress","Unusable/Junk"],
                    locked: false 
                },
                { header: "Serial Number/Other Info", type: "string", locked: false}
            ]
        },
        {
            tabName: "Shop Power Tools",
            tableName: "Shop_Power_Tools",
            isInventory: true,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tag_ID", type: "string", hidden: false, locked: false },
                { header: "Tag_Type", type: "dropdown", options: ["UNIQUE", "MULTIPLE"], hidden: false, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Item_Description", type: "string", locked: false },
                { header: "Purchase Price", type: "number", format: "$#,##0.00", locked: false},
                { 
                    header: "Functional Category",
                    type: "dropdown",
                    options: ["Drilling","Cutting","Grinding","Sanding","Fastening","Shaping/Routing","Other"],
                    hidden: true,
                    locked: false
                },
                {
                    header: "Operational Category",
                    type: "dropdown",
                    options: ["Portable/Handheld","Stationary/Bench Top","Outdoor Power Equipment"],
                    hidden: true,
                    locked: false
                },
                {
                    header: "Power Source",
                    type: "dropdown",
                    options: ["Corded(A/C)","Cordless/Battery","Pneumatic/Air","Fuel-Powered"],
                    hidden: true,
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
            isInventory: true,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tag_ID", type: "string", hidden: false, locked: false },
                { header: "Tag_Type", type: "dropdown", options: ["UNIQUE", "MULTIPLE"], hidden: false, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Item_Description", type: "string", locked: false },
                { header: "Purchase Price", type: "number", format: "$#,##0.00", locked: false},
                { 
                    header: "Category",
                    type: "dropdown",
                    options: ["Fastening/Turning","Measuring/Layout","Striking/Hammering","Gripping/Holding","Cutting/Shaping","Other"],
                    hidden: true,
                    locked: false 
                },
                {
                    header: "Condition",
                    type: "dropdown",
                    options: ["Operational","Needs Repair","Repair In-Progress","Unusable/Junk"],
                    locked: false
                },
                { 
                    header: "Stock_Level", 
                    type: "dropdown", 
                    options: ["Few", "Adequate", "Many", "Counted"],
                    hidden: false, 
                    locked: false },
                {
                    header: "Stock_Count",
                    type: "number",
                    format: "0",
                    locked: false
                },
                { header: "Remarks", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Consumables",
            tableName: "Shop_Consumables",
            isInventory: true,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Tag_ID", type: "string", hidden: false, locked: false },
                { header: "Tag_Type", type: "dropdown", options: ["UNIQUE", "MULTIPLE"], hidden: false, locked: true },
                { header: "Location_ID", type: "dropdown", locked: false},
                { header: "Item_Description", type: "string", locked: false },
                { 
                  header: "Category",
                  type: "dropdown",
                  options: [
                             "Fasteners/Hardware","Abrasives/Cutting","Fluids/Lubricants/Chemicals","Primer/Paint",
                             "Safety/PPE","Shop/Janitorial","Welding","Electrical","Other"
                            ],
                  locked: false 
                },
                { 
                    header: "Stock_Level",
                    type: "dropdown", 
                    options: ["None","Few", "Adequate", "Many", "Counted"], 
                    locked: false 
                },
                { header: "Unit of Measure", type: "string", locked: false },
                { header: "Unit Cost", type: "number", format: "$#,##0.00", locked: false },
                {
                    header: "Stock_Count",
                    type: "number",
                    format: "0",
                    locked: false
                },
                {
                    header:"Bulk_Value",
                    type: "number",
                    format: "$#,##0.00",
                    locked: false
                },
                { header: "Reorder Point", type: "number", format: "0", locked: false },
                {
                  header: "Current Inventory Value",
                  type: "formula",
                  formula: `
                            =IF(
                                [@[Stock_Level]]="Counted",
                                [@[Stock_Count]] * [@[Unit Cost]],
                                IF(
                                    [@[Stock_Level]]="None",
                                    0,
                                    [@[Bulk_Value]] * IFS(
                                        [@[Stock_Level]]="Many", 1.0, 
                                        [@[Stock_Level]]="Adequate", 0.5, 
                                        [@[Stock_Level]]="Few", 0.25, 
                                        TRUE, 0
                                    )
                                )
                            )
                            `.trim(),
                  format: "$#,##0.00",
                  locked: true
                },
                { header: "Supplier/Remarks", type: "string", locked: false }
            ]
        },
        {
            tabName: "Maintenance Log",
            tableName: "Maintenance_Log",
            isInventory: false,
            active: true,
            columns: [
                { header: "mae_id", type: "string", hidden: true, locked: true },
                { header: "Asset and Service Description", type: "string", locked: false},
                { 
                    header: "Performed By",
                    type: "dropdown",
                    options: ["Self in Shop","Contractor in Shop","Outside Facility"],
                    locked: false 
                    },
                { header: "Cost", type: "number", format: "$#,##0.00", locked: false },
                { header: "Scheduled Service Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Complete", type: "boolean", locked: false },
                { header: "Completion Date", type: "date", format: "mm/dd/yyyy", locked: false },
                { header: "Remarks", type: "string", locked: false }
            ]
        },
        {
            tabName: "Shop Overhead",
            tableName: "Shop_Overhead",
            isInventory: false,
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
            isInventory: false,
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

