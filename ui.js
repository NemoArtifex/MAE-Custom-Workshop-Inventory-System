/**
 * ui.js - MAE Custom Digital Solutions
 * Purpose: Handle all DOM manipulation and visual states.
 * Philosophy: Practical, Functional, Simple, Rugged.
 */

// Helper to format numbers as currency $0.00; 
//NOTE OUTSIDE of main UI object to keep it available to all functions
// without cluttering the main UI object
const formatCurrency = (value) => {
    const num = parseFloat(value);
    if (isNaN(num)) return value; // Return as-is if it's not a number
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 2 // Ensures .00 always appears
    }).format(num);
};

import { Dashboard } from './dashboard.js';

export const UI = {
    
    // 1. AUTH UI STATES
    // Updates the connection button and clear/set messages based on login
    setConnected(username, signOutCallback) {
        const authButton = document.getElementById("auth-btn");
        authButton.innerText = `Sign Out: ${username}`;
        authButton.style.background = "#c0392b"; // Red for Sign Out
        authButton.style.color = "white";
        authButton.disabled = false;
        authButton.onclick = signOutCallback;
    },

    setDisconnected(signInCallback) {
        const authButton = document.getElementById("auth-btn");
        authButton.innerText = "Connect Microsoft Office 365";
        authButton.style.background = ""; // Reverts to CSS default (Green)
        authButton.style.color = "";
        authButton.onclick = signInCallback;
        
        // Clear all data zones
        document.getElementById("menu").innerHTML = "";
        document.getElementById("table-container").innerHTML = "";
        document.getElementById("action-bar-zone").innerHTML = ""; // Clear the buttons
        document.getElementById("current-view-title").innerText = "Please connect to view inventory data.";
    },

    // 2. DYNAMIC MENU RENDERING
    // Builds the sidebar buttons based on maeSystemConfig
    // ui.js - Updated renderMenu
renderMenu(activeWorksheets, onClickCallback) {
    const menu = document.getElementById("menu");
    menu.innerHTML = ""; // Clear existing

    // 1. ADD STATIC HOME/DASHBOARD BUTTON
    const homeLi = document.createElement("li");
    const homeBtn = document.createElement("button");
    homeBtn.innerText = "🏠 Workshop Dashboard";
    
    // RUGGED: We start with the dashboard as 'active' on login
    homeBtn.className = "menu-btn home-btn active"; 
    
    homeBtn.onclick = () => {
        // Handle visual state
        document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
        homeBtn.classList.add('active');
        
        // USE THE CALLBACK: This sends "Master_Dashboard" back to app.js
        // where loadTableData is defined.
        onClickCallback("Master_Dashboard"); 
    };

    homeLi.appendChild(homeBtn);
    menu.appendChild(homeLi);

    // 2. RENDER DYNAMIC INVENTORY LINKS
    activeWorksheets.forEach(sheet => {
        const li = document.createElement("li");
        const btn = document.createElement("button");
        btn.innerText = sheet.tabName;
        btn.className = "menu-btn";
        
        btn.onclick = () => {
            document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Standard inventory table loading
             onClickCallback(sheet.tableName);
            };

         li.appendChild(btn);
         menu.appendChild(li);
        });
    },

    // 3. TABLE RENDERING (The "Worker" logic refactored from app.js)
    // Practical: Uses the Config "Blueprint" to filter out hidden technical columns.
    // Rugged: Handles empty states and Microsoft Graph's row structure.
    renderTable(rows, tableName, sheetConfig, customTitle = null) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");

    if (!sheetConfig) {
        container.innerHTML = "Error: Worksheet configuration not found.";
        return;
    }

    title.innerHTML = customTitle || `View: ${sheetConfig.tabName}`;

    // Check if this is an "Operational Issues" view
    const isRepairsView = customTitle && customTitle.includes("Operational Issues");

    if (isRepairsView) {
        this.renderSubdividedRepairs(rows, tableName, sheetConfig);
        return; 
    }

    // NEW: Find the index of the mae_id column to protect it
    const idIndex = sheetConfig.columns.findIndex(c => c.header === "mae_id");

    // 1. Identify visible columns from the Manifest
    const visibleIndices = [];
    let html = `<table class="inventory-table" id="main-data-table"><thead><tr>`;
    
    html += `<th class="edit-only-cell">Action</th>`;
    
    sheetConfig.columns.forEach((col, index) => {
        if (col.hidden !== true) { 
            html += `<th>${col.header}</th>`;
            visibleIndices.push(index);
        }
    });
    html += `</tr></thead><tbody>`;

    // 2. Render Rows
    if (rows && rows.length > 0) {
        rows.forEach((row) => {
            const persistentIndex = row.index; 
            const allCells = Array.isArray(row.values[0]) ? row.values[0] : row.values; 

            // RUGGED: Extract the actual ID and anchor it to the row attribute
            // This prevents data loss if the cell text is edited or cleared
            const rawMaeId = (idIndex !== -1) ? allCells[idIndex] : '';

            html += `<tr data-row-index="${persistentIndex}" data-mae-id="${rawMaeId}">`;

            html += `<td class="edit-only-cell">
                        <button class="delete-row-btn" onclick="requestDelete(${persistentIndex})">🗑️</button>
                     </td>`;

            visibleIndices.forEach(idx => {
                const colDef = sheetConfig.columns[idx];
                const isEditable = !colDef.locked && colDef.type !== 'formula';
                
                const isCurrentStock = colDef.header === "Current Stock";
                const isReorderPoint = colDef.header === "Reorder Point";
                const isQuantity = colDef.header === "Quantity" || colDef.header === "Current Stock";
                const isCurrency = colDef.format && colDef.format.includes("$");
                
                let displayValue = allCells[idx] ?? '';

                if (isCurrency) {
                    displayValue = formatCurrency(displayValue);
                }

                if (colDef.type === 'boolean') {
                    const isChecked = displayValue.toString().toUpperCase() === "TRUE";
                    displayValue = `<input type="checkbox" disabled ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
                }

                html += `<td 
                        class="${isEditable ? 'editable-cell' : 'locked-cell'}
                               ${isCurrentStock ? 'col-type-stock-alert' : ''}
                               ${isReorderPoint ? 'col-type-reorder-point' : ''}
                               ${isQuantity ? 'col-type-qty' : ''} 
                               ${isCurrency ? 'col-type-currency' : ''}" 
                        data-col-index="${idx}">${displayValue}</td>`;
            });
            html += `</tr>`;
        });
    } else {
        const colSpan = visibleIndices.length + 1;
        html += `<tr><td colspan="${colSpan}" style="text-align:center; padding:20px;">No records found.</td></tr>`;
    }

    html += `</tbody></table>`;
    container.innerHTML = html;
},
    // ============ END RENDER TABLE ============

    //=========== RENDER SUBDIVIDED REPAIRS Specialized Renderer ==============
   renderSubdividedRepairs(rows, tableName, sheetConfig) {
    const container = document.getElementById("table-container");
    const conditions = ["Needs Repair", "Repair In-Progress", "Unusable/Junk"];

    // NEW: Find the index of the mae_id column to protect it
    const idIndex = sheetConfig.columns.findIndex(c => c.header === "mae_id");

    const condIdx = sheetConfig.columns.findIndex(c => c.header === "Condition");
    const visibleIndices = sheetConfig.columns
        .map((col, i) => col.hidden !== true ? i : -1)
        .filter(i => i !== -1);

    let html = `<table class="inventory-table" id="main-data-table"><thead><tr>`;
    html += `<th class="edit-only-cell">Action</th>`;
    visibleIndices.forEach(idx => html += `<th>${sheetConfig.columns[idx].header}</th>`);
    html += `</tr></thead>`;

    conditions.forEach(status => {
        const groupRows = rows.filter(r => {
            const rowCells = r.values[0]; 
            return rowCells && rowCells[condIdx] === status;
        });

        if (groupRows.length > 0) {
            html += `
                <tbody class="repair-group-header">
                    <tr>
                        <td colspan="${visibleIndices.length + 1}" 
                            style="background: #34495e; color: white; font-weight: bold; padding: 10px; border-left: 10px solid ${this.getRepairColor(status)}">
                            ${status.toUpperCase()} (${groupRows.length} Items)
                        </td>
                    </tr>
                </tbody>
                <tbody>`;

            groupRows.forEach(row => {
                const rowData = row.values[0]; 
                
                // RUGGED: Extract actual ID and anchor to the row metadata
                const rawMaeId = (idIndex !== -1) ? rowData[idIndex] : '';

                html += `<tr data-row-index="${row.index}" data-mae-id="${rawMaeId}">`;
                html += `<td class="edit-only-cell"><button class="delete-row-btn" onclick="requestDelete(${row.index})">🗑️</button></td>`;
        
                visibleIndices.forEach(idx => {
                    const colDef = sheetConfig.columns[idx];
                    const isEditable = !colDef.locked && colDef.type !== 'formula';
                    const displayValue = rowData[idx] || ''; 
            
                    html += `<td class="${isEditable ? 'editable-cell' : 'locked-cell'}" data-col-index="${idx}">${displayValue}</td>`;
                });
                html += `</tr>`;
            });
            html += `</tbody>`;
        }
    });

    html += `</table>`;
    container.innerHTML = html;
},

    //=========== END RENDER SUBDIVIDED REPAIRS Specialized Renderer==============

    // 4. STATUS HELPERS
    showLoading(tableName) {
        document.getElementById("table-container").innerHTML = `<div class="loader">Loading ${tableName} data...</div>`;
    },

    showError(message) {
        document.getElementById("table-container").innerHTML = `<p style="color:red; padding:20px;">${message}</p>`;
    },

    getRepairColor(status) {
        const colors = {
            "Needs Repair": "#e74c3c",       // Red
            "Repair In-Progress": "#f1c40f", // Yellow/Gold
            "Unusable/Junk": "#7f8c8d"       // Gray
        };
        return colors[status] || "#34495e";  // Default Navy
    },

    //===== Updated setHealthStatus reflecting the "Read-Only Database" Policy========
    setHealthStatus(isHealthy, firstTableName) {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");

        if (isHealthy) return; 

        title.innerText = "Database Connection Error";
        container.innerHTML = `
            <div style="padding:25px; border-left: 6px solid #e74c3c; background: #fff8f8; color: #2c3e50;">
                <h3 style="margin-top:0; color: #c0392b;">⚠️ System Integrity Alert</h3>
                <p>The application cannot connect to the required data structures. As per your <b>Operational Acknowledgement</b>, manual changes to the Excel file in OneDrive can break the system link.</p>
            
             <p><strong>Common causes for this error:</strong></p>
                <ul style="line-height: 1.6;">
                    <li>The file <b>${maeSystemConfig.spreadsheetName}</b> was renamed or moved out of the OneDrive root.</li>
                    <li>A table header or name (specifically <b>${firstTableName}</b>) was modified manually.</li>
                </ul>

             <hr style="border:0; border-top: 1px solid #eccaca; margin: 20px 0;">
            
                <p><strong>Recovery Path:</strong></p>
                <p>To restore functionality, you must revert any manual changes made to the spreadsheet. If the file is beyond repair, you may <b>Delete</b> it from your OneDrive and <b>Refresh</b> this page. The system will deploy a fresh, formatted database template.</p>
            
                <p style="font-size: 0.85rem; color: #7f8c8d; font-style: italic;">
                    Warning: A "Fresh Start" deployment will permanently erase any inventory data previously stored in the broken file.
                </p>
            </div>`;
},
//======END setHealthStatus =========
//================RENDER COMMAND BAR==========
 
renderCommandBar(tableName) {
    const container = document.getElementById("action-bar-zone");
    if (!container) return;

    // 1. Access the global config
    const config = window.maeSystemConfig; 
    if (!config) return;

    // 2. Find the specific sheet blueprint
    const sheetConfig = config.worksheets.find(s => s.tableName === tableName);
    if (!sheetConfig) return; 

    // Normalize name for reliable comparison
    const normalizedName = tableName.trim().toLowerCase();
    const isDashboard = normalizedName.includes("dashboard");

    // 3. Scan for "Manual" keywords (Quantity / Current Stock)
    const hasManualField = sheetConfig.columns.some(col => 
        col.header === "Quantity" || col.header === "Current Stock"
    );

    // 4. Build the Button HTML
    let buttons = `<button class="action-btn" id="btn-print">Print Table</button>`;

    // Add Manual Log button if keywords match
    if (hasManualField) {
        buttons += `<button class="action-btn" id="btn-manual-print">Print Manual Log</button>`;
    }

    // Only show "Action" buttons if we are NOT on a dashboard
    if (!isDashboard) {
        
        // Define the specific five tables allowed to show the Scan button (in lowercase)
        const scannableTables = [
            "shop_machinery", 
            "shop_power_tools", 
            "shop_hand_tools", 
            "shop_consumables", 
            "resell_inventory"
        ];
        
        // Check if current table is scannable AND feature is enabled in config
        if (scannableTables.includes(normalizedName) && config.features?.enableScanning) {
            buttons += `<button class="action-btn" id="btn-scan" style="background:#8e44ad;">📷 Scan Item</button>`;
        }

        buttons += `
            <button class="action-btn" id="btn-add">Add Item</button>
            <button class="action-btn" id="btn-edit">Edit Table</button>
        `;

        // Only show Quick Update if it's an inventory-style sheet
        if (hasManualField) {
            buttons += `<button class="action-btn" id="btn-inventory-update">Quick Update</button>`;
        }
    }

    // 5. Inject into the UI
    container.innerHTML = `<div class="command-bar">${buttons}</div>`;
},

//========== END RENDER COMMAND BAR ================

// ================RENDER ENTRY FORM===============
   renderEntryForm(mode, tableName, sheetConfig, onSaveCallback, rowIndex = null, existingData = null) {
    const container = document.getElementById("table-container");
    const isEdit = mode === 'edit';
    
    let formHtml = `
        <div class="form-card" id="entry-form">
            <div class="form-header">
                <h3>${isEdit ? 'Edit' : 'Add New'} Entry: ${sheetConfig.tabName}</h3>
                <button class="close-x" onclick="document.getElementById('entry-form').remove()">×</button>
            </div>
            <div class="form-grid">`;

    sheetConfig.columns.forEach((col, index) => {
        const fieldId = `field-${col.header.replace(/\s+/g, '')}`;
        const val = (isEdit && existingData) ? existingData[index] : "";

        // RUGGED LOGIC: If the field is hidden or a formula, we still need the input 
        // in the DOM for the data to be submitted, but we hide it from the user.
            if (col.hidden || col.type === "formula") {
                formHtml += `<input type="hidden" id="${fieldId}" value="${val}">`;
            } 
            else {
                // Render visible fields normally
                formHtml += `<div class="input-group"><label>${col.header}</label>`;

                if (col.type === "boolean") {
                    // Render as a checkbox for boolean types
                    const isChecked = val.toString().toUpperCase() === "TRUKE";
                    formHtml += `<input type="checkbox" id="${fieldId}" ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
                }

                else if (col.type === "dropdown") {
                    // INVIOLATE: Forces user to pick from config options only
                    formHtml += `
                        <select id="${fieldId}" required>
                            <option value="">-- Select ${col.header} --</option>
                            ${col.options.map(opt => 
                                `<option value="${opt}" ${opt == val ? 'selected' : ''}>${opt}</option>`
                            ).join('')}
                        </select>`;
                } 
                else if (col.type === "number") {
                    const isCurrency = col.format && col.format.includes("$");
                
                    if (isCurrency) {
                        formHtml += `
                            <input type="number" step="0.01" id="${fieldId}" value="${val}" 
                                placeholder="0.00" 
                                onblur="if(this.value) this.value = parseFloat(this.value).toFixed(2)">`;
                    } else {
                        formHtml += `
                            <input type="number" step="1" id="${fieldId}" value="${val}" 
                                placeholder="Whole number only"
                                onkeydown="if(['.', ',', 'e', 'E'].includes(event.key)) event.preventDefault();"
                                onblur="if(this.value) this.value = Math.floor(this.value)">`;
                    }
                } 
                else if (col.type === "date") {
                    formHtml += `<input type="date" id="${fieldId}" value="${val}">`;
                } 
                else {
                    formHtml += `<input type="text" id="${fieldId}" value="${val}" placeholder="Enter ${col.header}...">`;
                }
                formHtml += `</div>`;
            }
        }); 

    formHtml += `</div>
        <div class="form-actions">
            <button class="save-btn" id="submit-form-btn">${isEdit ? 'Update' : 'Save'} to OneDrive</button>
            <button class="cancel-btn" onclick="document.getElementById('entry-form').remove()">Cancel</button>
        </div>
    </div>`;

    container.insertAdjacentHTML('beforebegin', formHtml);

     // FIX: Define the button AFTER it is injected into the DOM
    const submitBtn = document.getElementById("submit-form-btn");

    if (submitBtn) {
        submitBtn.onclick = async () => {
            // RUGGED LOCK
            submitBtn.disabled = true;
            submitBtn.innerText = "Saving to OneDrive...";
            submitBtn.style.opacity = "0.5";
            submitBtn.style.cursor = "not-allowed";

            // RUN THE SAVE: Call the callback from app.js
            // We 'await' it so we know when the network request is finished
            await onSaveCallback(rowIndex, existingData); 

            // UNLOCK (If the form wasn't removed by the callback)
            if (document.getElementById("submit-form-btn")) {
                submitBtn.disabled = false;
                submitBtn.innerText = isEdit ? 'Update to OneDrive' : 'Save to OneDrive';
                submitBtn.style.opacity = "1";
                submitBtn.style.cursor = "pointer";
            }
            
        }
    }
},
//======= END RENDER ENTRY FORM ============

//========== EXIT EDIT MODE ==============
    exitEditMode() {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    // 1. Reset Global Table States
    table.classList.remove("is-editing", "is-quick-updating", "saving-active");
    table.style.opacity = "1";
    table.style.pointerEvents = "auto";

    // 2. Clean up any active forms
    const entryForm = document.getElementById("entry-form");
    if (entryForm) entryForm.remove();

    // 3. THE FIX: Process every cell to remove inputs and dropdowns
    const cells = table.querySelectorAll("td");
    cells.forEach(cell => {
        // --- A. Handle Number Inputs (The "Sticky" Culprit) ---
        const input = cell.querySelector('input');
        if (input) {
            cell.innerText = input.value; // Capture the final number
        }

        // --- B. Handle Dropdowns ---
        const select = cell.querySelector('select');
        if (select) {
            cell.innerText = select.value; // Capture the final choice
        }

        // --- C. Reset Visual & Functional States ---
        cell.contentEditable = "false";
        cell.style = ""; // Wipes all inline z-index, background-colors, etc.
        cell.classList.remove("quick-edit-focus", "dropdown-edit-zone", "text-edit-focus");
        
        // Remove all temporary listeners
        cell.onclick = null;
        cell.onkeydown = null;
        cell.onblur = null;
    });

    console.log("MAE System: UI Sanitized. All inputs removed.");
},

//===== END EXiT EDIT MODE ==========

//=====PRINT TABLE ===========

// ui.js - Updated printTable function
printTable(tableName, sheetConfig, customTitle = null) {
    // Target the main content area so the title stays aligned with the table
    const container = document.getElementById("app-content");
    if (!container) return;

    // Use custom title if provided, otherwise fallback to config
    const finalTitle = customTitle || `MAE Workshop Inventory System: ${sheetConfig.tabName}`;


    // Create the temporary print header
    const printHeader = document.createElement("div");
    printHeader.className = "print-only-title";
    
    // RUGGED: Simple, clear branding for the hardcopy
    printHeader.innerHTML = `
        <h1>${finalTitle}</h1>
    `;
    
    // 1. Inject at the top of the content zone
    container.prepend(printHeader);

     // Ensure we aren't in "Edit Mode" visuals during print
    this.exitEditMode(); 

    // 2. Trigger the browser print dialog
    window.print();

    // 3. Clean up immediately so it doesn't show on the screen
    printHeader.remove();
},
//====== END PRINT TABLE ============

//============ print MANUAL LOG =============
// ui.js - New Function
printManualLog(tableName, sheetConfig) {
    const container = document.getElementById("app-content");
    const table = document.getElementById("main-data-table");
    if (!table || !container) return;

    // 1. Mark the table for manual log styling
    table.classList.add("manual-log-mode");

    const printHeader = document.createElement("div");
    printHeader.className = "print-only-title";
    printHeader.innerHTML = `<h1>MANUAL INVENTORY LOG: ${sheetConfig.tabName}</h1>`;
    
    container.prepend(printHeader);
    UI.exitEditMode(); //prevents extraneous css styling if print function called from Edit Table
    window.print();
    
    // 2. Clean up
    printHeader.remove();
    table.classList.remove("manual-log-mode");
},

//========== END PRINT MANUAL LOG ==============

//========== RENDER DASHBOARD ==========
// Inside UI object in ui.js
/**
 * UI.renderDashboard - MAE Master Dashboard
 * Mapping: Hero numbers pulled from the single-row Master_Dashboard table.
 */
renderDashboard(row, config) {
    // save data to a temporary "state" so Chart function can see it
    this.currentDashboardData = Dashboard.parseSummary(row, config);

    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    
    const dashboardData = Dashboard.parseSummary(row, config);
    const maintenanceDue = dashboardData["Maintenance Items Due in Next 30 Days"];
    title.innerText = "Workshop Master Dashboard";
    

    container.innerHTML = `
        <div class="dashboard-grid">
            
            <!-- Snapshot A: Resell Inventory (Drill down to WIP/For Sale/Complete) -->
            <div class="dash-card" onclick="loadTableData('Resell_Inventory')">
                <h4>Resell Inventory</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Actual Sales"])}</div>
                <p>Total Actual Sales</p>
                <div class="card-sub-actions">
                    <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Resell_Inventory', 'resell-active')">
                        View WIP, Complete & For Sale
                    </button>
                    <small style="display:block; margin-top:5px; color: #7f8c8d;">
                        Total Invested: ${formatCurrency(dashboardData["Total Resell Investment"])}
                    </small>
                </div>
            </div>

            <!-- Snapshot B: Total Asset Value (Drill down pie chart of Assets) -->
            <div class="dash-card">
                <h4>Total Shop Assets</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Shop Asset Value"])}</div>
                <p>Machinery, Tools & Supplies</p>
                <div class="card-sub-actions">
                    <button class="action-btn" onclick="UI.showAssetBreakdown()">📊 View Asset Breakdown</button>
                </div>
            </div>


            <!-- Snapshot C: Low Stock Alerts (Drill down to ONLY Low Stock) -->
            <div class="dash-card ${dashboardData["Low Stock Items Count"] > 0 ? 'alert' : ''}" 
                 onclick="loadTableData('Shop_Consumables', 'low-stock')">
                <h4>Low Stock Alerts</h4>
                <div class="hero-num">${dashboardData["Low Stock Items Count"]}</div>
                <p>Items below reorder point</p>
                <small>Click to view shopping list</small>
            </div>

            <!-- Snapshot E: Overhead Snapshot-->
            <div class="dash-card" onclick="loadTableData('Shop_Overhead')">
                <h4>Total Amount Due Next 30 Days</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Amount Due Next 30 Days"])}</div>
    
                <div class="card-sub-actions">
                    <!-- New Annual Breakdown Button -->
                    <button class="action-btn" onclick="event.stopPropagation(); UI.showAnnualOverhead()">
                        📊 Annual Overhead Breakdown
                    </button>
        
                    <p style="margin-top:15px; font-weight:bold; font-size:0.9rem;">Upcoming Bills by Time Period:</p>
                    <div style="display: flex; flex-wrap: wrap; gap: 5px; justify-content: center;">
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Overhead', 'due-7')">Next 7 Days</button>
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Overhead', 'due-30')">Next 30 Days</button>
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Overhead', 'due-90')">Next 90 Days</button>
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Overhead', 'due-180')">Next 180 Days</button>
                    </div>
                </div>
            </div>
            

            <!-- Snapshot F: Equipment Operational Issues -->
            <div class="dash-card ${dashboardData["Equipment With Operational Issues"] > 0 ? 'warning' : ''}">
                <h4>Equipment With Operational Issues</h4>
                <div class="hero-num">${dashboardData["Equipment With Operational Issues"]}</div>
                <p>Total Items Needing Attention</p>
    
                <div class="card-sub-actions" style="display: flex; flex-direction: column; gap: 8px; margin-top: 15px;">
                    <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Machinery', 'needs-repair')">
                        Shop Machinery: Operational Issues
                    </button>
                    <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Power_Tools', 'needs-repair')">
                        Shop Power Tools: Operational Issues
                    </button>
                    <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Shop_Hand_Tools', 'needs-repair')">
                        Shop Hand Tools: Operational Issues
                    </button>
                    <small style="color: #7f8c8d; margin-top: 5px;">Click a category to view specific repair lists</small>
                </div>
            </div>
      
            <!-- Snapshot G: Maintenance Items -->
            <div class="dash-card ${maintenanceDue > 0 ? 'warning' : ''}">
                <h4>Upcoming Maintenance</h4>
                <div class="hero-num">${maintenanceDue}</div>
                <p>Items Due in Next 30 Days</p>
                <div class="card-sub-actions">
                    <div style="display: flex; flex-wrap: wrap; gap: 5px; justify-content: center;">
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Maintenance_Log', 'maint-7')">Next 7 Days</button>
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Maintenance_Log', 'maint-30')">Next 30 Days</button>
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Maintenance_Log', 'maint-90')">Next 90 Days</button>
                        <button class="mini-btn" onclick="event.stopPropagation(); loadTableData('Maintenance_Log', 'maint-180')">Next 180 Days</button>
                    </div>
                    <small style="display:block; margin-top:10px; color: #7f8c8d;">
                        Stay ahead of shop downtime.
                    </small>
                </div>
            </div>

        </div>
    `;
},
//====END RENDER DASHBOARD =============

//==== FOR CHART.JS ===============

showAssetBreakdown() {
    const data = this.currentDashboardData;
    if (!data) return;

    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    title.innerText = "SHOP ASSETS: Value Distribution";
    
    // 1. Fixed height for a massive "Hero" view
    // The button is now absolutely positioned to the top-left
    container.innerHTML = `
        <div style="width: 100%; height: 500px; position: relative; padding: 10px;">
            <button class="cancel-btn" 
                    onclick="loadTableData('Master_Dashboard')" 
                    style="position: absolute; top: 10px; left: 10px; z-index: 1000; padding: 8px 15px;">
                ← Back
            </button>
            <canvas id="assetChart"></canvas>
        </div>`;

    const ctx = document.getElementById('assetChart').getContext('2d');
    
    new Chart(ctx, {
    type: 'pie',
    plugins: [ChartDataLabels],
    data: {
        labels: ['Machinery', 'Power Tools', 'Hand Tools', 'Consumables'],
        datasets: [{
            data: [
                data["Total Machinery Value"],
                data["Total Power Tool Value"],
                data["Total Hand Tool Value"],
                data["Total Consumables Value"]
            ],
            backgroundColor: ['#2c3e50', '#d35400', '#27ae60', '#2980b9'],
            borderWidth: 3
        }]
    },
    options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
            padding: { left: 10, right: 30, top: 20, bottom: 20 }
        },
        plugins: {
            legend: {
                display: true,
                position: 'left', // Legend on left to maximize pie size
                align: 'center',
                labels: {
                    font: { weight: 'bold', size: 14 },
                    padding: 20,
                    boxWidth: 20
                }
            },
            datalabels: {
                color: '#fff',
                font: { weight: 'bold', size: 12 },
                // Use percentages + values for high-speed situational awareness
                formatter: (value, ctx) => {
                    let sum = 0;
                    let dataArr = ctx.chart.data.datasets[0].data;
                    dataArr.map(d => { sum += d; });
                    let percentage = (value * 100 / sum).toFixed(0) + "%";
                    // Only render labels for non-zero values to prevent overlap
                    return value > 0 ? `${percentage}\n${formatCurrency(value)}` : null;
                },
                anchor: 'center',
                align: 'center',
                textAlign: 'center',
                textShadowColor: 'rgba(0,0,0,0.6)',
                textShadowBlur: 4
                }
            }
        }
    });
//========= END ASSET BREAKDOWN ==============
  },
// ========= SHOW ANNUAL OVERHEAD =============
async showAnnualOverhead() {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    title.innerText = "ANNUAL OVERHEAD: Expenses by Category";

    UI.showLoading("Fetching Overhead Calculations...");

    try {
        const tableName = "Overhead_Summary";
        const sheetConfig = maeSystemConfig.worksheets.find(s => s.tableName === tableName);
        
        // 1. Fetch the rows directly from your new Excel Summary Table
        const tableData = await Dashboard.getFullTableData(tableName); 

        // 2. DYNAMIC INDEX DETECTION
        const catIdx = sheetConfig.columns.findIndex(c => c.header === "Expense Category");
        const valIdx = sheetConfig.columns.findIndex(c => c.header === "Annual Total");

        // 3. Extract labels and values (Graph API returns values in a nested array: values[0][index])
        //const labels = tableData.map(row => row.values[catIdx]);
        //const values = tableData.map(row => parseFloat(row.values[valIdx]) || 0);
        //const labels = tableData.map(row => String(row.values[0][catIdx] || ""));
        //const values = tableData.map(row => parseFloat(row.values[0][valIdx]) || 0);

        // RUGGED: Check if row.values[0] exists before mapping to prevent empty bars
        const labels = tableData
            .filter(row => row.values && row.values[0]) 
            .map(row => String(row.values[0][catIdx] || ""));

        const values = tableData
            .filter(row => row.values && row.values[0])
            .map(row => parseFloat(row.values[0][valIdx]) || 0);


        container.innerHTML = `
            <div style="width: 100%; height: 500px; position: relative; padding: 20px;">
                <button class="cancel-btn" onclick="loadTableData('Master_Dashboard')" 
                        style="position: absolute; top: 10px; left: 10px; z-index: 1000; padding: 8px 15px;">
                    ← Back
                </button>
                <canvas id="overheadChart"></canvas>
            </div>`;

        const ctx = document.getElementById('overheadChart').getContext('2d');

        // 4. THE RUGGED CLEANUP: Destroy existing chart to prevent "ghost" overlaps
        if (window.myChart) {
            window.myChart.destroy();
        }
        

        
        // 5. THE ASSIGNMENT: Initialize and store the chart instance
        window.myChart = new Chart(ctx, {
            type: 'bar',
            data: {
                // Shorten labels for clean display; handles the split-view on the axis
                labels: labels.map(l => String(l || "").split('/')), 
                datasets: [{
                    label: 'Annual Cost',
                    data: values,
                    backgroundColor: '#2c3e50', // Deep Navy
                    borderColor: '#d35400',     // Safety Orange
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { 
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (context) => `Total: ${formatCurrency(context.raw)}`
                        }
                    }
                },
                scales: {
                    y: { 
                        beginAtZero: true, 
                        ticks: { 
                            callback: (value) => '$' + value.toLocaleString(),
                            font: { weight: 'bold' } 
                        } 
                    },
                    x: {
                        ticks: { font: { size: 11 } }
                    }
                }
            }
        });
    } catch (error) {
        console.error("MAE System: Chart Load Error", error);
        UI.showError("Could not load overhead breakdown. Ensure Excel formulas are intact.");
    }
},
//===== END SHOW ANNUAL OVERHEAD ===========


//============ INDUSTRIAL SCAN RESULT UI (Rugged Tablet Optimized) =============

renderScanResultCard(rowData, tableName, sheetConfig, rowIndex) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
        
    // Define headers for high situational awareness on the shop floor
    const industrialManifest = {
        "Resell_Inventory": ["Asset ID", "Item Name", "Current Status", "Target Sale Price"],
        "Shop_Machinery": ["Machine Name/Model", "Manufacturer/Brand", "Serial Number", "Condition"],
        "Shop_Power_Tools": ["Tool Name/Model", "Power Source", "Condition"],
        "Shop_Hand_Tools": ["Tool Name/Model", "Category", "Condition", "Quantity"],
        "Shop_Consumables": ["Item Name", "Current Stock", "Reorder Point", "Unit of Measure"]
    };

    const allowedHeaders = industrialManifest[tableName] || [];
    title.innerText = `Scan Found: ${sheetConfig.tabName}`;

    let html = `
        <div class="industrial-card" style="padding: 30px; border-top: 8px solid var(--accent); background: #fff; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px;">
    `;
        
    sheetConfig.columns.forEach((col, index) => {
        if (allowedHeaders.includes(col.header)) {
            let displayVal = rowData[index] ?? "---";
                
            if (col.format && col.format.includes("$")) {
                displayVal = formatCurrency(displayVal);
            }

            html += `
                <div class="data-point">
                    <label style="display:block; font-size: 0.85rem; color: #7f8c8d; font-weight: bold; text-transform: uppercase;">${col.header}</label>
                    <div style="font-size: 1.5rem; font-weight: 700; color: var(--primary);">${displayVal}</div>
                </div>`;
        }
    });

    html += `</div>`; 

    html += `
        <div style="display: flex; gap: 15px;">
            <button class="action-btn" 
                    style="flex: 2; height: 70px; font-size: 1.3rem; background: var(--accent);" 
                    onclick="UI.openEditFormFromScan('${tableName}', ${JSON.stringify(rowData).replace(/"/g, '&quot;')}, ${rowIndex})">
                ✏️ Update Record
            </button>
            <button class="action-btn" 
                    style="flex: 1; height: 70px; font-size: 1.3rem; background: #7f8c8d;" 
                    onclick="window.loadTableData('Master_Dashboard')">
                Done
            </button>
        </div>
    </div>`;

    container.innerHTML = html;
    this.renderCommandBar(tableName);
},
openEditFormFromScan(tableName, rowData, rowIndex) {
    const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === tableName);
    
    this.renderEntryForm('edit', tableName, sheetConfig, async () => {
        // CALL THE NEW SINGLE ROW UPDATE INSTEAD OF THE TABLE SCRAPER
        const success = await window.updateSingleRowFromForm(tableName, rowIndex, sheetConfig);
        if (success) {
            this.exitEditMode();
            window.loadTableData(tableName);
        }
    }, rowIndex, rowData); 
}
//========= END SCAN RESULT UI LOGIC =============


};

window.UI = UI;

