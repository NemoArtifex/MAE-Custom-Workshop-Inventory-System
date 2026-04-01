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
    renderTable(rows, tableName, sheetConfig, customTitle= null) {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        
        if (!sheetConfig) {
            container.innerHTML = "Error: Worksheet configuration not found.";
            return;
        }

        title.innerText = customTitle || `View: ${sheetConfig.tabName}`;

        // 1. Identify visible columns from the Manifest (config.js)
        const visibleIndices = [];
        let html = `<table class="inventory-table" id="main-data-table"><thead><tr>`;
        
        // Add "Delete" Header
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
                // RUGGED: Use the persistent 'index' from Graph API row object
                // This ensures we always update the correct row in Excel
                const persistentIndex = row.index; 

                html += `<tr data-row-index="${persistentIndex}">`;

                // Add Delete Icon Cell using the persistent index
                html += `<td class="edit-only-cell">
                            <button class="delete-row-btn" onclick="requestDelete(${persistentIndex})">🗑️</button>
                         </td>`;
                
                // Extract cell data
                const allCells = Array.isArray(row.values[0]) ? row.values[0] : row.values; 

                visibleIndices.forEach(idx => {
                    const colDef = sheetConfig.columns[idx];
                    const isEditable = !colDef.locked && colDef.type !== 'formula';
                    const isQuantity = colDef.header === "Quantity" || colDef.header === "Current Stock";
                    
                    // RUGGED: Identify if this is a currency column for CSS styling
                    const isCurrency = colDef.format && colDef.format.includes("$");
                    
                    let displayValue = allCells[idx] ?? '';

                    // Format visually for the UI
                    if (isCurrency) {
                        displayValue = formatCurrency(displayValue);
                    }

                    // 3. Build the Cell with specific MAE classes
                    html += `<td 
                            class="${isEditable ? 'editable-cell' : 'locked-cell'} 
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

    // 4. STATUS HELPERS
    showLoading(tableName) {
        document.getElementById("table-container").innerHTML = `<div class="loader">Loading ${tableName} data...</div>`;
    },

    showError(message) {
        document.getElementById("table-container").innerHTML = `<p style="color:red; padding:20px;">${message}</p>`;
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
        

    const normalizedName = tableName.trim().toLowerCase();
    const dashboardTables = ["master_dashboard", "test_dashboard", "master dashboard", "test dashboard"];

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

    // Add Add/Edit/Quick Update for non-dashboard tables
    if (!dashboardTables.includes(normalizedName)) {
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
        // RUGGED: Skip ID, Hidden, and Formulas
        if (!col.hidden && col.type !== "formula") {
            const fieldId = `field-${col.header.replace(/\s+/g, '')}`;
            const val = (isEdit && existingData) ? existingData[index] : "";

            formHtml += `<div class="input-group"><label>${col.header}</label>`;

            if (col.type === "dropdown") {
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
                    // CURRENCY: Allows decimals, forces 2-decimal format on blur
                    formHtml += `
                        <input type="number" step="0.01" id="${fieldId}" value="${val}" 
                            placeholder="0.00" 
                            onblur="if(this.value) this.value = parseFloat(this.value).toFixed(2)">`;
                } else {
                    // INTEGER (Rugged): Blocks decimal/scientific keys, floors any pasted values
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

            <!-- Snapshot E: Monthly Overhead (Drill down to Full Overhead) -->
            <div class="dash-card" onclick="loadTableData('Shop_Overhead')">
                <h4>Monthly Overhead</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Monthly Overhead"])}</div>
                <p>Total Fixed Costs</p>
                <small style="color: #7f8c8d;">Click to manage expenses</small>
            </div>

            <!-- Snapshot F: Equipment Repairs (Drill down to ONLY broken tools) -->
            <div class="dash-card ${dashboardData["Equipment Needing Repair"] > 0 ? 'warning' : ''}" 
                 onclick="loadTableData('Shop_Machinery', 'needs-repair')">
                <h4>Repairs Needed</h4>
                <div class="hero-num">${dashboardData["Equipment Needing Repair"]}</div>
                <p>Out-of-Service Items</p>
                <small>Click to view repair list</small>
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
    
    
    container.innerHTML = `
    <div style="width: 100%; height: 350px; position: relative; margin: 0 auto; padding: 10px;">
        <canvas id="assetChart"></canvas>
        <div style="text-align:center; margin-top:10px;">
            <button class="cancel-btn" onclick="loadTableData('Master_Dashboard')">Back to Dashboard</button>
        </div>
    </div>`;

    const ctx = document.getElementById('assetChart').getContext('2d');
    
    new Chart(ctx, {
        type: 'pie',
        plugins: [ChartDataLabels], // Only load the compatible plugin
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
                borderWidth: 2,
                // RUGGED: This pushes the pie slightly inward to make room for labels
                hoverOffset: 20 
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            aspectRatio: 1,
            layout: {
                padding: 45 
            },
            plugins: {
                legend: { position: 'bottom' },
                datalabels: {
                    // POSITIONING: This spaces them out around the perimeter
                    anchor: 'end',
                    align: 'end',
                    offset: 10,
                    
                    color: '#2c3e50',
                    font: { weight: 'bold', size: 12 },
                    textAlign: 'center',
                    
                    // FORMATTER: Separates label and value with a newline
                    formatter: (value, context) => {
                        const label = context.chart.data.labels[context.dataIndex];
                        return label + '\n' + formatCurrency(value);
                }
            }
        }
    });
//========= END ASSET BREAKDOWN ==============
  }
};

window.UI = UI;

