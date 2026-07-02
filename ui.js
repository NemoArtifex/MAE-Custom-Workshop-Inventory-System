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
        
    // 🌟 AUTHENTICATION DISCONNECT GUARD SHIELD 🌟
        authButton.onclick = () => {
        const isProtectedSessionActive = 
            window.currentTable === "inventory_registration" || 
            window.currentTable === "untagged_audit_grid_view" ||
            window.currentTable === "resell_status_pivot" ||
            window.currentTable === "location_inspector";

        if (isProtectedSessionActive) {
            const confirmSignOut = confirm("MAE SYSTEM DISCONNECT WARNING:\n\nYou are currently inside an active batch registration or compliance audit session.\n\nDisconnecting your Microsoft Office 365 link right now will drop your workspace state and discard all uncommitted items on your screen.\n\nAre you sure you want to sign out?");
            if (!confirmSignOut) {
            return; // Abort sign-out request, keep session locked
            }
        }
      
        // Safe boundary confirmed, proceed to sign out macro execution
        signOutCallback();
        };
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
        // 🌟 SIDEBAR NAVIGATION SHIELD 🌟
        const isProtectedSessionActive = 
            window.currentTable === "inventory_registration" || 
            window.currentTable === "untagged_audit_grid_view" ||
            window.currentTable === "resell_status_pivot" ||
            window.currentTable === "location_inspector";

        if (isProtectedSessionActive) {
            const confirmExit = confirm("MAE SYSTEM WARNING:\n\nYou are currently inside an active batch session or filtered workspace. Navigating away right now will discard any unsaved entries on your screen.\n\nAre you sure you want to exit this view?");
            if (!confirmExit) {
            return; // Abort navigation change, keep session pinned!
            }
        }
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
           // 🌟 SIDEBAR NAVIGATION SHIELD 🌟
        const isProtectedSessionActive = 
            window.currentTable === "inventory_registration" || 
            window.currentTable === "untagged_audit_grid_view" ||
            window.currentTable === "resell_status_pivot" ||
            window.currentTable === "location_inspector";

        if (isProtectedSessionActive) {
            const confirmExit = confirm("MAE SYSTEM WARNING:\n\nYou are currently inside an active batch session or filtered workspace. Navigating away right now will discard any unsaved entries on your screen.\n\nAre you sure you want to exit this view?");
            if (!confirmExit) {
            return; // Abort navigation change, keep session pinned!
            }
        }
                
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

        const isRepairsView = customTitle && customTitle.includes("Operational Issues");
        if (isRepairsView) {
            this.renderSubdividedRepairs(rows, tableName, sheetConfig);
           return; 
        }

        const idIndex = sheetConfig.columns.findIndex(c => c.header === "mae_id");
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

        if (rows && rows.length > 0) {
            rows.forEach((row) => {
                const persistentIndex = row.index; 
                const allCells = row.values; 

                if (!allCells || allCells.length === 0) return;

                const rawMaeId = (idIndex !== -1) ? allCells[idIndex] : '';

                html += `<tr data-row-index="${persistentIndex}" data-mae-id="${rawMaeId}">`;
                html += `<td class="edit-only-cell">
                            <button class="delete-row-btn" onclick="requestDelete(${persistentIndex})">🗑️</button>
                        </td>`;

                visibleIndices.forEach(idx => {
                    const colDef = sheetConfig.columns[idx];
                    //  Allow Tag_Type to toggle, but explicitly freeze Tag_ID columns during inline edits
                    const isEditable = (!colDef.locked || colDef.header === "Tag_Type") && colDef.header !== "Tag_ID" && colDef.type !== 'formula';
                    let displayValue = allCells[idx] ?? '';

                    // --- 🌟 NEW: VALUATION GOVERNANCE SILO GUARD ---
                    // Prevents "Methodology Bleed" by hiding abandoned data
                    if (tableName === "Shop_Consumables") {
                        const levelIdx = sheetConfig.columns.findIndex(c => c.header === "Stock_Level");
                        const currentMethod = (allCells[levelIdx] || "").toString().trim();
                    
                        const isBulkMode = ["None", "Few", "Adequate", "Many"].includes(currentMethod);
                        const isCountedMode = currentMethod === "Counted";

                        const unitSiloHeaders = ["Unit Cost", "Stock_Count","Reorder Point"];
                        const bulkSiloHeaders = ["Bulk_Value"];

                        // Wipe data for the inactive silo to avoid $0.00 confusion
                        if (isBulkMode && unitSiloHeaders.includes(colDef.header)) {
                            displayValue = ""; 
                        } else if (isCountedMode && bulkSiloHeaders.includes(colDef.header)) {
                            displayValue = "";
                        }
                    }

                    const isCurrency = colDef.format && colDef.format.includes("$");
                    const isLowStockText = (displayValue === "Few" || displayValue === "None");
                    const isSubjective = ["Few", "Adequate", "Many"].includes(displayValue);

                    let sellPriceAlert = "";
                    if (tableName === "Resell_Inventory" && colDef.header === "Actual Sale Price") {
                        const statusIdx = sheetConfig.columns.findIndex(c => c.header === "Current Status");
                        const itemStatus = allCells[statusIdx];
                        const itemPrice = parseFloat(displayValue.toString().replace(/[^0-9.-]+/g, "")) || 0;

                        if (itemStatus === "Sold" && itemPrice <= 0) {
                            sellPriceAlert = "col-resell-price-missing";
                        }
                    }

                    let overdueClass = "";
                    if (colDef.type === 'date' && displayValue && displayValue !== "") {
                        const dueDate = new Date(displayValue);
                        const today = new Date();
                        today.setHours(0, 0, 0, 0);
                        const isDeadlineCol = colDef.header.includes("Due") || colDef.header.includes("Service");

                        if (isDeadlineCol && !isNaN(dueDate) && dueDate < today) {
                            overdueClass = "col-date-overdue";
                        }
                    }

                    // Only format currency if the silo hasn't been wiped to blank
                    if (isCurrency && displayValue !== "") {
                        displayValue = formatCurrency(displayValue);
                    }

                    if (colDef.type === 'boolean') {
                        const isChecked = displayValue.toString().toUpperCase() === "TRUE";
                        displayValue = `<input type="checkbox" disabled ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
                    }

                    let stockAlertClasses = "";
                    if (tableName === "Shop_Consumables") {
                        if (colDef.header === "Stock_Count") stockAlertClasses = "col-type-stock-alert";
                        if (colDef.header === "Reorder Point") stockAlertClasses = "col-type-reorder-point";
                    }

                    html += `<td class="${isEditable ? 'editable-cell' : 'locked-cell'} 
                                    ${overdueClass}
                                    ${sellPriceAlert}
                                    ${stockAlertClasses}
                                    ${isLowStockText ? 'col-stock-alert-red' : ''}
                                    ${isSubjective ? 'col-subjective-level' : ''}
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
    //=======  refreshSiloLocks ===========
    refreshSiloLocks(row, tableName, sheetConfig) {
        if (tableName !== "Shop_Consumables") return;

        const levelIdx = sheetConfig.columns.findIndex(c => c.header === "Stock_Level");
        const levelCell = row.querySelector(`td[data-col-index="${levelIdx}"]`);
        if (!levelCell) return;

        const currentMethod = (levelCell.innerText || "").trim();
        const isCounted = currentMethod === "Counted";

        const unitSilo = ["Unit Cost", "Stock_Count", "Reorder Point"];
        const bulkSilo = ["Bulk_Value"];

        sheetConfig.columns.forEach((col, idx) => {
            const targetCell = row.querySelector(`td[data-col-index="${idx}"]`);
            if (!targetCell) return;

            const isUnitCol = unitSilo.includes(col.header);
            const isBulkCol = bulkSilo.includes(col.header);

            if ((isCounted && isBulkCol) || (!isCounted && isUnitCol)) {
                // 🌟 RUGGED WIPE: Clear the text content so the harvester sends 'null' to Excel
                targetCell.innerHTML = ""; 
            
                targetCell.classList.add('silo-locked');
                targetCell.style.pointerEvents = "none";
                targetCell.style.opacity = "0.5";
                targetCell.contentEditable = "false";
            } 
            else if (isUnitCol || isBulkCol) {
                targetCell.classList.remove('silo-locked');
                targetCell.style.pointerEvents = "auto";
                targetCell.style.opacity = "1";
        
                if (window.isEditing === true && !targetCell.querySelector('input')) {
                    const isCurrency = col.format && col.format.includes("$");
                    const currentVal = targetCell.innerText.replace(/[^0-9.-]+/g, "") || 0;
            
                    targetCell.innerHTML = `
                        <input type="number" 
                            class="edit-number-input" 
                            value="${currentVal}" 
                            step="${isCurrency ? '0.01' : '1'}">`;
                }
            }
        });
        console.log(`MAE System: Silo state enforced for ${currentMethod} mode.`);
    },
    //====== END refreshSiloLocks ==========

    //=========== RENDER SUBDIVIDED REPAIRS Specialized Renderer ==============
    renderSubdividedRepairs(rows, tableName, sheetConfig) {
    const container = document.getElementById("table-container");
    const conditions = ["Needs Repair", "Repair In-Progress", "Unusable/Junk"];

    // 1. PRIMARY KEY INTEGRITY: Find index of mae_id
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
            const rowCells = r.values; 
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
                const rowData = row.values; 
                
                // 2. RUGGED ANCHOR: Attach ID to the row attribute
                const rawMaeId = (idIndex !== -1) ? rowData[idIndex] : '';

                html += `<tr data-row-index="${row.index}" data-mae-id="${rawMaeId}">`;
                html += `<td class="edit-only-cell"><button class="delete-row-btn" onclick="requestDelete(${row.index})">🗑️</button></td>`;
        
                visibleIndices.forEach(idx => {
                    const colDef = sheetConfig.columns[idx];
                    const isEditable = !colDef.locked && colDef.type !== 'formula';
                    let displayValue = rowData[idx] || ''; 
            
                    // 3. BOOLEAN LOGIC: Match renderTable's checkbox conversion
                    if (colDef.type === 'boolean') {
                        const isChecked = displayValue.toString().toUpperCase().trim() === "TRUE";
                        displayValue = `<input type="checkbox" disabled ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
                    }

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

        const normalizedName = tableName.trim().toLowerCase();
        const config = window.maeSystemConfig;
        const sheetConfig = config.worksheets.find(s => s.tableName === tableName);
        const table = document.getElementById("main-data-table");
        const isQuickUpdating = table && table.classList.contains("is-quick-updating");

        if (window.isEditing || isQuickUpdating) {
            let buttons = "";
            if (isQuickUpdating && sheetConfig) {
                buttons += `<button class="action-btn" onclick="UI.printManualLog('${tableName}', maeSystemConfig.worksheets.find(s => s.tableName === '${tableName}'))" style="background:#34495e;">🖨️ Print Manual Log</button>`;
            }
            buttons += `
                <button class="action-btn" id="btn-commit-sync" style="background:#27ae60;">💾 Commit Changes</button>
                <button class="action-btn" id="btn-discard-edit" style="background:#7f8c8d;">Discard Changes</button>
            `;
            container.innerHTML = `<div class="command-bar">${buttons}</div>`;
            return;
        }

        if (normalizedName === "location_audit") {
            container.innerHTML = `<div class="command-bar">
                <button class="action-btn" onclick="loadTableData('Location')">← Back to Map</button>
                <button class="action-btn" id="btn-print-audit">Print TBD Audit</button>
            </div>`;
            return;
        }

        if (normalizedName === "location_inspector" || window.currentTable === "location_inspector") {
            container.innerHTML = `
                <div class="command-bar" style="justify-content: center;">
                    <button class="action-btn" onclick="UI.printInspectedLocationTable()" style="background:#27ae60;">🖨️ Print Inspected View</button>
                    <button class="action-btn" onclick="window.loadTableData('Location')">← Back to Location Map</button>
                </div>`;
            return;
        }

        if (normalizedName === "inventory_search" || window.currentTable === "inventory_search") {
            container.innerHTML = `
                <div class="command-bar" style="justify-content: center;">
                    <button class="action-btn" onclick="renderSearchControls()" style="background:#2980b9;">🔍 Run New Search</button>
                    <button class="action-btn" onclick="window.loadTableData('Location')">← Back to Location Map</button>
                </div>`;
            return;
        }

        if (!sheetConfig) return;
        const isDashboard = normalizedName.includes("dashboard");
        const hasManualField = sheetConfig.columns.some(col => ["Quantity", "Current Stock", "Stock_Count"].includes(col.header));
        let buttons = "";

        if (normalizedName === "location") {
            buttons = `
                <button class="action-btn" onclick="UI.renderLocationInspectorControls()" style="background:#8e44ad; font-weight:bold;">📊 Inspect on Location_ID</button>
                <button class="action-btn" onclick="renderSearchControls()" style="background:#2980b9; font-weight:bold;">🔍 Search Inventory</button>
                <button class="action-btn" onclick="UI.manageLocationMap()">Manage Shop Location Map</button>
                <button class="action-btn" onclick="runLocationAudit()">Audit of TBD Locations</button>
                <button class="action-btn" id="btn-print">Print Location Map</button>
            `;
        } else if (!isDashboard) {
            // 🌟 REGULATORY GATEWAY ENGAGED: HIDE '+' BUTTON ONLY IF SHEET IS AN INVENTORY TABLE
            buttons = `
                <button class="action-btn" id="btn-print">🖨️ Print Sheet</button>
                ${hasManualField ? `<button class="action-btn" id="btn-manual-print">📋 Print Manual Log</button>` : ''}
                ${!sheetConfig.isInventory ? `<button class="action-btn" id="btn-add">➕ Add Item</button>` : ''} 
                <button class="action-btn" id="btn-edit">✏️ Edit Table</button>
                ${hasManualField ? `<button class="action-btn" id="btn-inventory-update" style="background:#e67e22;">⚡ Quick Update</button>` : ''}
                ${normalizedName === "resell_inventory" ? `<button class="action-btn" id="btn-resell-status-pivot" style="background:#8e44ad;">📊 Sort By Status</button>` : ''}
            `;
        } else {
            buttons = `
                <button class="action-btn" onclick="UI.renderCentralRegistrationWizard()" style="background:#e67e22; font-weight:bold;">⚡ Central Item Registration</button>
                <button class="action-btn" onclick="UI.renderTagMaintenanceWizard()" style="background:#8e44ad; font-weight:bold;">🔧 Manage Lost/Damaged Tags</button>
                <button class="action-btn" onclick="runUntaggedAudit()" style="background:#c0392b; font-weight:bold;">⚠️ Audit Untagged Items</button>
                <button class="action-btn" id="btn-print">Print Dashboard</button>
            `;
        }
        container.innerHTML = `<div class="command-bar">${buttons}</div>`;
    },

    // 2. STAGE ONE WIZARD: TOKEN IDENTIFICATION GATE
    renderCentralRegistrationWizard() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        title.innerText = "Administrative: Centralized Item Intake Portal";
        
        // Lock the router configuration state flag
        window.currentTable = "inventory_registration";

        let html = `
            <div class="form-card" style="border-left: 6px solid var(--accent); background:#fff; padding: 25px; margin-bottom: 25px;">
                <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">⚡ Central Asset Registration Wizard</h4>
                <p style="font-size:0.85rem; color:#666; margin:0 0 15px 0;">STAGE 1: Token Identification Gate. Select your target table, then choose to register an UNTAGGED bulk item or scan a fresh sticker token.</p>
                
                <div style="display: flex; flex-direction: column; gap: 15px; max-width: 500px; margin-bottom: 20px;">
                    <div style="display: flex; flex-direction: column;">
                        <label style="font-size:0.8rem; font-weight:bold; color:var(--primary); margin-bottom:5px;">Target Inventory Classification Sheet</label>
                        <select id="mae-central-table-selector" class="edit-dropdown" style="height:45px; font-size:0.95rem;">
                            <option value="">-- Choose Target Table --</option>
                            <option value="Shop_Machinery">Shop Machinery</option>
                            <option value="Shop_Power_Tools">Shop Power Tools</option>
                            <option value="Shop_Hand_Tools">Shop Hand Tools</option>
                            <option value="Shop_Consumables">Shop Consumables</option>
                            <option value="Resell_Inventory">Resell Inventory</option>
                        </select>
                    </div>

                    <div style="display: flex; flex-direction: column; position: relative;">
                        <label style="font-size:0.8rem; font-weight:bold; color:var(--primary); margin-bottom:5px;">Scan Fresh Sticker Token (Advanced Tier Focus)</label>
                        <input type="text" id="field-Tag_ID" placeholder="Click here and scan physical label roll..." style="height:45px; border:2px solid var(--border); padding:0 12px; font-weight:bold; font-size:1rem; background: #fffde7;" autofocus>
                        <div id="wizard-tag-feedback" style="margin-top: 5px; font-size: 0.8rem; font-weight: bold;"></div>
                    </div>
                </div>

                <div style="display: flex; gap: 15px;">
                    <button class="action-btn" onclick="UI.processWizardStageOneScan()" style="background:var(--primary); height:45px; font-weight:bold; flex: 1;">⚡ Verify Scanned Tag</button>
                    <button class="action-btn" onclick="UI.processWizardStageOneUntagged()" style="background:#7f8c8d; height:45px; font-weight:bold; flex: 1;">📦 Proceed as UNTAGGED</button>
                </div>
            </div>
            <div id="central-form-render-zone"></div>
        `;
        container.innerHTML = html;
        this.renderCommandBar("");

        // Set real-time listener to handle direct input swipes into the box
        setTimeout(() => {
            const input = document.getElementById("field-Tag_ID");
            if (input) {
                input.focus();
                input.addEventListener("change", () => UI.processWizardStageOneScan());
            }
        }, 100);
    },

//========== END RENDER COMMAND BAR ================

// ================RENDER ENTRY FORM===============
   renderEntryForm(mode, tableName, sheetConfig, onSaveCallback, rowIndex = null, existingData = null, wizardIntakeData=null) {
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

        // If it's a new item, expose Tag_Type so the user can classify the tag
        const forceExposeTagType = (!isEdit && col.header === "Tag_Type");

        if ((col.hidden || col.type === "formula" || col.locked) && !forceExposeTagType) {     
            formHtml += `<input type="hidden" id="${fieldId}" value="${val}">`;
        } 
        else {
            formHtml += `<div class="input-group"><label>${col.header}</label>`;

            // 1A. BOOLEAN
            if (col.type === "boolean") {
                const isChecked = val.toString().toUpperCase() === "TRUE";
                formHtml += `<input type="checkbox" id="${fieldId}" ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
            }
            // 1B. TAG_TYPE DISCIPLINE EXPLICIT ENTRY (Injected for Advanced Tier Intake)
            else if (col.header === "Tag_Type") {
                // MAE PROTECTION: Force-lock type selection to MULTIPLE when working inside the multi-item container panel view
                const isAppendingToContainerView = document.getElementById("current-view-title")?.innerText.includes("Multiple Items");
                let currentSelection = val;

                // 🌟 WIZARD INTAKE DATA ALIGNMENT GUARD 🌟
                if (wizardIntakeData !== null && wizardIntakeData.tagType) {
                currentSelection = wizardIntakeData.tagType;
                } else if (!isEdit && isAppendingToContainerView) {
                currentSelection = "MULTIPLE";
                } else if (!currentSelection) {
                currentSelection = "UNIQUE";
                }

                const isFormLocked = (!isEdit && isAppendingToContainerView);

                formHtml += `
                    <select id="${fieldId}" required ${isFormLocked ? 'disabled style="background-color:#eeeeee; color:#888888; border:1px solid var(--border); width: 100%; height: 45px;"' : 'style="border: 2px solid var(--accent); background: #fffde7; font-weight: bold; height: 45px;"'}>
                        <option value="UNIQUE" ${currentSelection === "UNIQUE" ? "selected" : ""}>UNIQUE (One Tag for One Single Machine/Asset)</option>
                        <option value="MULTIPLE" ${currentSelection === "MULTIPLE" ? "selected" : ""}>MULTIPLE (One Tag for a Container/Bin/Drawer Group)</option>
                    </select>
                `;
                
                // Secret Payload Pass: Feeds the background harvester when the dropdown element is disabled
                if (isFormLocked) {
                    formHtml += `<input type="hidden" id="${fieldId}" value="MULTIPLE">`;
                }
            }    
            // 2. LOCATION_ID (Foundation Discipline)
            else if (col.header === "Location_ID") {
                formHtml += `
                    <div class="location-control-group">
                        <select id="${fieldId}" required>
                            ${window.maeLocations.map(loc => 
                                `<option value="${loc}" ${loc === val ? 'selected' : ''}>${loc}</option>`
                            ).join('')}
                        </select>
                        <span class="foundation-alert">FOUNDATION FIELD: Managed via Location Table</span>
                    </div>`;
            }
            // 3. DROPDOWNS (Includes your new Stock_Level column)
            else if (col.type === "dropdown") {
                const availableOptions = col.options || [];
                formHtml += `
                    <select id="${fieldId}" required>
                        <option value="">-- Select ${col.header} --</option>
                        ${availableOptions.map(opt => `<option value="${opt}" ${opt == val ? 'selected' : ''}>${opt}</option>`).join('')}
                    </select>`;
            }
            // 4. NUMBERS (Includes your new Stock_Count column)
            else if (col.type === "number") {
                const isCurrency = col.format && col.format.includes("$");
                formHtml += `<input type="number" step="${isCurrency ? '0.01' : '1'}" id="${fieldId}" value="${val}" placeholder="${isCurrency ? '0.00' : 'Whole number'}">`;
            } 
            // 5. DATE
            else if (col.type === "date") {
                formHtml += `<input type="date" id="${fieldId}" value="${val}">`;
            } 
            // 6. STANDARD TEXT
            else {
                const isTag = col.header === "Tag_ID";
                
                // PATH A: If editing an existing item, freeze the Tag_ID permanently
                if (isTag && isEdit) {
                    formHtml += `
                        <div style="position: relative;">
                            <input type="text" id="${fieldId}" value="${val}" disabled 
                                   style="background-color: #eeeeee; color: #888888; border: 1px solid var(--border); font-weight: bold; cursor: not-allowed; width: 100%; box-sizing: border-box;">
                            <span style="display: block; font-size: 0.65rem; color: var(--accent); font-weight: bold; margin-top: 4px; text-transform: uppercase;">
                                🔒 PERMANENT MATRIX ANCHOR: Locked Against Editing
                            </span>
                        </div>`;
                } 
                // PATH B: If registering a fresh item via scanner injection, check for a pending mailbox value
                else if (isTag && window.pendingScanValue) {
                    formHtml += `
                        <div style="position: relative;">
                            <input type="text" id="${fieldId}" value="${window.pendingScanValue}" readonly 
                                   style="background-color: #e8f8f5; color: #27ae60; border: 2px solid #27ae60; font-weight: bold; cursor: not-allowed; width: 100%; box-sizing: border-box;">
                            <span style="display: block; font-size: 0.65rem; color: #27ae60; font-weight: bold; margin-top: 4px; text-transform: uppercase;">
                                🔒 SCANNED HARDWARE TOKEN: Locked Against Manual Typing
                            </span>
                        </div>`;
                }
                // PATH C: Fallback to open text field for manual entry (Base Tier / Manual Clicks)
                else {
                    const isTag = col.header === "Tag_ID";
  
                    // 🌟 DETERMINISTIC INTRA-WIZARD INJECTION PASS 🌟
                    if (isTag && wizardIntakeData !== null && wizardIntakeData.tagId) {
                        formHtml += `
                        <div style="position: relative;">
                            <input type="text" id="${fieldId}" value="${wizardIntakeData.tagId}" readonly style="background-color: #e8f8f5; color: #27ae60; border: 2px solid #27ae60; font-weight: bold; cursor: not-allowed; width: 100%; box-sizing: border-box;">
                            <span style="display: block; font-size: 0.65rem; color: #27ae60; font-weight: bold; margin-top: 4px; text-transform: uppercase;"> 🔒 WIZARD INTAKE TOKEN: Locked Against Manual Typing </span>
                        </div>`;
                    } else {
                        // Standard baseline text box generation remains un-impacted
                        formHtml += `<input type="text" id="${fieldId}" value="${val}" placeholder="Enter ${col.header}..." ${isTag ? 'autofocus' : ''}>`;
                    }
                }
            }
            formHtml += `</div>`; // Closes the <div class="input-group"> row container
        }
    });

    formHtml += `</div>
        <div class="form-actions">
            <button class="save-btn" id="submit-form-btn">${isEdit ? 'Update' : 'Save'} to OneDrive</button>
            <button class="cancel-btn" onclick="document.getElementById('entry-form').remove()">Cancel</button>
        </div>
    </div>`;

    container.insertAdjacentHTML('beforebegin', formHtml);

    // RUGGED SCANNER INJECTION: Auto-fills Tag_ID if a scan is pending
    if (window.pendingScanValue) {
        // MAE ENGINE REPAIR: Explicitly isolate unhidden tracking target field
        const tagInput = document.getElementById("field-Tag_ID");
        if (tagInput) {
            tagInput.value = window.pendingScanValue;
            tagInput.style.backgroundColor = "#fffde7"; 
            tagInput.style.border = "2px solid var(--accent)";
            console.log("MAE System Focus Control: Scan successfully bound to field-Tag_ID layout cell.");
        } else {
            console.warn("MAE Intake Exception: Hardware burst intercepted, but this sheet does not contain a visible Tag_ID column layout anchor.");
        }
        window.pendingScanValue = null; 
    }

    //====== SILO CONTROLLER for Methodology Enforcement in Shop_Consumables 
    if (tableName === "Shop_Consumables") {
        const levelSelect = document.getElementById('field-Stock_Level');
        
        // Methodology Silo Mapping
        const unitSiloIds = ['field-UnitCost', 'field-Stock_Count', 'field-ReorderPoint'];
        const bulkSiloIds = ['field-Bulk_Value'];

        const applySiloGovernance = () => {
            const method = levelSelect.value; // "None", "Few", "Adequate", "Many", "Counted"
            
            // 1. Reset all fields to 'Active' first
            const allSiloIds = [...unitSiloIds, ...bulkSiloIds];
            allSiloIds.forEach(id => {
                const el = document.getElementById(id);
                if (!el) return;
                el.disabled = false;
                el.parentElement.classList.remove('silo-locked', 'silo-active');
            });

            // 2. ENFORCE SILOS
            if (method === "Counted") {
                // SILO A: Unit-Based is Active
                bulkSiloIds.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) {
                        el.disabled = true;
                        el.parentElement.classList.add('silo-locked');
                        el.value = ""; // Methodology Wipe
                    }
                });
                unitSiloIds.forEach(id => document.getElementById(id)?.parentElement.classList.add('silo-active'));
            } 
            else if (["Few", "Adequate", "Many"].includes(method)) {
                // SILO B: Bulk-Based is Active
                unitSiloIds.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) {
                        el.disabled = true;
                        el.parentElement.classList.add('silo-locked');
                        el.value = ""; // Methodology Wipe
                    }
                });
                bulkSiloIds.forEach(id => document.getElementById(id)?.parentElement.classList.add('silo-active'));
            }
            else {
                // "None" or Empty: Lock both to prevent accidental data entry
                allSiloIds.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) {
                        el.disabled = true;
                        el.parentElement.classList.add('silo-locked');
                    }
                });
            }
        };

        // Initialize state on form open and listen for changes
        levelSelect.addEventListener('change', applySiloGovernance);
        applySiloGovernance();
    }

    // 7. SAVE HANDLER
    const submitBtn = document.getElementById("submit-form-btn");
    if (submitBtn) {
        submitBtn.onclick = async () => {
            submitBtn.disabled = true;
            submitBtn.innerText = "Syncing with Ledger...";
            
            await onSaveCallback(rowIndex, existingData); 

            if (document.getElementById("entry-form")) {
                document.getElementById("entry-form").remove();
            }
        };
    }
},
//======= END RENDER ENTRY FORM ============


//========== EXIT EDIT MODE ==============
    exitEditMode(forceRefresh = false) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    // 1. Reset Global Flags
    window.isEditing = false; // TURNS OFF THE LOCK

     // If we are discarding, don't even look at the inputs, just reload from OneDrive
    if (forceRefresh === true) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    
    // 1. Force a "Nuclear" UI wipe
    container.innerHTML = `<div class="loader">Restoring Data from OneDrive...</div>`;
    title.innerText = "Reverting Changes...";

    // 2. Reset the command bar
    this.renderCommandBar(window.currentTable); 
    
    // 3. Trigger reload
    window.loadTableData(window.currentTable);
    return; 
}

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
        const select = cell.querySelector('select');
        if (input) {
            //==== if checkbox, don't show "on"
            if (input.type === "checkbox"){
                const isChecked = input.checked;
                // restore visual disabled checkbox immediately
                cell.innerHTML = `<input type="checkbox" disabled ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
            } else {
                cell.innerText = input.value; // Capture the final number
            }
        }

        // --- B. Handle Dropdowns ---
        //const select = cell.querySelector('select');
        else if (select) {
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

    this.renderCommandBar(window.currentTable);

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
    const table = document.getElementById("main-data-table");
    if (!table) return;

    // 1. RUGGED WHITELIST: Case-insensitive fragments that match your config.js headers
    const keepHeaders = ["location_id", "machine name", "tool name", "item name", "stock_count"];

    // 2. DOM-BASED COLUMN SCANNER
    // We scan the actual rendered headers to avoid "index drift" from hidden columns
    const headers = table.querySelectorAll("thead th");
    headers.forEach((th, pos) => {
        const headerText = th.innerText.trim().toLowerCase();
        const isActionCol = th.classList.contains("edit-only-cell");
        
        // Logic: Should we keep this column?
        const isKeep = keepHeaders.some(h => headerText.includes(h));

        // If it's not a keep-header and not the action column, Nuke it
        if (!isKeep && !isActionCol && headerText !== "") {
            // Apply the V1.9 High-Specificity Class
            th.classList.add("print-force-hide");
            
            // Hide every cell in this specific 1-based column position
            // nth-child is 1-based, so pos + 1 is the exact match
            table.querySelectorAll(`tbody td:nth-child(${pos + 1})`).forEach(td => {
                td.classList.add("print-force-hide");
            });
        }
    });

    // 3. UI PREPARATION
    const container = document.getElementById("app-content");
    const printHeader = document.createElement("div");
    printHeader.className = "print-only-title";
    printHeader.innerHTML = `<h1>MANUAL INVENTORY LOG: ${sheetConfig.tabName}</h1>`;
    
    // Inject the header and the special mode class
    container.prepend(printHeader);
    table.classList.add("manual-log-mode");
    
    // 4. TRIGGER PRINT WITH RUGGED SETTLE TIME
    // 450ms ensures the browser "paints" the hidden classes into the print buffer
    setTimeout(() => {
        window.print();

        // 5. RUGGED CLEANUP: Restore the UI to its digital state
        printHeader.remove();
        table.classList.remove("manual-log-mode");
        table.querySelectorAll(".print-force-hide").forEach(el => {
            el.classList.remove("print-force-hide");
        });
        console.log("MAE System: Print Manual Log Complete. UI Restored.");
    }, 450);
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
        <div style="display: flex; flex-direction: column; gap: 12px;">
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
            <!-- MAE REPAIR MODULE WORKFLOW CARRIER -->
            <button class="action-btn" 
                    style="width: 100%; height: 50px; font-size: 1.1rem; background: #c0392b; font-weight: bold;" 
                    onclick="window.initiateTagReplacementWorkflow('${tableName}', ${rowIndex}, '${rowData[sheetConfig.columns.findIndex(c => c.header === "Tag_ID")]}')">
                ⚠️ Replace Damaged / Missing Physical Tag
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
},
//========= END SCAN RESULT UI LOGIC =============

//===== PROMPT LOGIC for Location_ID =====
async promptNewLocation() {
    const newLoc = prompt("ESTABLISH NEW SHOP LOCATION:\nThis permanently adds a new storage spot to your Control Tower.");
    
    if (newLoc && newLoc.trim() !== "") {
        const cleanLoc = newLoc.trim().toUpperCase();
        
        // VISUAL FEEDBACK: Let the user know the network request is starting
        console.log(`MAE System: Registering new location [${cleanLoc}]...`);
        
        const success = await window.submitNewLocationToTable(cleanLoc);
        
        if (success) {
            // 1. Sync the local list of locations
            await window.refreshLocationCache();
            
            // 2. Alert the user (Essential for shop-floor confirmation)
            alert(`Location ${cleanLoc} successfully registered.`);
            
            // 3. Refresh the UI so the new location appears in dropdowns immediately
            window.loadTableData(window.currentTable);
        } else {
            alert("Error: Could not save location. Check your internet connection.");
        }
    }
},
//====== END PROMPT LOGIC For Location_ID ====

//=============================================
//========== Manage Location Map ===========
//==================================================
    async applyLocationChange(rowIndex, oldId) {
        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        const rowDataMap = {};

        // 1. DYNAMIC DATA GATHERING
        // We look for inputs by a predictable ID pattern based on the Header Name
        sheetConfig.columns.forEach(col => {
            if (col.type !== 'formula' && col.header !== 'mae_id') {
                const inputId = `loc-${col.header.replace(/\s+/g, '')}-${rowIndex}`;
                const input = document.getElementById(inputId);
                if (input) {
                    rowDataMap[col.header] = col.header === "Location_ID" ? input.value.trim().toUpperCase() : input.value.trim();
                }
            }
        });

        const newId = rowDataMap["Location_ID"];
        const confirmed = confirm(`Update Foundation Point [${oldId}]? \n\nNote: If the ID changed, existing items in other tables will need to be re-assigned.`);
        
        if (confirmed) {
            this.showLoading(`Updating ${newId}...`);
            const success = await window.updateLocationRecord(rowIndex, rowDataMap);
            if (success) {
                await window.refreshLocationCache();
                this.manageLocationMap();
            }
        }
    },

    manageLocationMap() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        title.innerText = "Administrative: Manage Shop Location Map";

        this.showLoading("Fetching Foundation Data...");
    
        window.Dashboard.getFullTableData("Location").then(data => {
            const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Location");
            if (!sheetConfig) { this.showError("Configuration for 'Location' table not found."); return; }

            let html = `<div class="form-card" id="location-manager">
                <div style="border-bottom: 2px solid var(--accent); margin-bottom: 15px; padding-bottom: 10px;">
                    <h4>+ Establish New Foundation Point</h4>
                    <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">`;

            // Build input layouts dynamically based on system schema configuration parameters
            sheetConfig.columns.forEach(col => {
                if (col.hidden || col.header === "mae_id") return;
                const inputId = `new-loc-${col.header.replace(/\s+/g, '')}`;
            
                html += `
                    <div style="flex: 1; min-width: 150px; display: flex; flex-direction: column;">
                        <input type="text" id="${inputId}" placeholder="${col.header}">
                        ${col.header === "Location_ID" ? `<div id="new-loc-id-feedback" style="margin-top:4px; min-height:1rem;"></div>` : ''}
                    </div>`;
            });

            html += `<button class="action-btn" onclick="UI.saveNewLocation()" style="background:#27ae60; margin-bottom:20px;">Establish</button>
                </div>
            </div>
            <div id="location-list-scroll" style="max-height: 500px; overflow-y: auto;">
                <table class="inventory-table">
                    <thead><tr>`;

            sheetConfig.columns.forEach(col => { if (!col.hidden) html += `<th>${col.header}</th>`; });
            html += `<th>Actions</th></tr></thead><tbody>`;

            // Render row item entries contextually
            data.forEach((row, idx) => {
                const vals = (row.values && Array.isArray(row.values[0])) ? row.values[0] : row.values;
                const locIdIdx = sheetConfig.columns.findIndex(c => c.header === "Location_ID");
                const currentLocName = vals[locIdIdx];
                if (currentLocName === "TBD") return;

                html += `<tr>`;
                sheetConfig.columns.forEach((col, colIdx) => {
                    if (col.hidden) return;
                    const fieldId = `loc-${col.header.replace(/\s+/g, '')}-${idx}`;
                    const displayVal = vals[colIdx] || '';
                
                    html += `
                        <td>
                            <input type="text" id="${fieldId}" value="${displayVal}" style="width:100%;">
                            ${col.header === "Location_ID" ? `<div id="${fieldId}-feedback" style="margin-top:4px; min-height:1rem;"></div>` : ''}
                        </td>`;
                });

                html += `
                    <td>
                        <button class="mini-btn" onclick="UI.applyLocationChange(${idx}, '${currentLocName}')" style="background:#2980b9;">Update</button>
                        <button class="mini-btn" onclick="UI.removeLocation('${currentLocName}')" style="background:#c0392b;">Delete</button>
                    </td>
                </tr>`;
            });

            html += `</tbody></table></div>
                <div class="form-actions" style="margin-top: 20px;">
                     <button class="cancel-btn" 
                        style="background: var(--primary); width: 100%; height: 50px; font-size: 1.1rem;" 
                        onclick="window.UI.exitLocationManagerAndRefresh()">
                        💾 Save Map Configuration & Return to Inventory
                    </button>   
                /div>`;
            
            container.innerHTML = html;

            // 🌟 ATTACH REAL-TIME RE-ORIENTATION LISTENERS AFTER CONTENT IS INJECTED INTO DOM
            const primaryNewInput = document.getElementById("new-loc-Location_ID");
            this.attachLocationValidationGuard(primaryNewInput, "new-loc-id-feedback");

            data.forEach((row, idx) => {
                const rowInput = document.getElementById(`loc-Location_ID-${idx}`);
                if (rowInput) {
                    this.attachLocationValidationGuard(rowInput, `loc-Location_ID-${idx}-feedback`);
             }
            });

        }).catch(err => {
            console.error("MAE Manager Load Crash Intercepted:", err);
            this.showError("Could not load Location data from OneDrive.");
        });
    },

    async removeLocation(locName) {
    const confirmed = confirm(`CRITICAL WARNING: You are about to DELETE the location [${locName}]. \n\nAny items currently assigned to this spot will lose their physical reference. Proceed?`);
    
    if (confirmed) {
        this.showLoading(`Decommissioning ${locName}...`);
        
        try {
            const data = await window.Dashboard.getFullTableData("Location");
            const locConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Location");
            
            // --- HEADER DISCOVERY ---
            const locIdx = locConfig.columns.findIndex(c => c.header === "Location_ID");

            // RUGGED SEARCH: Dig into row.values[0] where Graph API hides the data
            const rowIndex = data.findIndex(row => {
                const rowCells = row.values[0]; 
                return rowCells[locIdx] === locName;
            });

            if (rowIndex !== -1) {
                // This calls the engine in app.js
                const success = await window.deleteExcelRow("Location", rowIndex);
                if (success) {
                    await window.refreshLocationCache();
                    this.manageLocationMap(); // Reload Manager
                }
            } else {
                console.error("MAE System: Could not find row for " + locName);
                this.showError("Search failed: Location not found in OneDrive.");
            }
        } catch (err) {
            console.error("MAE System: Removal failed", err);
            this.showError("Failed to delete location. Check connection.");
        }
    }
},

    async renameLocation(oldName) {
        const newName = prompt(`RENAME FOUNDATION POINT: \nChanging [${oldName}] will break the link for all items currently assigned to it in Excel. \n\nEnter new name:`, oldName);
        
        if (newName && newName.trim() !== "" && newName.trim().toUpperCase() !== oldName) {
            const cleanNewName = newName.trim().toUpperCase();
            this.showLoading(`Renaming ${oldName} to ${cleanNewName}...`);

            try {
                const data = await window.Dashboard.getFullTableData("Location");
                const locConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Location");
                const locIdx = locConfig.columns.findIndex(c => c.header === "Location_ID");
                const rowIndex = data.findIndex(row => row.values[0][locIdx] === oldName);

                if (rowIndex !== -1) {
                    const success = await window.updateLocationName(rowIndex, cleanNewName);
                    if (success) {
                        await window.refreshLocationCache();
                        this.manageLocationMap(); // Reload Manager
                    }
                }
            } catch (err) {
                console.error("MAE System: Rename failed", err);
                this.showError("Failed to rename location in OneDrive.");
            }
        }
    },

    async saveNewLocation() {
        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Location");
        const rowDataMap = {};
        let missingField = false;

        // Guardrail Check: Abort if submission field holds an active validation error layout template state
        const targetKeyInput = document.getElementById("new-loc-Location_ID");
        if (targetKeyInput && targetKeyInput.classList.contains("input-error")) {
            alert("CRITICAL SYSTEM RULES VIOLATION:\n\nYou cannot write this Location_ID to the spreadsheet ledger until naming compliance rules are met.");
            targetKeyInput.focus();
            return;
        }

        sheetConfig.columns.forEach(col => {
            if (col.hidden || col.header === "mae_id") return;
            const inputId = `new-loc-${col.header.replace(/\s+/g, '')}`;
            const input = document.getElementById(inputId);
            if (input) {
                const val = input.value.trim();
                if (col.header === "Location_ID" && !val) { missingField = true; }
                rowDataMap[col.header] = col.header === "Location_ID" ? val.toUpperCase() : val;
            }
        });

        if (missingField) { alert("Error: Location_ID is mandatory."); return; }

        this.showLoading("Writing New Foundation Point...");
        const success = await window.submitNewLocationToTable(rowDataMap);
    
        if (success) {
            await window.refreshLocationCache();
            this.manageLocationMap();
        } else {
            alert("Error: Failed to register location on OneDrive ledger.");
            this.manageLocationMap();
        }
    },

//=======  Core Real-Time Location ID Validation Guard
attachLocationValidationGuard(inputElement, feedbackElementId) {
    if (!inputElement) return;

    inputElement.addEventListener("input", (e) => {
        // 1. SILENT AUTO-CORRECT: Instantly transform spaces to hyphens and lowercase to uppercase
        let cursorPosition = e.target.selectionStart;
        let originalLength = e.target.value.length;
        
        let cleanText = e.target.value.toUpperCase().replace(/\s+/g, "-");
        e.target.value = cleanText;
        
        // Preserve tablet typing cursor position during automated string conversion adjustments
        let lengthDelta = cleanText.length - originalLength;
        e.target.setSelectionRange(cursorPosition + lengthDelta, cursorPosition + lengthDelta);

        // 2. REGEX SPECIFICATION EVALUATION (Enforces leading zeros and dash separation constraints)
        const locationPatternRegex = /^([A-Z0-9]+)(-[A-Z0-9]{2,})+$/;
        const isValid = locationPatternRegex.test(cleanText);
        const feedbackEl = document.getElementById(feedbackElementId);

        if (cleanText === "") {
            inputElement.style.borderColor = "var(--border)";
            inputElement.style.backgroundColor = "#ffffff";
            if (feedbackEl) feedbackEl.innerHTML = "";
            return;
        }

        if (!isValid) {
            // 3. INDUSTRIAL HIGH-AWARENESS FAT-FINGER ALERT TINTS
            inputElement.style.borderColor = "#e74c3c";     // Red Alert Border Outline
            inputElement.style.backgroundColor = "#fadbd8"; // Soft Red Alert Sheet Tint
            inputElement.classList.add("input-error");
            
            if (feedbackEl) {
                feedbackEl.style.color = "#c0392b";
                feedbackEl.style.fontWeight = "bold";
                feedbackEl.style.fontSize = "0.85rem";
                feedbackEl.innerHTML = `⚠️ Format Variance: Use padded double-digits and hyphens. (e.g., SHOP-RACK-02-BIN-04)`;
            }
        } else {
            // 4. SOLID DISCIPLINARY STRUCTURAL STATE CONFIRMED
            inputElement.style.borderColor = "#27ae60";     // Operational Green Border
            inputElement.style.backgroundColor = "#e8f8f5"; // Operational Mint Sheet Tint
            inputElement.classList.remove("input-error");
            
            if (feedbackEl) {
                feedbackEl.style.color = "#27ae60";
                feedbackEl.style.fontWeight = "bold";
                feedbackEl.style.fontSize = "0.85rem";
                feedbackEl.innerHTML = `✅ Structure Verified: Solid physical map anchor matrix alignment.`;
            }
        }
    });
},

//====== Secure COnfiguration Exit and Sync Router
async exitLocationManagerAndRefresh() {
    this.showLoading("Synchronizing physical workshop map parameters...");
    
    try {
        // 1. Force a raw background cache sweep to ensure all memory indexes align
        await window.refreshLocationCache();
        
        // 2. Identify what view context was open before entering administrative mode
        // If the user was editing a table, default back to the Location map to view coordinates
        const targetView = (window.currentTable && window.currentViewTitle !== "location_audit") 
            ? window.currentTable 
            : "Location";

        console.log(`MAE Engine: Administrative map locked. Re-drawing sheet layout workspace: [${targetView}]`);
        
        // 3. Clear active input visual variables completely
        this.exitEditMode();
        
        // 4. Force a clean, re-sorted download straight from your OneDrive backend ledger
        window.loadTableData(targetView);

    } catch (err) {
        console.error("MAE Engine Sync Failure during manager teardown:", err);
        window.loadTableData("Master_Dashboard");
    }
},
//====================================================
// ======== END Manage location Map Logic ========
//==================================================


//======= "Virtual table renderer" for TBD ======
renderAuditGrid(auditData) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    
    title.innerText = "Audit: Items Awaiting Final Location Assignment";

    // 1. Check for empty state (Audit Complete)
    if (!auditData || auditData.length === 0) {
        container.innerHTML = `
            <div class="form-card" style="text-align:center; padding:40px;">
                <h3 style="color:#27ae60;">✅ Audit Complete</h3>
                <p>All items have been assigned a location in the Control Tower.</p>
                <button class="action-btn" onclick="loadTableData('Location')">Return to Location Map</button>
            </div>`;
        return;
    }

    // 2. Group the data by category for visual organization
    const grouped = auditData.reduce((acc, item) => {
        if (!acc[item.category]) acc[item.category] = [];
        acc[item.category].push(item);
        return acc;
    }, {});

    // 3. Build the two-column grouped table
    let html = `<table class="inventory-table" id="main-data-table">`;
    
    for (const [category, items] of Object.entries(grouped)) {
        // Render Category Header and Sub-Headers (Item / Location_ID)
        html += `
            <thead>
                <tr>
                    <th colspan="2" style="background:var(--primary); color:white; padding:12px;">
                        ${category.toUpperCase()}
                    </th>
                </tr>
                <tr>
                    <th style="width:60%;">Item</th>
                    <th style="width:40%;">Location_ID</th>
                </tr>
            </thead>
            <tbody>`;
        
        items.forEach(item => {
            // Option A: Live Cleanup - includes the unique ID for the fade-out animation
            html += `
            <tr id="audit-row-${item.mae_id}" style="transition: opacity 0.5s ease, transform 0.5s ease;">
                <td class="locked-cell">${item.itemName}</td>
                <td>
                    <select class="edit-dropdown" 
                            style="width:100%; height:40px; background:#fffde7; border:1px solid var(--accent);"
                            onchange="handleAuditUpdate('${item.tableName}', '${item.mae_id}', this.value, 'audit-row-${item.mae_id}')">
                        <option value="TBD">TBD</option>
                        ${window.maeLocations
                            .filter(loc => loc !== "TBD")
                            .map(loc => `<option value="${loc}">${loc}</option>`)
                            .join('')}
                    </select>
                </td>
            </tr>`;
        });
        html += `</tbody>`;
    }
    
    html += `</table>`;
    container.innerHTML = html;

    // Ensure we show the correct command bar context
    this.renderCommandBar("location_audit"); 
},
//======= END   "Virtual table renderer" for TBD ======

// =======Virtual TBD Item Print =========
printVirtualAudit(auditData, title) {
    const printContainer = document.createElement("div");
    printContainer.className = "print-only-container";
    printContainer.id = "temp-print-zone";

    const grouped = auditData.reduce((acc, item) => {
        if (!acc[item.category]) acc[item.category] = [];
        acc[item.category].push(item);
        return acc;
    }, {});

    let html = `<div class="print-only-title"><h1>${title}</h1></div>`;
    // We use your 'manual-log-mode' class to ensure rugged table borders and row height
    html += `<table class="inventory-table manual-log-mode" style="width:100%; border-collapse:collapse;">`;
    
    for (const [category, items] of Object.entries(grouped)) {
        html += `<thead>
                    <tr><th colspan="2" style="background:#eee; border:1px solid black; padding:10px;">${category.toUpperCase()}</th></tr>
                    <tr><th style="border:1px solid black; width:60%;">Item</th><th style="border:1px solid black; width:40%;">Location Assignment</th></tr>
                 </thead><tbody>`;
        items.forEach(item => {
            html += `<tr><td style="border:1px solid black; height:40px; padding:5px;">${item.itemName}</td><td style="border:1px solid black;"></td></tr>`;
        });
    }
    html += `</tbody></table>`;

    printContainer.innerHTML = html;
    document.body.appendChild(printContainer);

    // Small delay to let the DOM settle before opening the print dialog
    setTimeout(() => {
        window.print();
        printContainer.remove();
    }, 50);

    window.print();
    printContainer.remove();
},
//======= END    Virtual TBD Item Print =========

//========== Re-Homing Table for deletion of Location_ID ====
async removeLocation(locName) {
    const container = document.getElementById("table-container");
    this.showLoading(`Scanning Workshop for items in ${locName}...`);
    
    const deps = await window.getLocationDependencies(locName);

    if (deps.length > 0) {
        // SCENARIO A: Dependencies Found - Enforce Re-homing
        let html = `
            <div class="form-card" style="border-left: 10px solid var(--accent);">
                <h3>⚠️ Warning: Location Dependency Detected</h3>
                <p>There are <b>${deps.length}</b> items assigned to <b>${locName}</b>. Please assign them to a new Location_ID now.</p>
                <p style="font-size: 0.9rem; color: #7f8c8d;">Items not updated will revert to <b>TBD</b> for future assignment via the Audit tool.</p>
                
                <table class="inventory-table">
                    <thead>
                        <tr><th>Item Name</th><th>New Location Assignment</th></tr>
                    </thead>
                    <tbody>`;
        
        deps.forEach((item, idx) => {
            // RUGGED FIX: Define cleanMaeId INSIDE the loop so it maps to the current 'item'
            // We extract the first element if it's an array, otherwise use as-is
            const cleanMaeId = Array.isArray(item.mae_id) ? item.mae_id[0] : item.mae_id;

            html += `
                <tr>
                    <td>${item.itemName}</td>
                    <td>
                        <select class="rehome-select" data-table="${item.tableName}" data-id="${cleanMaeId}">
                            <option value="TBD">TBD (Unassigned)</option>
                            ${window.maeLocations.filter(l => l !== 'TBD' && l !== locName).map(l => `<option value="${l}">${l}</option>`).join('')}
                        </select>
                    </td>
                </tr>`;
        });

        html += `</tbody></table>
                <div style="margin-top: 20px; display: flex; gap: 10px;">
                    <button class="save-btn" onclick="UI.finalizeDecommission('${locName}')">Confirm Location Deletion</button>
                    <button class="cancel-btn" onclick="UI.manageLocationMap()">Cancel</button>
                </div>
            </div>`;
        container.innerHTML = html;
    } else {
        // SCENARIO B: No Dependencies - The "Discipline" Confirm
        container.innerHTML = `
            <div class="form-card" style="text-align: center; padding: 40px;">
                <h3>Confirm System Change</h3>
                <p>You are deleting <b>${locName}</b> from your Location Map.</p>
                <p>This structural change is permanent. Click below to proceed.</p>
                <div style="margin-top: 20px;">
                    <button class="save-btn" style="background:#c0392b;" onclick="UI.finalizeDecommission('${locName}')">Confirm Location Deletion</button>
                    <button class="cancel-btn" onclick="UI.manageLocationMap()">Cancel</button>
                </div>
            </div>`;
    }
},

//========== END Re-Homing Table for deletion of Location_ID ====

//====== Finalize Decommission: ensures location_id default to TBD if not selected ====
async finalizeDecommission(locName) {
    // 🌟 1. HARVEST DATA FIRST: Grab the user's selections while they are still on the screen
    const selects = document.querySelectorAll('.rehome-select');
    
    // Create an in-memory array of the changes to be made
    const changesToProcess = [];
    selects.forEach(select => {
        changesToProcess.push({
            newLoc: select.value,
            tableName: select.getAttribute('data-table'),
            maeId: select.getAttribute('data-id')
        });
    });

    console.log(`MAE System: Found ${changesToProcess.length} items to re-home.`);

    // 2. SYSTEM LOCK: Prevent any other syncs from firing during re-homing
    if (window.globalClickOffHandler) {
        document.removeEventListener('mousedown', window.globalClickOffHandler);
        console.log("MAE System: UI Locked for Decommissioning.");
    }
    
    // Now it is perfectly safe to clear the table container and show loading
    this.showLoading(`LOCKING SYSTEM: Re-homing items from ${locName}...`);

    if (changesToProcess.length === 0) {
        console.warn("MAE System: No items found to re-home. Proceeding to direct deletion.");
    }
    
    // 3. SEQUENTIAL SYNC: Loop and WAIT through our memory array, NOT the deleted DOM dropdowns
    for (const item of changesToProcess) {
        console.log(`MAE System: Re-homing ${item.maeId} to ${item.newLoc} in ${item.tableName}`);
        
        // Silent update (no row removal from UI)
        await window.handleAuditUpdate(item.tableName, item.maeId, item.newLoc, null);
        
        // Throttling: 400ms pause to respect Microsoft Graph API limits
        await new Promise(r => setTimeout(r, 400));
    }

    // 4. FINAL DELETION: Only happens after the re-homing loop finishes
    const data = await window.Dashboard.getFullTableData("Location");
    const locConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Location");
    const locIdx = locConfig.columns.findIndex(c => c.header === "Location_ID");
    
    // RUGGED LOOKUP: Find the specific row index for the Location being deleted
    const rowIndex = data.findIndex(row => {
        const rowCells = (row.values && Array.isArray(row.values[0])) ? row.values[0] : row.values;
        return rowCells[locIdx] === locName;
    });

    if (rowIndex !== -1) {
        const success = await window.deleteExcelRow("Location", rowIndex);
        if (success) {
            // 5. REFRESH EVERYTHING: Wipe cache and refresh local data
            await window.refreshLocationCache();
            alert("System Integrity Verified: Items re-homed and Location removed.");
            
            // 6. SYSTEM UNLOCK: Re-attach the click-off listener
            if (window.globalClickOffHandler) {
                document.addEventListener('mousedown', window.globalClickOffHandler);
            }
            
            this.manageLocationMap(); // Reload the manager UI
        }
    } else {
        // RUGGED RECOVERY: If location deletion fails, unlock the system so user can retry
        if (window.globalClickOffHandler) {
            document.addEventListener('mousedown', window.globalClickOffHandler);
        }
        this.showError("Location record not found for final deletion.");
    }
},
//====== END   Finalize Decommission: ensures location_id default to TBD if not selected ====

  // ====== UNIFIED WORKSPACE SEARCH HUB FOR MULTIPLE ITEMS ON A SINGLE TAG ======
  renderVirtualSearchHub(auditData, scannedTagId = null, activeTableName = null, itemCategory = "By_Location") {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    title.innerText = "Search Results: Multiple Items on Scanned Tag";

    // Group the data by category for visual organization on the shop floor
    const grouped = auditData.reduce((acc, item) => {
      if (!acc[item.category]) acc[item.category] = [];
      acc[item.category].push(item);
      return acc;
    }, {});

    let html = "";
    if (scannedTagId && activeTableName) {
      const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === activeTableName);
      const classificationLabel = sheetConfig ? sheetConfig.tabName : "this classification";
      
      // --- 🌟 NEW ARCHITECTURE: CONTEXT-AWARE INDUSTRIAL STATUS BANNERS 🌟 ---
      let bannerTitleText = `CONTAINER ACTIVE: [Tag ID: ${scannedTagId}]`;
      let bannerSubtitleText = `Quickly append a new tool or component record into this physical storage space.`;
      
      if (itemCategory === "By_Topic") {
        bannerTitleText = `THEMATIC TOPIC ACTIVE: [Tag ID: ${scannedTagId}]`;
        bannerSubtitleText = `Quickly append a new distributed part record into this virtual thematic list tracker.`;
      }

      // 🌟 THE METADATA BRIDGE PASS: Injects 'itemCategory' safely as an escaped string literal payload 🌟
      html += `
        <div class="form-card" style="border-left: 6px solid var(--accent); background: #fff; padding: 20px; margin-bottom: 25px; display: flex; align-items: center; justify-content: space-between; gap: 20px;">
          <div>
            <h4 style="margin:0 0 5px 0; color:var(--primary); text-transform:uppercase;">${bannerTitleText}</h4>
            <p style="margin:0; font-size:0.85rem; color:#666;">${bannerSubtitleText}</p>
          </div>
          <button class="action-btn" style="background: #e67e22; font-weight: bold; padding: 12px 25px; font-size: 1rem; flex-shrink: 0;" 
                  onclick="window.pendingScanValue='${scannedTagId}'; window.maeWizardActiveCategory='${itemCategory}'; window.handleAddClick('${activeTableName}')">
            ➕ Add Item to this Container
          </button>
        </div>
      `;
    }

    html += `<table class="inventory-table" id="main-data-table">`;
    for (const [category, items] of Object.entries(grouped)) {
      html += `
        <thead>
          <tr>
            <th colspan="2" style="background:var(--primary); color:white; padding:12px;">
              ${category.toUpperCase()} (${items.length} Items)
            </th>
          </tr>
          <tr>
            <th style="width:70%;">Item Description</th>
            <th style="width:30%;">Action</th>
          </tr>
        </thead>
        <tbody>`;

      items.forEach(item => {
        html += `
          <tr>
            <td class="locked-cell">${item.itemName}</td>
            <td style="text-align:center;">
              <button class="action-btn" style="padding: 5px 12px; font-size: 0.85rem; background: var(--accent);" onclick="loadTableData('${item.tableName}')">
                ✏%EF%B8%8F View / Edit in Table
              </button>
            </td>
          </tr>`;
      });
      html += `</tbody>`;
    }
    html += `</table>`;
    html += `
      <div class="form-actions" style="margin-top: 20px; text-align: center;">
        <button class="cancel-btn" style="width:50%;" onclick="loadTableData('Master_Dashboard')">
          ← Back to Dashboard
        </button>
      </div>`;

    container.innerHTML = html;

    const actionZone = document.getElementById("action-bar-zone");
    if (actionZone) {
      actionZone.innerHTML = `
        <div class="command-bar" style="justify-content: center;">
          <button class="action-btn" onclick="loadTableData('Master_Dashboard')">← Return to Dashboard</button>
        </div>`;
    }
  },
//====== END Virtual Search Hub Generator

//========== Resell Status Pivot Viewport Modulator
renderStatusPivotControls() {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    const config = window.maeSystemConfig;
    
    const sheetConfig = config.worksheets.find(s => s.tableName === "Resell_Inventory");
    const statusCol = sheetConfig.columns.find(c => c.header === "Current Status");
    const statusOptions = statusCol ? (statusCol.options || []) : [];

    title.innerText = "Resell Inventory: Filter View By Status Selection";
    window.currentTable = "resell_status_pivot"; // Establish virtual routing state context

    let html = `
        <div class="form-card" style="border-left:5px solid #8e44ad; padding:25px; background:#fff; margin-bottom:20px;">
            <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">Select Current Status Target</h4>
            <div style="display:flex; gap:15px; flex-wrap:wrap; align-items:center;">
                <select id="mae-resell-status-selector" class="edit-dropdown" style="flex:1; max-width:400px; height:50px;">
                    <option value="">-- Choose Status Group --</option>
                    ${statusOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                </select>
                <button class="action-btn" style="background:#8e44ad; height:50px; font-size:1rem;" onclick="UI.executeResellStatusFilter()">📊 Generate Filtered Table</button>
            </div>
        </div>
        <div id="status-filtered-table-mount"></div>
    `;

    container.innerHTML = html;

    // Load custom, context-sensitive back-to-base routing action options row
    const actionZone = document.getElementById("action-bar-zone");
    if (actionZone) {
        actionZone.innerHTML = `
            <div class="command-bar">
                <button class="action-btn" onclick="window.loadTableData('Resell_Inventory')">← Return to Resell Inventory</button>
            </div>`;
    }
},
//=========  END: Resell Status Pivot Viewport Modulator

//=========  Compiling and Printing the Filtered Subset (Resell Inventory)
async executeResellStatusFilter() {
    const selector = document.getElementById("mae-resell-status-selector");
    const selectedStatus = selector ? selector.value : "";
    const tableMount = document.getElementById("status-filtered-table-mount");
    const actionZone = document.getElementById("action-bar-zone");

    if (!selectedStatus) { alert("Please select an operational category status parameter first."); return; }

    this.showLoading("Isolating records from local data cache buffer...");

    try {
        // Pull fresh dataset straight from your decoupled ledger using your dashboard core tool
        const rawRows = await window.Dashboard.getFullTableData("Resell_Inventory");
        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Resell_Inventory");
        const statusIdx = sheetConfig.columns.findIndex(c => c.header === "Current Status");

        // 🌟 Step 1: Filter the rows based on the dropdown selection (With Nested Array Fix)
        const matchingRows = rawRows.filter(rowObj => {
            // Dig down into the inner array array wrapper just like app.js does
            const cells = (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values[0] : rowObj.values;
            return cells && String(cells[statusIdx]).trim() === selectedStatus;
        });

        // Save this specific subset to global memory so the print subsystem can see it seamlessly
        this.activeStatusPivotData = matchingRows;
        this.activeStatusPivotLabel = selectedStatus;

        if (matchingRows.length === 0) {
            tableMount.innerHTML = `<p style="padding:20px; text-align:center; font-style:italic; border:1px dashed #ccc; background:#f9f9f9; color:#333;">Zero records currently match the status criteria: "${selectedStatus}".</p>`;
            
            // Clean out the loading screen wrapper to allow user navigation to recover safely
            const container = document.getElementById("table-container");
            if (container) {
                // Preserves the selection form card structure on screen
                container.innerHTML = `
                    <div class="form-card" style="border-left:5px solid #8e44ad; padding:25px; background:#fff; margin-bottom:20px;">
                        <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">Select Current Status Target</h4>
                        <div style="display:flex; gap:15px; flex-wrap:wrap; align-items:center;">
                            <select id="mae-resell-status-selector" class="edit-dropdown" style="flex:1; max-width:400px; height:50px;">
                                <option value="">-- Choose Status Group --</option>
                                ${sheetConfig.columns.find(c => c.header === "Current Status").options.map(opt => `<option value="${opt}" ${opt === selectedStatus ? 'selected' : ''}>${opt}</option>`).join('')}
                            </select>
                            <button class="action-btn" style="background:#8e44ad; height:50px; font-size:1rem;" onclick="UI.executeResellStatusFilter()">📊 Generate Filtered Table</button>
                        </div>
                    </div>
                    <div id="status-filtered-table-mount"><p style="padding:20px; text-align:center; font-style:italic; border:1px dashed #ccc; background:#f9f9f9; color:#333;">Zero records currently match the status criteria: "${selectedStatus}".</p></div>
                `;
            }
            this.renderCommandBar(""); 
            return;
        }

        // 🌟 Step 2: Normalize dates and data elements (With Nested Array Fix)
        const flattenedSubset = matchingRows.map(rowObj => {
            const cells = (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values[0] : rowObj.values;
            const cleanValues = cells.map((val, idx) => {
                const colDef = sheetConfig.columns[idx];
                return (colDef && colDef.type === 'date') ? window.excelSerialToDate(val) : val;
            });
            return { ...rowObj, values: cleanValues };
        });

        // Step 3: Run your Default Location_ID sorting engine over the filtered rows
        const locIdx = sheetConfig.columns.findIndex(c => c.header === "Location_ID");
        if (locIdx !== -1) {
            flattenedSubset.sort((a, b) => {
                const locA = String(a.values[locIdx] || "").trim();
                const locB = String(b.values[locIdx] || "").trim();
                if (locA === locB) return 0;
                if (locA === "TBD") return -1;
                if (locB === "TBD") return 1;
                return locA.localeCompare(locB, undefined, { numeric: true, sensitivity: 'base' });
            });
        }

        // Step 4: Draw a regular data grid cleanly inside your mount zone view layout parameters
        const idIndex = sheetConfig.columns.findIndex(c => c.header === "mae_id");
        const visibleIndices = [];
        
        let htmlTable = `<table class="inventory-table" id="main-data-table"><thead><tr>`;
        sheetConfig.columns.forEach((col, index) => {
            if (col.hidden !== true) { htmlTable += `<th>${col.header}</th>`; visibleIndices.push(index); }
        });
        htmlTable += `</tr></thead><tbody>`;

        flattenedSubset.forEach(row => {
            const cells = row.values;
            const rawMaeId = (idIndex !== -1) ? cells[idIndex] : '';
            htmlTable += `<tr data-row-index="${row.index}" data-mae-id="${rawMaeId}">`;
            visibleIndices.forEach(idx => {
                const colDef = sheetConfig.columns[idx];
                let displayValue = cells[idx] ?? '';
                if (colDef.format && colDef.format.includes("$") && displayValue !== "") {
                    displayValue = formatCurrency(displayValue);
                }
                htmlTable += `<td class="locked-cell">${displayValue}</td>`;
            });
            htmlTable += `</tr>`;
        });
        htmlTable += `</tbody></table>`;
        
        // Restore selection form card along with newly mounted table subset data array 
        const container = document.getElementById("table-container");
        container.innerHTML = `
            <div class="form-card" style="border-left:5px solid #8e44ad; padding:25px; background:#fff; margin-bottom:20px;">
                <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">Select Current Status Target</h4>
                <div style="display:flex; gap:15px; flex-wrap:wrap; align-items:center;">
                    <select id="mae-resell-status-selector" class="edit-dropdown" style="flex:1; max-width:400px; height:50px;">
                        <option value="">-- Choose Status Group --</option>
                        ${sheetConfig.columns.find(c => c.header === "Current Status").options.map(opt => `<option value="${opt}" ${opt === selectedStatus ? 'selected' : ''}>${opt}</option>`).join('')}
                    </select>
                    <button class="action-btn" style="background:#8e44ad; height:50px; font-size:1rem;" onclick="UI.executeResellStatusFilter()">📊 Generate Filtered Table</button>
                </div>
            </div>
            <div id="status-filtered-table-mount">${htmlTable}</div>
        `;

        // Step 5: Mount custom context action bar holding your required Print and Back selectors
        if (actionZone) {
            actionZone.innerHTML = `
                <div class="command-bar">
                    <button class="action-btn" onclick="window.UI.printStatusPivotTable()" style="background:#27ae60;">🖨️ Print Filtered View</button>
                    <button class="action-btn" onclick="window.loadTableData('Resell_Inventory')">← Back to Resell Inventory</button>
                </div>`;
        }

    } catch (err) {
        console.error("MAE Engine Pivot System Failure:", err);
        tableMount.innerHTML = "<p style='color:red; padding:20px;'>Failed to correctly slice status array parameters.</p>";
    }
},

printStatusPivotTable() {
    const tableMount = document.getElementById("status-filtered-table-mount");
    if (!tableMount || !this.activeStatusPivotLabel) {
        alert("System Error: No active filtered data grid available to print.");
        return;
    }

    const tableElement = tableMount.querySelector("table");
    if (!tableElement) {
        alert("System Error: Table asset element layout mismatch.");
        return;
    }

    const customPrintTitle = `Resell Item with Current Status: ${this.activeStatusPivotLabel}`;
    const printWindow = window.open('', '_blank');

    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>${customPrintTitle}</title>
            <style>
                /* 🌟 MAE ENGINE UPGRADE: FORCE LANDSCAPE PAGE DIRECTIVES 🌟 */
                @page {
                    size: landscape;
                    margin: 0.4in;
                }

                body { 
                    background: #ffffff !important; 
                    color: #000000 !important; 
                    padding: 20px; 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }

                /* 🌟 MAE ENGINE UPGRADE: SCALE UP BRANDED HERO TYPOGRAPHY 🌟 */
                h2 { 
                    margin: 0 0 5px 0; 
                    text-transform: uppercase; 
                    letter-spacing: 1px; 
                    font-size: 22pt !important; /* Scaled up to match standard high-contrast headers */
                    font-weight: 800;
                    color: #000000;
                }
                h4 { 
                    margin: 0 0 25px 0; 
                    color: #333333; 
                    font-size: 14pt !important; /* Expanded descriptor row sizing */
                    font-weight: bold;
                    border-bottom: 3px solid #000000; /* Rugged thick dividing border line */
                    padding-bottom: 8px;
                    text-transform: uppercase;
                }

                table { 
                    width: 100% !important; 
                    border-collapse: collapse !important; 
                    margin-top: 15px;
                    page-break-inside: auto;
                }
                tr {
                    page-break-inside: avoid;
                    page-break-after: auto;
                }
                th, td { 
                    border: 1px solid #000000 !important; 
                    padding: 10px 12px !important; /* Slightly padded for clear clipboard viewing */
                    text-align: left; 
                    font-size: 10pt !important;
                    color: #000000 !important;
                    background: #ffffff !important;
                }
                th { 
                    background-color: #f2f2f2 !important; 
                    font-weight: bold !important; 
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                .edit-only-cell, .print-force-hide, button, .form-card, input[type="hidden"] { 
                    display: none !important; 
                    width: 0 !important;
                    height: 0 !important;
                    visibility: hidden !important;
                }
            </style>
        </head>
        <body>
            <h2>MAE Workshop Inventory System</h2>
            <h4>${customPrintTitle}</h4>
            <div>${tableElement.outerHTML}</div>
            <script>
                window.onload = function() {
                    setTimeout(() => {
                        window.print(); 
                        window.close();
                    }, 250);
                };
            </script>
        </body>
        </html>
    `);
    printWindow.document.close();
},
//======  END Compiling and Printing the Filtered Subset (Resell Inventory)

//======= Print Location_ID search results ========
printInspectedLocationTable() {
        const table = document.getElementById("main-data-table");
        if (!table || !this.activeInspectedLocationLabel) {
            alert("System Error: No active inspection data grid available to print.");
            return;
        }

        const today = new Date().toLocaleDateString('en-US');
        const customPrintTitle = `Inventory Items Held At Location_ID: ${this.activeInspectedLocationLabel}`;
        const printWindow = window.open('', '_blank');

        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>${customPrintTitle}</title>
                <style>
                    @page { size: landscape; margin: 0.4in; }
                    body { background: #ffffff !important; color: #000000 !important; padding: 20px; font-family: 'Segoe UI', Tahoma, Verdana, sans-serif; }
                    h2 { margin: 0 0 5px 0; text-transform: uppercase; letter-spacing: 1px; font-size: 22pt !important; font-weight: 800; color: #000000; }
                    h4 { margin: 0 0 25px 0; color: #333333; font-size: 14pt !important; font-weight: bold; border-bottom: 3px solid #000000; padding-bottom: 8px; text-transform: uppercase; }
                    table { width: 100% !important; border-collapse: collapse !important; margin-top: 15px; }
                    th, td { border: 1px solid #000000 !important; padding: 10px 12px !important; text-align: left; font-size: 10pt !important; color: #000000 !important; background: #ffffff !important; }
                    th { background-color: #f2f2f2 !important; font-weight: bold !important; text-transform: uppercase; }
                    button, .action-btn, td:nth-child(2), th:nth-child(2) { display: none !important; } /* Cleanly drop the "Action" column from final hardcopy sheets */
                </style>
            </head>
            <body>
                <h2>MAE Workshop Inventory System</h2>
                <h4>${customPrintTitle} (as of ${today})</h4>
                <div>${table.outerHTML}</div>
                <script>
                    window.onload = function() {
                        setTimeout(() => { window.print(); window.close(); }, 250);
                    };
                </script>
            </body>
            </html>
        `);
        printWindow.document.close();
    },
//=== END Print Location_ID search results ======

// =========================================================================
    //  LOCATION_ID CONTENTS INSPECTOR MODULE
    // =========================================================================
    renderLocationInspectorControls() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        
        title.innerText = "Location Inventory: Filter View By Storage Spot";
        window.currentTable = "location_inspector"; // Establish virtual routing state context

        let html = `
            <div class="form-card" style="border-left:5px solid #8e44ad; padding:25px; background:#fff; margin-bottom:20px;">
                <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">Select Physical Storage Spot</h4>
                <div style="display:flex; gap:15px; flex-wrap:wrap; align-items:center;">
                    <select id="mae-inspection-location-selector" class="edit-dropdown" style="flex:1; max-width:400px; height:50px;">
                        <option value="">-- Choose Active Location_ID --</option>
                        ${window.maeLocations.map(loc => `<option value="${loc}">${loc}</option>`).join('')}
                    </select>
                    <button class="action-btn" style="background:var(--primary); height:50px; font-size:1rem;" onclick="UI.executeLocationInspection()">🔍 Inspect Storage Spot</button>
                </div>
            </div>
            <div id="location-inspected-table-mount"></div>
        `;

        container.innerHTML = html;
        this.renderCommandBar("location_inspector");
    },

    async executeLocationInspection() {
        const selector = document.getElementById("mae-inspection-location-selector");
        const targetLocation = selector ? selector.value : "";
        const tableContainer = document.getElementById("table-container");

        if (!targetLocation) { 
            alert("Please select a physical storage spot parameter first."); 
            return; 
        }

        this.showLoading(`Scanning cache for items in [${targetLocation}]...`);

        // The exact priority sequence of your active inventory worksheets
        const tablesToScan = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
        let aggregatedResults = [];

        try {
            for (const tableName of tablesToScan) {
                const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === tableName);
                if (!sheetConfig) continue;

                // Query cached items through your dashboard data module
                const dataRows = await window.Dashboard.getFullTableData(tableName);
                if (!dataRows || dataRows.length === 0) continue;

                const locIdx = sheetConfig.columns.findIndex(c => c.header === "Location_ID");
                // Discover the first active descriptor header tracking asset name strings natively
                const nameIdx = sheetConfig.columns.findIndex(c => !c.hidden && (c.header.includes("Name") || c.header.includes("Tool") || c.header.includes("Description")));

                if (locIdx === -1) continue;

                dataRows.forEach(rowObj => {
                    // Extract values cleanly from index 0 inside Graph's 2D array wrapper
                    const rawCells = (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values[0] : rowObj.values;
                    
                    if (rawCells && rawCells[locIdx] !== undefined && rawCells[locIdx] !== null) {
                        const currentLocValue = String(rawCells[locIdx]).trim().toUpperCase();
                        
                        if (currentLocValue === targetLocation.trim().toUpperCase()) {
                            aggregatedResults.push({
                                category: sheetConfig.tabName,
                                itemName: rawCells[nameIdx] || "N/A",
                                tableName: tableName
                            });
                        }
                    }
                });
            }

            // Save matched arrays locally so the print subsystem can access them cleanly
            this.activeInspectedLocationData = aggregatedResults;
            this.activeInspectedLocationLabel = targetLocation;

            // Handle the empty storage spot layout scenario gracefully
            if (aggregatedResults.length === 0) {
                tableContainer.innerHTML = `
                    <div class="form-card" style="text-align:center; padding:40px; margin:20px;">
                        <h3 style="color:var(--primary); margin-top:0;">📋 Empty Storage Spot</h3>
                        <p>Zero active inventory items are currently mapped to Location_ID: <b>${targetLocation}</b>.</p>
                        <button class="action-btn" onclick="window.loadTableData('Location')">Return to Location Map</button>
                    </div>`;
                window.currentTable = "location_inspector";
                this.renderCommandBar("location_inspector");
                return;
            }

            // Group the metrics dynamically by tab name for professional workshop presentation
            const grouped = aggregatedResults.reduce((acc, item) => {
                if (!acc[item.category]) acc[item.category] = [];
                acc[item.category].push(item);
                return acc;
            }, {});

            // 🌟 THE SOLUTION GRID: Build a clean view mapping exactly what sits inside the spot 🌟
            let htmlGrid = `
                <div style="padding: 20px 0;">
                    <div class="form-card" style="border-left:5px solid #8e44ad; padding:20px; background:#fff; margin-bottom:20px; display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <h3 style="margin:0; color:var(--primary);">Active Inspection: Storage Spot [${targetLocation}]</h3>
                            <p style="margin:5px 0 0 0; font-size:0.9rem; color:#666;">Displaying all inventory records matching this physical landmark vector.</p>
                        </div>
                        <button class="action-btn" onclick="UI.renderLocationInspectorControls()" style="background:#8e44ad;">🔄 Change Location</button>
                    </div>
                    <table class="inventory-table" id="main-data-table">`;

            for (const [category, items] of Object.entries(grouped)) {
                htmlGrid += `
                    <thead>
                        <tr>
                            <th colspan="2" style="background:#34495e; color:white; padding:12px; position:sticky; top:0; z-index:9999;">
                                ${category.toUpperCase()} (${items.length} Items Present)
                            </th>
                        </tr>
                        <tr>
                            <th style="width:75%; background: var(--primary) !important;">Item Description / Model Identification</th>
                            <th style="width:25%; text-align:center; background: var(--primary) !important;">Operational Action</th>
                        </tr>
                    </thead>
                    <tbody>`;
                
                items.forEach(item => {
                    htmlGrid += `
                        <tr>
                            <td class="locked-cell" style="padding: 12px 15px;"><b>${item.itemName}</b></td>
                            <td style="text-align:center; padding: 12px 15px;">
                                <button class="action-btn" style="padding: 5px 12px; font-size: 0.85rem; background: var(--accent);" onclick="window.loadTableData('${item.tableName}')">
                                    ✏️ Go to Table
                                </button>
                            </td>
                        </tr>`;
                });
                htmlGrid += `</tbody>`;
            }
            htmlGrid += `</table></div>`;

            // Overwrite the table viewport container with our targeted structural list
            tableContainer.innerHTML = htmlGrid;

            // Set current state flags to swap bottom command bar button actions contextually
            window.currentTable = "location_inspector";
            this.renderCommandBar("location_inspector");

        } catch (err) {
            console.error("MAE Engine Location Inspection Failure:", err);
            this.showError("Failed to safely load inventory content records.");
        }
    },
//=======  END: LOCATION_ID CONTENTS INSPECTOR MODULE  ================

// =========  GLOBAL FOCUSED-FIELD SEARCH INTERFACE ===============
    renderSearchControls() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        
        title.innerText = "Inventory Search: Find Item Physical Location";
        window.currentTable = "inventory_search"; // Establish virtual routing state context

        let html = `
            <div class="form-card" style="border-left:5px solid #2980b9; padding:25px; background:#fff; margin-bottom:20px;">
                <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">🔍 Narrow Focused Item Lookup</h4>
                <p style="font-size:0.85rem; color:#666; margin:0 0 15px 0;">Enter a partial item name, brand, model description, or category keyword to cross-reference all tables and find where it is stored.</p>
                <div style="display:flex; gap:15px; flex-wrap:wrap; align-items:center;">
                    <input type="text" id="mae-global-search-input" placeholder="Enter keyword (e.g., DeWalt, Bolt, Bandsaw)..." style="flex:1; max-width:400px; height:50px; padding: 0 15px; border:1px solid var(--border); border-radius:4px; font-size:1rem;" autofocus>
                    <button class="action-btn" style="background:var(--accent); height:50px; font-size:1rem;" onclick="triggerAssetSearch()">Run Search Matrix</button>
                </div>
            </div>
            <div id="location-search-results-mount"></div>
        `;

        container.innerHTML = html;
        this.renderCommandBar("inventory_search");

        // Set focus to text field automatically for fluid HID keyboard/scanner compatibility
        setTimeout(() => document.getElementById("mae-global-search-input")?.focus(), 50);
    },

    triggerAssetSearch() {
        const inputEl = document.getElementById("mae-global-search-input");
        const queryText = inputEl ? inputEl.value.trim() : "";
        if (!queryText) {
            alert("Please enter a tool or part description search keyword phrase first.");
            return;
        }
        // Direct execution routing back down to your core app.js scanning logic script block
        window.executeFocusedAssetSearch(queryText);
    },
//==========  END: GLOBAL FOCUSED-FIELD SEARCH INTERFACE ===============

// =========================================================================
    // THE CENTRAL INTAKE REGISTRATION PORTAL 
    //   CENTRAL ASSET WIZARD 
    // =========================================================================
    renderCentralRegistrationWizard() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        title.innerText = "Administrative: Centralized Item Intake Portal";
        
        // Lock the router configuration state flag
        window.currentTable = "inventory_registration";

        let html = `
            <div class="form-card" style="border-left: 6px solid var(--accent); background:#fff; padding: 25px; margin-bottom: 25px;">
                <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">⚡ Central Asset Registration Wizard</h4>
                <p style="font-size:0.85rem; color:#666; margin:0 0 15px 0;">STAGE 1: Token Identification Gate. Select your target table, then choose to register an UNTAGGED bulk item or scan a fresh sticker token.</p>
                
                <div style="display: flex; flex-direction: column; gap: 15px; max-width: 500px; margin-bottom: 20px;">
                    <div style="display: flex; flex-direction: column;">
                        <label style="font-size:0.8rem; font-weight:bold; color:var(--primary); margin-bottom:5px;">Target Inventory Classification Sheet</label>
                        <select id="mae-central-table-selector" class="edit-dropdown" style="height:45px; font-size:0.95rem;">
                            <option value="">-- Choose Target Table --</option>
                            <option value="Shop_Machinery">Shop Machinery</option>
                            <option value="Shop_Power_Tools">Shop Power Tools</option>
                            <option value="Shop_Hand_Tools">Shop Hand Tools</option>
                            <option value="Shop_Consumables">Shop Consumables</option>
                            <option value="Resell_Inventory">Resell Inventory</option>
                        </select>
                    </div>

                    <div style="display: flex; flex-direction: column; position: relative;">
                        <label style="font-size:0.8rem; font-weight:bold; color:var(--primary); margin-bottom:5px;">Scan Fresh Sticker Token (Advanced Tier Focus)</label>
                        <input type="text" id="field-Tag_ID" placeholder="Click here and scan physical label roll..." style="height:45px; border:2px solid var(--border); padding:0 12px; font-weight:bold; font-size:1rem; background: #fffde7;">
                        <div id="wizard-tag-feedback" style="margin-top: 5px; font-size: 0.8rem; font-weight: bold;"></div>
                    </div>
                </div>

                <div style="display: flex; gap: 15px;">
                    <button class="action-btn" onclick="UI.processWizardStageOneScan()" style="background:var(--primary); height:45px; font-weight:bold; flex: 1;">⚡ Verify Scanned Tag</button>
                    <button class="action-btn" onclick="UI.processWizardStageOneUntagged()" style="background:#7f8c8d; height:45px; font-weight:bold; flex: 1;">📦 Proceed as UNTAGGED</button>
                    <button class="action-btn" onclick="UI.resetCentralRegistrationWizard()" style="background:#c0392b; height:45px; font-weight:bold; flex: 1;">🔄 Clear / Reset Form</button>
                </div>
            </div>
            <div id="central-form-render-zone"></div>
        `;
        container.innerHTML = html;
        this.renderCommandBar("");

        // Focus management pass without any browser-native change listeners attached
        setTimeout(() => {
            const input = document.getElementById("field-Tag_ID");
            if (input) {
                input.focus();
                
                // 🌟 THE SAFE INTERCEPT GATE 🌟
                // If a user manual types or if an enter event hits this box, block native browser form 
                // submission completely, and route it through our unified verification instead.
                input.onkeydown = (e) => {
                    if (e.key === 'Enter') {
                        e.preventDefault(); // Abort browser form triggers
                        UI.processWizardStageOneScan(); // Run our controlled lookup pass instead
                    }
                };
            }
        }, 100);
    },
    // 2. ADD NEW function to process an actual scanned barcode in Stage 1
    async processWizardStageOneScan() {
        // REGISTER TRANSACTION FOR MANUAL ENTRIES / CLICK ACTIONS
        const currentTransactionId = Date.now();
        window.activeScanTransactionId = currentTransactionId;

        const tableSelect = document.getElementById("mae-central-table-selector");
        const tagInput = document.getElementById("field-Tag_ID");
        const feedback = document.getElementById("wizard-tag-feedback");
        const formZone = document.getElementById("central-form-render-zone");

        const targetTable = tableSelect ? tableSelect.value : "";
        const rawTag = tagInput ? tagInput.value.trim() : "";

        if (!targetTable) {
            alert("Mandatory Selection Required:\n\nPlease choose a target inventory classification table before checking the tag.");
            if (tagInput) tagInput.value = "";
            return;
        }

        if (!rawTag) {
            alert("Scan Required:\n\nPlease focus the input field and scan a sticker or choose the UNTAGGED track.");
            return;
        }

        const cleanTag = window.Labels.extractCleanId(rawTag).toUpperCase();
        if (tagInput) tagInput.value = cleanTag;

        if (feedback) {
            feedback.style.color = "var(--primary)";
            feedback.innerText = "⌛ Verifying label uniqueness across partitions...";
        }

        // --- ANTI-COLLISION SECURITY CHECK ---
        const isCollision = await window.verifyTagUniquenessCrossTable(cleanTag);
        
        // CIRCUIT BREAKER CHECK
        if (window.activeScanTransactionId !== currentTransactionId) {
            console.warn("MAE Circuit Breaker: Wizard execution terminated mid-fetch via form reset.");
            return;
        }

        if (isCollision) {
            if (tagInput) {
                tagInput.value = "";
                tagInput.style.borderColor = "#e74c3c";
                tagInput.style.backgroundColor = "#fadbd8";
            }
            if (feedback) {
                feedback.style.color = "#c0392b";
                feedback.innerHTML = `❌ COLLISION ERROR: Barcode [${cleanTag}] is already registered to an asset row inside the spreadsheet ledger. Choose another sticker.`;
            }
            formZone.innerHTML = "";
            return;
        }

        // 🌟 DEPLOY THE CUSTOM INDUSTRIAL MODAL OVERLAY 🌟
        // Bypasses fragile browser confirm() windows and routes selections through your system layout
        this.renderTagTypeWizardModal(targetTable, cleanTag, currentTransactionId);
    },
    // 🌟 ADD NEW custom layout popup model to clean up the intake handshake selection process 🌟
    // 🌟 MODIFIED WIZARD OVERLAY: MULTI-TIER CHOICES WITH DISCIPLINARY FALLBACKS 🌟
  renderTagTypeWizardModal(targetTable, cleanTag, currentTransactionId) {
    const existingModal = document.getElementById("mae-wizard-modal-overlay");
    if (existingModal) existingModal.remove();

    const modalHtml = `
      <div id="mae-wizard-modal-overlay" style="position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0, 0, 0, 0.6); z-index: 10000; display: flex; align-items: center; justify-content: center; padding: 20px; box-sizing: border-box;">
        <div id="mae-wizard-modal-card" style="background: #ffffff; border-top: 8px solid var(--accent); border-radius: 8px; width: 100%; max-width: 550px; padding: 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.3); box-sizing: border-box; text-align: center; transition: all 0.3s ease;">
          <h2 style="margin: 0 0 15px 0; color: var(--primary); text-transform: uppercase; font-weight: 800; font-size: 1.6rem; letter-spacing: 1px;">
            SELECT ITEM CATEGORY
          </h2>
          <p style="font-size: 0.95rem; color: #444; margin-bottom: 25px; line-height: 1.5;" id="mae-modal-text-prompt">
            Barcode <b style="background: #fffde7; padding: 2px 6px; border: 1px dashed var(--accent); font-family: monospace;">${cleanTag}</b> is verified UNIQUE and available.<br>
            Specify how this physical label will anchor inside your warehouse layout map.
          </p>
          <div style="display: flex; flex-direction: column; gap: 12px; width: 100%;" id="mae-modal-buttons-silo">
            <button class="action-btn" id="modal-btn-unique" style="background: var(--primary); height: 55px; font-size: 1.1rem; font-weight: bold; width: 100%;">
              🎯 UNIQUE (One Tag for One Single Asset)
            </button>
            <button class="action-btn" id="modal-btn-multiple" style="background: #e67e22; height: 55px; font-size: 1.1rem; font-weight: bold; width: 100%;">
              📦 MULTIPLE (One Tag for a Bin / Drawer Group)
            </button>
            <button class="cancel-btn" id="modal-btn-return" style="background: #7f8c8d; height: 45px; font-size: 1rem; font-weight: bold; margin-left: 0; width: 100%;">
              ↩️ Return to Wizard
            </button>
          </div>
        </div>
      </div>
    `;
    
    document.body.insertAdjacentHTML("beforeend", modalHtml);

    const modalOverlay = document.getElementById("mae-wizard-modal-overlay");
    const modalCard = document.getElementById("mae-wizard-modal-card");
    const tableSelect = document.getElementById("mae-central-table-selector");
    const tagInput = document.getElementById("field-Tag_ID");
    const feedback = document.getElementById("wizard-tag-feedback");

    // --- TRACK A: UNIQUE SELECTION ENGINE ---
    document.getElementById("modal-btn-unique").onclick = () => {
      modalOverlay.remove();
      window.maeWizardActiveCategory = "UNIQUE"; // Explicitly register Unique status
      if (feedback) {
        feedback.style.color = "#27ae60";
        feedback.innerHTML = `✅ Structure Verified: Tag [${cleanTag}] locked as UNIQUE.`;
        tagInput.style.borderColor = "#27ae60";
        tagInput.style.backgroundColor = "#e8f8f5";
        tagInput.disabled = true;
        tableSelect.disabled = true;
      }
      window.renderCentralRegistrationWizardStageTwo(targetTable, cleanTag, "UNIQUE");
    };

    // --- TRACK B: MULTIPLE CHOOSER GATE (WITH DEFENSIVE FALLBACK) ---
    document.getElementById("modal-btn-multiple").onclick = () => {
      // 🌟 INTERLOCK ENGAGED: Pre-load the stricter "By_Location" as your default background state!
      window.maeWizardActiveCategory = "By_Location";
      
      modalCard.style.borderTopColor = "#e67e22"; // Shift aesthetic tint to warning orange
      
      document.getElementById("mae-modal-text-prompt").innerHTML = `
        <span style="color:#e67e22; font-weight:bold; font-size:1.1rem; display:block; margin-bottom:10px;">⚠️ MANDATORY TRACK SPECIFICATION</span>
        You selected a <b>MULTIPLE</b> shared grouping tag.<br>
        Specify how this cluster is organized, or choose "Accept Default" to apply By_Location discipline.
      `;

      document.getElementById("mae-modal-buttons-silo").innerHTML = `
        <button class="action-btn" id="modal-btn-sub-location" style="background: #e67e22; height: 55px; font-size: 1.05rem; font-weight: bold; width: 100%; text-align:left; padding-left:20px;">
          📦 BY LOCATION (Physical box/bin. All items stay together)
        </button>
        <button class="action-btn" id="modal-btn-sub-topic" style="background: #2980b9; height: 55px; font-size: 1.05rem; font-weight: bold; width: 100%; text-align:left; padding-left:20px;">
          📋 BY TOPIC (Thematic cluster. Items live in different spots)
        </button>
        <button class="cancel-btn" id="modal-btn-sub-default-exit" style="background: var(--primary); color:white; height: 45px; font-size: 1rem; font-weight: bold; margin-left: 0; width: 100%;">
          🏁 Accept Default & Start Intake Workspace
        </button>
      `;

      // Option B1: Explicitly confirming BY LOCATION
      document.getElementById("modal-btn-sub-location").onclick = () => {
        modalOverlay.remove();
        window.maeWizardActiveCategory = "By_Location";
        if (feedback) {
          feedback.style.color = "#e67e22";
          feedback.innerHTML = `✅ Structure Verified: Tag [${cleanTag}] locked as MULTIPLE (By_Location).`;
          tagInput.style.borderColor = "#e67e22";
          tagInput.style.backgroundColor = "#fffde7";
          tagInput.disabled = true;
          tableSelect.disabled = true;
        }
        window.renderCentralRegistrationWizardStageTwo(targetTable, cleanTag, "MULTIPLE");
      };

      // Option B2: Explicitly changing profile to BY TOPIC
      document.getElementById("modal-btn-sub-topic").onclick = () => {
        modalOverlay.remove();
        window.maeWizardActiveCategory = "By_Topic";
        if (feedback) {
          feedback.style.color = "#2980b9";
          feedback.innerHTML = `✅ Structure Verified: Tag [${cleanTag}] locked as MULTIPLE (By_Topic).`;
          tagInput.style.borderColor = "#2980b9";
          tagInput.style.backgroundColor = "#e8f4f8";
          tagInput.disabled = true;
          tableSelect.disabled = true;
        }
        window.renderCentralRegistrationWizardStageTwo(targetTable, cleanTag, "MULTIPLE");
      };

      // Option B3: The Disciplinary Fallback Tracker Route
      // If they click to bypass, the pre-loaded default "By_Location" remains locked inside the session variable.
      document.getElementById("modal-btn-sub-default-exit").onclick = () => {
        modalOverlay.remove();
        console.log(`MAE Defensive Guard: Operator bypassed category specification. Auto-healing to stricter option: [${window.maeWizardActiveCategory}].`);
        if (feedback) {
          feedback.style.color = "#e67e22";
          feedback.innerHTML = `🔒 Auto-Assigned Fallback: Tag [${cleanTag}] locked as MULTIPLE (By_Location).`;
          tagInput.style.borderColor = "#e67e22";
          tagInput.style.backgroundColor = "#fffde7";
          tagInput.disabled = true;
          tableSelect.disabled = true;
        }
        window.renderCentralRegistrationWizardStageTwo(targetTable, cleanTag, "MULTIPLE");
      };
    };

    // --- TRACK C: ABORT CANCEL HANDLER ---
    document.getElementById("modal-btn-return").onclick = () => {
      modalOverlay.remove();
      window.activeScanTransactionId = null;
      window.maeWizardActiveCategory = null; // Clear the temporary mailbox entirely
      if (tagInput) {
        tagInput.value = "";
        tagInput.disabled = false;
        tagInput.style.borderColor = "var(--border)";
        tagInput.style.backgroundColor = "#fffde7";
        tagInput.focus();
      }
      if (feedback) feedback.innerHTML = "";
    };
  },

    // 3. ADD NEW function to process an untagged item path in Stage 1
    processWizardStageOneUntagged() {
        const tableSelect = document.getElementById("mae-central-table-selector");
        const targetTable = tableSelect ? tableSelect.value : "";
        const tagInput = document.getElementById("field-Tag_ID");
        const feedback = document.getElementById("wizard-tag-feedback");

        if (!targetTable) {
            alert("Mandatory Selection Required:\n\nPlease choose a target inventory classification table before proceeding.");
            return;
        }

        if (feedback) {
            feedback.style.color = "var(--accent)";
            feedback.innerHTML = `📦 Fallback Default Activated: Asset will be registered under structural token [UNTAGGED].`;
            if (tagInput) {
                tagInput.value = "UNTAGGED";
                tagInput.disabled = true;
                tagInput.style.borderColor = "var(--border)";
                tagInput.style.backgroundColor = "#eee";
            }
            tableSelect.disabled = true;
        }

        // Deploy fields with the default untagged parameters hardcoded
        this.renderCentralRegistrationWizardStageTwo(targetTable, "UNTAGGED", "UNIQUE");
    },
    async processWizardStageOneScan() {
        // REGISTER TRANSACTION FOR MANUAL ENTRIES / CLICK ACTIONS
        const currentTransactionId = Date.now();
        window.activeScanTransactionId = currentTransactionId;

        const tableSelect = document.getElementById("mae-central-table-selector");
        const tagInput = document.getElementById("field-Tag_ID");
        const feedback = document.getElementById("wizard-tag-feedback");
        const formZone = document.getElementById("central-form-render-zone");

        const targetTable = tableSelect ? tableSelect.value : "";
        const rawTag = tagInput ? tagInput.value.trim() : "";

        if (!targetTable) {
            alert("Mandatory Selection Required:\n\nPlease choose a target inventory classification table before checking the tag.");
            if (tagInput) tagInput.value = "";
            return;
        }

        if (!rawTag) {
            alert("Scan Required:\n\nPlease focus the input field and scan a sticker or choose the UNTAGGED track.");
            return;
        }

        const cleanTag = window.Labels.extractCleanId(rawTag).toUpperCase();
        if (tagInput) tagInput.value = cleanTag;

        if (feedback) {
            feedback.style.color = "var(--primary)";
            feedback.innerText = "⌛ Verifying label uniqueness across partitions...";
        }

        // --- ANTI-COLLISION SECURITY CHECK ---
        const isCollision = await window.verifyTagUniquenessCrossTable(cleanTag);
        
        // CIRCUIT BREAKER CHECK
        if (window.activeScanTransactionId !== currentTransactionId) {
            console.warn("MAE Circuit Breaker: Wizard execution terminated mid-fetch via form reset.");
            return;
        }

        if (isCollision) {
            if (tagInput) {
                tagInput.value = "";
                tagInput.style.borderColor = "#e74c3c";
                tagInput.style.backgroundColor = "#fadbd8";
            }
            if (feedback) {
                feedback.style.color = "#c0392b";
                feedback.innerHTML = `❌ COLLISION ERROR: Barcode [${cleanTag}] is already registered to an asset row inside the spreadsheet ledger. Choose another sticker.`;
            }
            formZone.innerHTML = "";
            return;
        }

        // 🌟 EXPLICITLY BOUND ENGINE CALL 🌟
        // Changed 'this' to 'UI' to guarantee execution scope stability
        UI.renderTagTypeWizardModal(targetTable, cleanTag, currentTransactionId);
    },
    // 🌟 STAGE TWO GENERATION: FORM COMPILATION & EXPLICIT ATTRIBUTE INJECTION PASS
    renderCentralRegistrationWizardStageTwo(targetTable, validatedTagId, tagType, isSubsequentEntry = false) {
        const formZone = document.getElementById("central-form-render-zone");
        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === targetTable);

        // 1. INITIALIZE SESSION MEMORY LEDGER (Only on the very first entry pass)
        if (!isSubsequentEntry) {
            window.maeWizardSessionItems = [];
        }

        // 2. COMPILE UNIFIED LAYOUT: Form Entry Input Grid on top, Visual Live List underneath
        formZone.innerHTML = `
            <div id="mae-wizard-form-mount"></div>
            <!-- 🌟 THE RUNNING LIVE SESSION LIST LEDGER 🌟 -->
            <div id="mae-wizard-live-list-panel" style="margin-top: 30px; background: #ffffff; border: 1px solid var(--border); border-top: 4px solid var(--primary); padding: 20px; border-radius: 4px; display: ${window.maeWizardSessionItems.length > 0 ? 'block' : 'none'};">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 15px;">
                <h4 style="margin:0; color:var(--primary); text-transform:uppercase; font-weight:800; font-size:0.95rem;">📋 Items Registered in this Session</h4>
                <span style="background:var(--primary); color:white; padding:2px 8px; border-radius:10px; font-size:0.8rem; font-weight:bold;" id="mae-session-badge-count">${window.maeWizardSessionItems.length}</span>
      <     /div>
            <div id="mae-wizard-session-grid-mount"></div>
            <!-- TERMINATION ACTION CONTROL -->
            <button class="action-btn" onclick="UI.finalizeWizardBatchSession()" style="width:100%; height:50px; background:var(--primary); font-weight:bold; font-size:1.1rem; margin-top:20px; text-transform:uppercase; letter-spacing:0.5px;">
                🏁 Finished Adding Items (Close Session)
            </button>
            </div>
        `;

        // 3. TRIGGER DETERMINISTIC ENTRY FORM GENERATION (Race-free argument pass)
        window.UI.renderEntryForm('add', targetTable, sheetConfig, async () => {
            // Re-enforce Stage 1 metrics right before submission
            const tagField = document.getElementById("field-Tag_ID");
            const typeField = document.getElementById("field-Tag_Type");
            if (tagField) tagField.value = validatedTagId;
            if (typeField) typeField.value = tagType;

            // Harvest values for our local visual list BEFORE submission clears them
            const descFieldId = `field-${sheetConfig.columns.find(c => c.header.includes("Description") || c.header.includes("Name")).header.replace(/\s+/g, '')}`;
            const locFieldId = `field-Location_ID`;
            const enteredDescription = document.getElementById(descFieldId)?.value || "N/A";
            const enteredLocation = document.getElementById(locFieldId)?.value || "TBD";

            // Submit row data asynchronously straight up to OneDrive
            const success = await window.submitNewRow(targetTable, sheetConfig);
            if (success) {
            // Append item specifications into our local session memory tracker array
            window.maeWizardSessionItems.push({
                description: enteredDescription,
                location: enteredLocation,
                timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
            });

            // If managing a MULTIPLE tag session, clear descriptive cells and re-load form inline
            if (tagType === "MULTIPLE") {
                // 🌟 SYSTEM LAYOUT PARKING SHIELD   🌟
                // If a global table load wiped our wizard canvas assembly, restore the container context smoothly
                let formZoneCheck = document.getElementById("central-form-render-zone");
                if (!formZoneCheck) {
                console.log("MAE Wizard Engine: Restoring intake portal rendering canvas layout anchors...");
                const tableContainer = document.getElementById("table-container");
                tableContainer.innerHTML = `
                    <div class="form-card" style="border-left: 6px solid var(--accent); background:#fff; padding: 25px; margin-bottom: 25px;">
                    <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">⚡ Central Asset Registration Wizard</h4>
                    <p style="font-size:0.85rem; color:#666; margin:0 0 15px 0;">CONTAINER BATCH MODE ACTIVE: Registering multiple entries to Tag ID [${validatedTagId}].</p>
                    </div>
                    <div id="central-form-render-zone"></div>`;
                }

                // Now safe to loop back inline, passing true to preserve the running array log rows
                window.UI.renderCentralRegistrationWizardStageTwo(targetTable, validatedTagId, "MULTIPLE", true);
        
                if (typeof window.UI.renderWizardSessionListGrid === "function") {
                window.UI.renderWizardSessionListGrid();
                }
            } else {
                // Unique track item complete: clean up and return to dashboard cockpit
                formZone.innerHTML = "";
                window.currentTable = "Master_Dashboard";
                window.loadTableData("Master_Dashboard");
            }
            }
        }, null, null, { tagId: validatedTagId, tagType: tagType });

        // 4. CLEAN AND FLUID ELEMENT ATTACHMENT MIGRATION
        const formCard = document.getElementById("entry-form");
        if (formCard) {
            // Instantly append form layout directly inside our wizard mount panel view frame
            document.getElementById("mae-wizard-form-mount").appendChild(formCard);
            const closeBtn = formCard.querySelector(".close-x");
            if (closeBtn) closeBtn.remove();

            // Re-direct active focus instantly to the contextual Item Description tracking row box
            const descriptiveInputField = formCard.querySelector("input[type='text']:not(#field-Tag_ID)");
            if (descriptiveInputField) {
            descriptiveInputField.focus();
            descriptiveInputField.style.backgroundColor = "#fffde7"; // Focus Highlight Tint
            }
        }

        // Draw current data lists if performing subsequent multi-entry passes
        if (isSubsequentEntry && typeof window.UI.renderWizardSessionListGrid === "function") {
            window.UI.renderWizardSessionListGrid();
        }
        },
            // 🌟 ADD NEW method to compile and render your live session list grid rows on the fly 🌟
            renderWizardSessionListGrid() {
        const gridMount = document.getElementById("mae-wizard-session-grid-mount");
        const panel = document.getElementById("mae-wizard-live-list-panel");
        const badge = document.getElementById("mae-session-badge-count");

        if (!gridMount || !window.maeWizardSessionItems || window.maeWizardSessionItems.length === 0) return;

        // Reveal the parent visual container panel now that items exist
        if (panel) panel.style.display = "block";
        if (badge) badge.innerText = window.maeWizardSessionItems.length;

        let htmlTable = `
            <table class="inventory-table" style="margin-top: 0; width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#f4f4f4;">
                        <th style="width:15%; font-size:0.85rem; padding:8px; background:#7f8c8d !important; color:white !important; text-align:center;">Logged Time</th>
                        <th style="width:55%; font-size:0.85rem; padding:8px; background:#7f8c8d !important; color:white !important;">Item Description / Size</th>
                        <th style="width:30%; font-size:0.85rem; padding:8px; background:#7f8c8d !important; color:white !important;">Assigned Location_ID</th>
                    </tr>
                </thead>
                <tbody>
        `;

        // Reverse map so the most recently saved tool floats right to the top of the clipboard list
        [...window.maeWizardSessionItems].reverse().forEach(item => {
            htmlTable += `
                <tr>
                    <td class="locked-cell" style="padding:8px; font-size:0.9rem; font-family:monospace; color:#666; text-align:center; vertical-align:middle;">${item.timestamp}</td>
                    <td class="locked-cell" style="padding:8px; font-size:0.9rem; vertical-align:middle;"><b>${item.description}</b></td>
                    <td class="locked-cell" style="padding:8px; font-size:0.9rem; color:var(--accent); font-weight:bold; vertical-align:middle;">${item.location}</td>
                </tr>
            `;
        });

        htmlTable += `</tbody></table>`;
        gridMount.innerHTML = htmlTable;
    },
    // 🌟 ADD NEW termination method to flush array trackers and clear out the intake view context
    finalizeWizardBatchSession() {
        const formZone = document.getElementById("central-form-render-zone");
        if (formZone) formZone.innerHTML = "";

        // Flush session array trackers out of active memory completely
        window.maeWizardSessionItems = [];
        
        // Reset global configuration table router states safely back to home landing page
        window.currentTable = "Master_Dashboard";
        
        alert("Batch Intake Session Successfully Sealed!\n\nAll items have been committed to your OneDrive data ledger partitions.");
        window.loadTableData("Master_Dashboard");
    },

    launchContextualFormFromCentral() {
        const tableSelect = document.getElementById("mae-central-table-selector");
        const targetTable = tableSelect ? tableSelect.value : "";

        if (!targetTable) {
            alert("Mandatory Selection Required:\n\nPlease choose an active inventory classification table target from the dropdown list first.");
            return;
        }

        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === targetTable);
        const formZone = document.getElementById("central-form-render-zone");
        formZone.innerHTML = ""; // Clean workspace slate

        // 🌟 MAE REGISTRATION LOCK FLAG 🌟
        // Instructs your global background listener that an intake process is active
        window.currentTable = "inventory_registration";

        // Trigger your proven, table-contextual form generator model
        this.renderEntryForm('add', targetTable, sheetConfig, async () => {
            const success = await window.submitNewRow(targetTable, sheetConfig);
            if (success) {
                formZone.innerHTML = ""; // Clear active form on successful commit
                window.currentTable = "Master_Dashboard"; // Revert routing lock safely
                alert("Central Entry Successfully Committed to OneDrive Ledger!");
            }
        });

        // Failsafe input cursor focus routing macro
        setTimeout(() => {
            const formCard = document.getElementById("entry-form");
            if (!formCard) return;

            const descriptiveInputField = formCard.querySelector("input[type='text']:not(#field-Tag_ID)");
            if (descriptiveInputField) {
                descriptiveInputField.focus();
                descriptiveInputField.style.backgroundColor = "#fffde7"; // Highlight active typing field yellow
            }
        }, 150);
    },
    // 🌟 ADD NEW function to clear and reset the onboarding wizard canvas completely
    resetCentralRegistrationWizard() {
        // 🌟 TRIP THE CIRCUIT BREAKER 🌟
        // Changing this value completely invalidates any background network loops currently running
        window.activeScanTransactionId = null;

        const tableSelect = document.getElementById("mae-central-table-selector");
        const tagInput = document.getElementById("field-Tag_ID");
        const feedback = document.getElementById("wizard-tag-feedback");
        const formZone = document.getElementById("central-form-render-zone");

        // 1. Clear text contents and reset visual tracking layouts
        if (formZone) formZone.innerHTML = "";
        if (feedback) feedback.innerHTML = "";
        
        if (tagInput) {
            tagInput.value = "";
            tagInput.disabled = false;
            tagInput.style.borderColor = "var(--border)";
            tagInput.style.backgroundColor = "#fffde7"; // Action yellow scan box hint
        }

        if (tableSelect) {
            tableSelect.value = "";
            tableSelect.disabled = false;
        }

        // 2. Clear any lingering mailbox variables inside global storage
        window.pendingScanValue = null;

        // 3. Re-orient focus instantly back to the scanner input field box
        setTimeout(() => {
            if (tagInput) tagInput.focus();
        }, 50);
        
        console.log("MAE Wizard System: Central registration canvas successfully sanitized and reset.");
    },

    //========  END THE CENTRAL INTAKE REGISTRATION PORTAL ==========

    
//====== Virtual Table View Renderer for UNTAGGED Audit with Inline Structural Layout CSS ======
renderUntaggedAuditGrid(auditData) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    title.innerText = "Audit: Items Awaiting Physical Tag Assignment";

    if (!auditData || auditData.length === 0) {
        container.innerHTML = `
            <div class="form-card" style="text-align:center; padding:40px; margin:20px;">
                <h3 style="color:#27ae60; margin-top:0;">✅ Tag Compliance Verified</h3>
                <p>Excellent discipline! Every single asset inside the workshop database ledger has a registered Tag_ID.</p>
                <button class="action-btn" onclick="window.loadTableData('Master_Dashboard')">Return to Dashboard</button>
            </div>`;
        
        const actionZone = document.getElementById("action-bar-zone");
        if (actionZone) actionZone.innerHTML = "";
        return;
    }

    // --- MAE ENHANCED MATRIX CONTAINER ---
    // This parent layout wrapper strictly partitions the stationary header from the scrolling data table underneath
    let html = `
        <div id="mae-audit-view-wrapper" style="display: flex; flex-direction: column; height: 100%; width: 100%; overflow: hidden;">
            
            <!-- STATIONARY CONTAINER ASSY BAR BLOCK -->
            <div id="mae-bulk-tagging-control-bar" style="display:none; flex-shrink: 0; margin-bottom:15px; width:100%; background:#f4f4f4; padding-bottom:5px; box-sizing:border-box;">
                <div class="form-card" style="border-left: 6px solid var(--accent); background: #ffffde; padding: 15px; display: flex; align-items: center; justify-content: space-between; gap: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); margin-bottom: 0;">
                    <div>
                        <h4 style="margin:0 0 5px 0; color:var(--accent); font-weight:800; text-transform:uppercase; font-size:0.95rem;">⚡ Bulk Container Assembly Active</h4>
                        <p style="margin:0; font-size:0.85rem; color:#444;">Scan or type a single barcode to group all <span id="mae-bulk-checked-count" style="font-weight:bold; color:var(--primary);">0</span> checked items into a shared space.</p>
                    </div>
                    <div style="display:flex; gap:10px; align-items:center;">
                        <input type="text" id="mae-bulk-container-input" placeholder="Scan shared label sticker here..." 
                               style="height:40px; width:280px; padding:0 12px; border:2px solid var(--accent); border-radius:4px; font-weight:bold; font-size:0.9rem; background:#ffffff;"
                               onchange="window.executeBulkContainerGroupingTransition(this.value)">
                        <button class="cancel-btn" style="height:40px; margin-left:5px; padding: 0 15px;" onclick="UI.clearBulkAuditSelection()">Clear Selection</button>
                    </div>
                </div>
            </div>

            <!-- INDEPENDENT SCROLLING VIEWPANEL MATRIX -->
            <div id="mae-audit-table-scroll-zone" style="flex: 1; overflow-y: auto; min-height: 0; background: #ffffff; border-radius: 4px;">
    `;

    const grouped = auditData.reduce((acc, item) => {
        if (!acc[item.category]) acc[item.category] = [];
        acc[item.category].push(item);
        return acc;
    }, {});

    html += `<table class="inventory-table" id="main-data-table" style="width:100%; border-collapse:collapse; min-width:100%; margin-top:0;">`;
    for (const [category, items] of Object.entries(grouped)) {
        html += `
            <thead>
                <tr>
                    <th colspan="3" style="background:#c0392b !important; color:white !important; padding:12px; position: sticky; top: 0; z-index: 99;">
                        ${category.toUpperCase()} (${items.length} Gaps)
                    </th>
                </tr>
                <tr>
                    <th style="width:8%; text-align:center; background: var(--primary) !important; color: white !important; position: sticky; top: 41px; z-index: 99;">Select</th>
                    <th style="width:52%; background: var(--primary) !important; color: white !important; position: sticky; top: 41px; z-index: 99;">Item Description Name</th>
                    <th style="width:40%; background: var(--primary) !important; color: white !important; position: sticky; top: 41px; z-index: 99;">Scan / Type Individual Tag_ID (UNIQUE)</th>
                </tr>
            </thead>
            <tbody>`;
        
        items.forEach(item => {
            const htmlRowId = `untagged-row-${item.rowIndex}`;
            html += `
                <tr id="${htmlRowId}" style="transition: opacity 0.4s ease, background-color 0.3s ease;">
                    <td style="text-align:center; vertical-align:middle; padding:10px;">
                        <input type="checkbox" class="mae-audit-bulk-checkbox" 
                               style="transform: scale(1.4); cursor:pointer; width:22px; height:22px;"
                               data-table="${item.tableName}" 
                               data-id="${item.mae_id}" 
                               data-row="${item.rowIndex}"
                               onchange="UI.evaluateAuditCheckboxStateChanges()">
                    </td>
                    <td class="locked-cell" style="padding: 12px 15px; vertical-align:middle;"><b>${item.itemName}</b></td>
                    <td style="padding: 8px 15px; vertical-align:middle;">
                        <input type="text" placeholder="Scan standalone label sticker..." class="mae-individual-scan-box"
                               style="width:100%; height:38px; background:#fffde7; border:2px solid var(--accent); padding:0 8px; font-weight:bold; box-sizing: border-box;"
                               onchange="window.handleAuditUpdate('${item.tableName}', '${item.mae_id}', this.value, '${htmlRowId}')">
                    </td>
                </tr>`;
        });
        html += `</tbody>`;
    }
    html += `</table>`;
    
    // Close scroll zone and parent wrapper divisions cleanly
    html += `
            </div> 
        </div>`;
    
    container.innerHTML = html;

    const actionZone = document.getElementById("action-bar-zone");
    if (actionZone) {
        actionZone.innerHTML = `
            <div class="command-bar" style="justify-content: center;">
                <button class="action-btn" onclick="window.loadTableData('Master_Dashboard')">← Return to Master Dashboard</button>
            </div>`;
    }
    
    window.currentTable = "untagged_audit_grid_view";
},

// ==========================================
// STATE OBSERVERS FOR THE BULK AUDIT GRID
// ==========================================
evaluateAuditCheckboxStateChanges() {
    const checkboxes = document.querySelectorAll('.mae-audit-bulk-checkbox:checked');
    const controlBar = document.getElementById('mae-bulk-tagging-control-bar');
    const countDisplay = document.getElementById('mae-bulk-checked-count');
    const individualInputs = document.querySelectorAll('.mae-individual-scan-box');

    if (checkboxes.length > 0) {
        // Reveal the bulk command header bar inline
        controlBar.style.display = "block";
        countDisplay.innerText = checkboxes.length;
        
        // Disable individual row inputs while checking boxes to prevent user confusion
        individualInputs.forEach(input => {
            input.disabled = true;
            input.style.opacity = "0.4";
            input.style.cursor = "not-allowed";
        });
    } else {
        this.clearBulkAuditSelection();
    }
},

clearBulkAuditSelection() {
    const checkboxes = document.querySelectorAll('.mae-audit-bulk-checkbox');
    const controlBar = document.getElementById('mae-bulk-tagging-control-bar');
    const individualInputs = document.querySelectorAll('.mae-individual-scan-box');
    const bulkInput = document.getElementById('mae-bulk-container-input');

    checkboxes.forEach(cb => cb.checked = false);
    if (controlBar) controlBar.style.display = "none";
    if (bulkInput) bulkInput.value = "";

    // Re-enable standalone inputs
    individualInputs.forEach(input => {
        input.disabled = false;
        input.style.opacity = "1";
        input.style.cursor = "text";
    });
},

// ==========================================
// STATE OBSERVERS FOR THE BULK AUDIT GRID
// ==========================================
evaluateAuditCheckboxStateChanges() {
    const checkboxes = document.querySelectorAll('.mae-audit-bulk-checkbox:checked');
    const controlBar = document.getElementById('mae-bulk-tagging-control-bar');
    const countDisplay = document.getElementById('mae-bulk-checked-count');
    const individualInputs = document.querySelectorAll('.mae-individual-scan-box');

    if (checkboxes.length > 0) {
        // 1. Reveal the bulk command header bar
        controlBar.style.display = "block";
        countDisplay.innerText = checkboxes.length;
        
        // 2. INDUSTRIAL SAFETY GUARD: Disable individual row inputs while checking boxes to prevent user confusion
        individualInputs.forEach(input => {
            input.disabled = true;
            input.style.opacity = "0.4";
            input.style.cursor = "not-allowed";
        });
    } else {
        // 3. Hide bar if selection returns to empty state
        this.clearBulkAuditSelection();
    }
},

clearBulkAuditSelection() {
    const checkboxes = document.querySelectorAll('.mae-audit-bulk-checkbox');
    const controlBar = document.getElementById('mae-bulk-tagging-control-bar');
    const individualInputs = document.querySelectorAll('.mae-individual-scan-box');
    const bulkInput = document.getElementById('mae-bulk-container-input');

    checkboxes.forEach(cb => cb.checked = false);
    if (controlBar) controlBar.style.display = "none";
    if (bulkInput) bulkInput.value = "";

    // Re-enable standalone inputs
    individualInputs.forEach(input => {
        input.disabled = false;
        input.style.opacity = "1";
        input.style.cursor = "text";
    });
},
//====== END  Virtual Table View Renderer for UNTAGGED Audit ========

//========== TAG MAINTENANCE ===================
//========  Render Tag Maintenance Wizard for Lost/Damaged Tags ========
renderTagMaintenanceWizard() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        
        title.innerText = "Maintenance: Decommission Lost or Damaged Tags";
        window.currentTable = "tag_maintenance"; // Establish explicit maintenance mode state context

        let html = `
            <div class="form-card" style="border-left:6px solid #8e44ad; padding:25px; background:#fff; margin-bottom:20px;">
                <h4 style="margin:0 0 10px 0; color:var(--primary); text-transform:uppercase;">🛠️ Isolate Broken Label Rows</h4>
                <p style="font-size:0.85rem; color:#666; margin:0 0 20px 0;">If a tag sticker fell off, is torn, or cannot be scanned, use either entry field below to pinpoint the asset in the ledger.</p>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; align-items:end;">
                    <!-- Option A: Keyword Text Lookup -->
                    <div class="input-group">
                        <label style="font-weight:bold; color:var(--primary); margin-bottom:5px;">A. Find by Keyword Description</label>
                        <input type="text" id="mae-maintenance-search-input" placeholder="Type name (e.g. bandsaw, bolt)..." style="height:45px; padding:0 12px; border:1px solid var(--border); border-radius:4px; font-size:1rem;">
                    </div>
                    
                    <!-- Option B: Location Dropdown Lookup -->
                    <div class="input-group">
                        <label style="font-weight:bold; color:var(--primary); margin-bottom:5px;">B. Find by Storage Spot Landmark</label>
                        <select id="mae-maintenance-location-selector" class="edit-dropdown" style="height:45px; font-size:1rem; border:1px solid var(--border);">
                            <option value="">-- Select Location_ID --</option>
                            ${window.maeLocations.map(loc => `<option value="${loc}">${loc}</option>`).join('')}
                        </select>
                    </div>
                </div>
                
                <div style="margin-top:20px; display:flex; gap:15px;">
                    <button class="action-btn" style="flex:1; background:var(--primary); height:45px; font-weight:bold;" onclick="UI.executeMaintenanceSearch('keyword')">🔍 Run Text Search Matrix</button>
                    <button class="action-btn" style="flex:1; background:#2980b9; height:45px; font-weight:bold;" onclick="UI.executeMaintenanceSearch('location')">📊 Filter by Location Spot</button>
                </div>
            </div>
            <div id="maintenance-results-mount-zone"></div>
        `;

        container.innerHTML = html;

        // Custom action command bar for navigation safety
        const actionZone = document.getElementById("action-bar-zone");
        if (actionZone) {
            actionZone.innerHTML = `
                <div class="command-bar" style="justify-content: center;">
                    <button class="action-btn" onclick="window.loadTableData('Master_Dashboard')">← Return to Dashboard</button>
                </div>`;
        }
    },
//====== END  Render Tag Maintenance Wizard for Lost/Damaged Tags ========

//========  Execute Search Logic for Tag Maintenance Wizard ========
async executeMaintenanceSearch(searchType) {
        const resultsZone = document.getElementById("maintenance-results-mount-zone");
        resultsZone.innerHTML = `<div class="loader">Sweeping records for matching rows...</div>`;

        let cleanQuery = "";
        let filterField = "Item_Description";

        if (searchType === 'keyword') {
            cleanQuery = document.getElementById("mae-maintenance-search-input").value.trim().toLowerCase();
            if (!cleanQuery) { alert("Please input a keyword descriptor parameter first."); resultsZone.innerHTML = ""; return; }
        } else {
            cleanQuery = document.getElementById("mae-maintenance-location-selector").value.trim().toUpperCase();
            if (!cleanQuery) { alert("Please select a physical storage spot first."); resultsZone.innerHTML = ""; return; }
            filterField = "Location_ID";
        }

        const tablesToScan = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
        let maintenanceMatches = [];

        try {
            for (const tableName of tablesToScan) {
                const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === tableName);
                const dataRows = await window.Dashboard.getFullTableData(tableName);
                if (!dataRows || dataRows.length === 0) continue;

                const targetColIdx = sheetConfig.columns.findIndex(c => c.header === filterField);
                const descColIdx = sheetConfig.columns.findIndex(c => c.header === "Item_Description");
                const tagColIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");

                dataRows.forEach(rowObj => {
                    const cells = (rowObj.values && Array.isArray(rowObj.values)) ? rowObj.values[0] : rowObj.values;
                    if (!cells) return;

                    let matchConfirmed = false;
                    if (searchType === 'keyword') {
                        // Scan descriptions for substring match
                        const itemText = String(cells[descColIdx] || "").toLowerCase();
                        if (itemText.includes(cleanQuery)) matchConfirmed = true;
                    } else {
                        // Strict check on physical landmark vector
                        const itemLoc = String(cells[targetColIdx] || "").toUpperCase();
                        if (itemLoc === cleanQuery) matchConfirmed = true;
                    }

                    if (matchConfirmed) {
                        maintenanceMatches.push({
                            category: sheetConfig.tabName,
                            itemDescription: cells[descColIdx] || "N/A",
                            currentTag: cells[tagColIdx] || "UNTAGGED",
                            tableName: tableName,
                            rowIndex: rowObj.index
                        });
                    }
                });
            }

            if (maintenanceMatches.length === 0) {
                resultsZone.innerHTML = `<p style="padding:20px; text-align:center; font-style:italic; border:1px dashed #ccc; background:#f9f9f9;">No active records were found matching the parameters.</p>`;
                return;
            }

            // Build an interactive layout row block with custom decommissioning parameters
            let htmlTable = `
                <table class="inventory-table" style="margin-top:15px;">
                    <thead>
                        <tr>
                            <th style="width:25%;">Sheet Classification</th>
                            <th style="width:40%;">Item Description</th>
                            <th style="width:15%;">Active Tag ID</th>
                            <th style="width:20%; text-align:center;">Maintenance Action</th>
                        </tr>
                    </thead>
                    <tbody>`;

            maintenanceMatches.forEach(row => {
                htmlTable += `
                    <tr>
                        <td class="locked-cell"><b>${row.category}</b></td>
                        <td class="locked-cell">${row.itemDescription}</td>
                        <td class="locked-cell" style="font-family:monospace; font-weight:bold;">${row.currentTag}</td>
                        <td style="text-align:center;">
                            ${row.currentTag === "UNTAGGED" ? 
                                `<span style="color:#7f8c8d; font-size:0.85rem; font-style:italic;">Already Untagged</span>` : 
                                `<button class="mini-btn" style="background:#c0392b; font-weight:bold;" onclick="UI.executeDirectTagWipe('${row.tableName}', ${row.rowIndex}, '${row.currentTag}')">⚠️ Decommission Tag</button>`
                            }
                        </td>
                    </tr>`;
            });

            htmlTable += `</tbody></table>`;
            resultsZone.innerHTML = htmlTable;

        } catch (err) {
            console.error("MAE Maintenance Core Sweep Crash:", err);
            resultsZone.innerHTML = `<p style="color:red; padding:20px;">Routine sweep was interrupted. Check connection bounds.</p>`;
        }
    },
//======  END Execute Search Logic for Tag Maintenance Wizard ========

// =======Direct Tag Decommissioning Handler
// const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(window.maeSystemConfig.spreadsheetName)}:/workbook/tables/${item.tableName}/rows/itemAt(index=${item.rowIndex})`;
//======

async executeDirectTagWipe(tableName, rowIndex, oldTagId) {
        if (oldTagId === "UNTAGGED") {
            alert("MAE System: This item is already marked as untagged.");
            return;
        }

        // ==========================================
        // 1. DUAL WORKFLOW SELECTION PROMPT GATES
        // ==========================================
        const promptMsg = `MAE INDUSTRIAL TAG MAINTENANCE SYSTEM:\n\n` +
                          `You are modifying the physical label configuration for Tag_ID [${oldTagId}].\n\n` +
                          `• Click OK to completely STRIP this tag. ALL items holding this identifier across ALL tables will reset to "UNTAGGED" status and move to the compliance audit queue.\n\n` +
                          `• Click CANCEL if you are holding a brand-new physical sticker and want to perform an INSTANT "HOT-SWAP" replacement right now.`;

        const choice = confirm(promptMsg);
        
        let newTagValue = "UNTAGGED";
        
        if (!choice) {
            // PATH B: The operator wants to perform an instantaneous "Hot-Swap" right now
            const reStickerInput = prompt(`RE-STICKERING WORKSPACE INTERFACE:\n\nPlease scan or type your NEW replacement label sticker value right now:`);
            
            if (!reStickerInput || reStickerInput.trim() === "") {
                console.log("MAE System: Re-stickering procedure canceled by user choice.");
                return;
            }
            
            // 🌟 MAE ENGINE REPAIR: Route the prompt string through your Labels module to strip out the web URLs 🌟
            newTagValue = window.Labels.extractCleanId(reStickerInput).toUpperCase();
            
            if (newTagValue === "UNTAGGED") {
                alert("CRITICAL REGULATION BLOCKED:\n\nYou cannot assign the absolute fallback string 'UNTAGGED' as a functional hardware sticker token.");
                return;
            }

            // 1B. ANTI-COLLISION DOUBLE CHECK: Prevent mapping an already active tag
            this.showLoading("Verifying new tag uniqueness across database ledger partitions...");
            const token = await window.getGraphToken();
            const priorityTables = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
            let isCollisionDetected = false;

            for (const table of priorityTables) {
                const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === table);
                const rowsData = await window.Dashboard.getFullTableData(table);
                if (!rowsData || rowsData.length === 0) continue;

                const tagColIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
                if (tagColIdx === -1) continue;

                const matchFound = rowsData.find(row => {
                    // MAE FIXED APPARATUS: Dig down safely to index 0 of Graph's double-nested array structure [[...]]
                    const cells = (row.values && Array.isArray(row.values[0])) ? row.values[0] : (Array.isArray(row.values) ? row.values : null);
                    return cells && String(cells[tagColIdx]).trim() === newTagValue;
                });

                if (matchFound) {
                    isCollisionDetected = true;
                    break;
                }
            }

            if (isCollisionDetected) {
                alert(`CRITICAL COLLISION ERROR:\n\nThe scanned tag [${newTagValue}] is ALREADY actively assigned to an asset row inside your database ledger.\n\nYou cannot cross-contaminate tracking tokens. Grab a completely fresh, unused sticker roll.`);
                this.renderTagMaintenanceWizard();
                return;
            }
        }

        // ==========================================
        // 2. TRANSACTION PREPARATION & BATCH ASSEMBLY
        // ==========================================
        this.showLoading(`Transmitting tracking data adjustments to OneDrive: [${oldTagId}] ➔ [${newTagValue}]...`);

        try {
            const priorityTables = ["Shop_Machinery", "Shop_Power_Tools", "Shop_Hand_Tools", "Shop_Consumables", "Resell_Inventory"];
            const token = await window.getGraphToken();
            let matchedRowsToUpdate = [];

            // Sweep tables to locate EVERY single row holding the old damaged identifier string
            for (const table of priorityTables) {
                const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === table);
                const rowsData = await window.Dashboard.getFullTableData(table);
                if (!rowsData || rowsData.length === 0) continue;

                const tagColIdx = sheetConfig.columns.findIndex(c => c.header === "Tag_ID");
                if (tagColIdx === -1) continue;

                rowsData.forEach(row => {
                    // 🌟 MAE FIXED APPARATUS: Explicitly unwrap Graph API's 2D double-nested array cell matrix container safely 🌟
                    const cells = (row.values && Array.isArray(row.values[0])) ? row.values[0] : (Array.isArray(row.values) ? row.values : null);
                    
                    if (cells && cells[tagColIdx] !== undefined && cells[tagColIdx] !== null) {
                        if (String(cells[tagColIdx]).trim() === oldTagId.toString().trim()) {
                            matchedRowsToUpdate.push({
                                tableName: table,
                                rowIndex: parseInt(row.index, 10),
                                config: sheetConfig
                            });
                        }
                    }
                });
            }

            console.log(`MAE Maintenance Engine: Located ${matchedRowsToUpdate.length} matching rows requiring transformation.`);

            if (matchedRowsToUpdate.length === 0) {
                alert(`MAE SYSTEM EXCEPTION:\n\nCould not identify any active rows mapping to Tag ID: [${oldTagId}] in the data structures. Verify that the file matches the current metadata sync indicators.`);
                this.renderTagMaintenanceWizard();
                return;
            }

            // ==========================================
            // 3. TRANSACTION EXECUTION (PRESERVES ITEM DATA)
            // ==========================================
            for (const item of matchedRowsToUpdate) {
                const tagIdIdx = item.config.columns.findIndex(c => c.header === "Tag_ID");
                const tagTypeIdx = item.config.columns.findIndex(c => c.header === "Tag_Type");

                // Sparse mapping: Array of nulls guarantees existing item specifications remain untouched
                const rowValues = new Array(item.config.columns.length).fill(null);
                rowValues[tagIdIdx] = newTagValue;
                
                if (tagTypeIdx !== -1) {
                    // Logic Guard: If hot-swapping a container, keep its status as MULTIPLE. If clearing it, drop back to safe UNIQUE initial default status.
                    rowValues[tagTypeIdx] = (newTagValue === "UNTAGGED") ? "UNIQUE" : "MULTIPLE";
                }

                const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(window.maeSystemConfig.spreadsheetName)}:/workbook/tables/${item.tableName}/rows/itemAt(index=${item.rowIndex})`;
                const response = await fetch(url, {
                    method: 'PATCH',
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ values: [rowValues] })
                });

                if (!response.ok) {
                    console.error(`MAE Fault Intercept: Network update lock failed on row index ${item.rowIndex} inside table ${item.tableName}`);
                }
                
                // 400ms Throttling protection delay to safeguard Microsoft Graph workbook concurrency locks
                await new Promise(r => setTimeout(r, 400));
            }

            //==alert(`System Integrity Verified!\n\nSuccessfully transformed ${matchedRowsToUpdate.length} database ledger row entries to hold Tag ID: [${newTagValue}]. All descriptive asset features remain preserved.`);
            //==this.renderTagMaintenanceWizard(); // Reload clean status wizard screen

            alert(`System Integrity Verified!\n\nSuccessfully transformed ${matchedRowsToUpdate.length} database ledger row entries to hold Tag ID: [${newTagValue}]. All descriptive asset features remain preserved.`);

        // 🌟 MAE ENGINE RUGGED FIXED APPARATUS: ABSOLUTE SYNC PIPELINE FLUSH 🌟
        // Instead of drawing the wizard context from local stale parameters,
        // we display a loading frame and force the global router to re-download 
        // your updated ground-truth ledger files directly from OneDrive.
        this.showLoading("Synchronizing local ledger partitions... please wait.");

        // 1200ms Industrial Settle Delay: Gives Microsoft's cloud workbook parameters
        // ample time to process, calculate, and write file metadata completely on their servers.
        await new Promise(r => setTimeout(r, 1200));

        // Reset the active state tracking variables
        window.currentTable = "Master_Dashboard";

        // Pull a clean download from OneDrive and safely bounce the view back to the main layout screen
        window.loadTableData("Master_Dashboard");
        } catch (err) {
            console.error("MAE Hot-Swap Transaction Sub-System Crash:", err);
            this.showError("Failed to safely complete tag re-homing updates. Check network links.");
        }
    }
//==== END Direct Tag Decommissioning Handler ========

};

window.UI = UI;


