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
import { excelSerialToDate} from './app.js';

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
                    const isEditable = !colDef.locked && colDef.type !== 'formula';

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

    // --- RULE 1: EDIT MODE / QUICK UPDATE GUARDRAIL ---
    if (window.isEditing || isQuickUpdating) {
        let buttons = "";
        
        // 🌟 MIGRATION: Print Manual Log ONLY appears here during Quick Update
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

    // --- RULE 2: VIRTUAL AUDIT VIEW ---
    if (normalizedName === "location_audit") {
        container.innerHTML = `<div class="command-bar">
            <button class="action-btn" onclick="loadTableData('Location')">← Back to Map</button>
            <button class="action-btn" id="btn-print-audit">Print TBD Audit</button>
        </div>`;
        return;
    }

    // --- RULE 3: STANDARD INVENTORY TABLES ---
    if (!sheetConfig) return;
    const isDashboard = normalizedName.includes("dashboard");
    const hasManualField = sheetConfig.columns.some(col => 
        ["Quantity", "Current Stock", "Stock_Count"].includes(col.header)
    );

    let buttons = "";
    if (normalizedName === "location") {
        buttons = `
            <button class="action-btn" onclick="UI.manageLocationMap()">Manage Shop Location Map</button>
            <button class="action-btn" onclick="runLocationAudit()">Audit of TBD Locations</button>
            <button class="action-btn" id="btn-print">Print Location Map</button>
        `;
    } 
    else if (!isDashboard) {
        buttons = `
            <button class="action-btn" id="btn-print">🖨️ Print Sheet</button>
            ${hasManualField ? `<button class="action-btn" id="btn-manual-print">📋 Print Manual Log</button>` : ''}
            <button class="action-btn" id="btn-add">➕ Add Item</button>
            <button class="action-btn" id="btn-edit">✏️ Edit Table</button>
            ${hasManualField ? `<button class="action-btn" id="btn-inventory-update" style="background:#e67e22;">⚡ Quick Update</button>` : ''}
            
            /* 🌟 MAE ENGINE UPGRADE: CONTEXTUAL PIVOT CONTROLLER 🌟 */
            ${normalizedName === "resell_inventory" ? `<button class="action-btn" id="btn-resell-status-pivot" style="background:#8e44ad;">📊 Sort By Status</button>` : ''}
        `;
    }
    else {
        buttons = `<button class="action-btn" id="btn-print">Print Dashboard</button>`;
    }

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

        if (col.hidden || col.type === "formula") {
            formHtml += `<input type="hidden" id="${fieldId}" value="${val}">`;
        } 
        else {
            formHtml += `<div class="input-group"><label>${col.header}</label>`;

            // 1. BOOLEAN
            if (col.type === "boolean") {
                const isChecked = val.toString().toUpperCase() === "TRUE";
                formHtml += `<input type="checkbox" id="${fieldId}" ${isChecked ? 'checked' : ''} class="mae-checkbox">`;
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
                formHtml += `<input type="text" id="${fieldId}" value="${val}" placeholder="Enter ${col.header}..." ${isTag ? 'autofocus' : ''}>`;
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

    // RUGGED SCANNER INJECTION: Auto-fills Tag_ID if a scan is pending
    if (window.pendingScanValue) {
        const tagInput = document.getElementById("field-Tag_ID") || document.getElementById("field-mae_id");
        if (tagInput) {
            tagInput.value = window.pendingScanValue;
            tagInput.style.backgroundColor = "#fffde7"; 
            tagInput.style.border = "2px solid var(--accent)";
            console.log("MAE System: Scan auto-filled into form.");
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

//========Virtual Search Hub for Multiple ITems on a Single Tag 
renderVirtualSearchHub(auditData) {
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    
    title.innerText = "Search Results: Multiple Items on Scanned Tag";

    // 1. Group the data by category for visual organization on the shop floor
    const grouped = auditData.reduce((acc, item) => {
        if (!acc[item.category]) acc[item.category] = [];
        acc[item.category].push(item);
        return acc;
    }, {});

    // 2. Build the grouped table
    let html = `<table class="inventory-table" id="main-data-table">`;
    
    for (const [category, items] of Object.entries(grouped)) {
        // Render Category Header and Sub-Headers
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
                    <button class="action-btn" 
                            style="padding: 5px 12px; font-size: 0.85rem; background: var(--accent);" 
                            onclick="loadTableData('${item.tableName}')">
                        ✏️ View / Edit in Table
                    </button>
                </td>
            </tr>`;
        });
        html += `</tbody>`;
    }
    
    html += `</table>`;
    
    // Add a back button to the bottom so they can clear the search
    html += `
        <div class="form-actions" style="margin-top: 20px; text-align: center;">
            <button class="cancel-btn" style="width:50%;" onclick="loadTableData('Master_Dashboard')">
                ← Back to Dashboard
            </button>
        </div>`;
        
    container.innerHTML = html;

    // Render a localized back-to-dashboard command bar
    const actionZone = document.getElementById("action-bar-zone");
    if (actionZone) {
        actionZone.innerHTML = `
            <div class="command-bar" style="justify-content: center;">
                <button class="action-btn" onclick="loadTableData('Master_Dashboard')">← Return to Dashboard</button>
            </div>`;
    }
},
//====== END Virtual Search Hub for Multiple ITems on a Single Tag 

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

    UI.showLoading("Isolating records from local data cache buffer...");

    try {
        // Pull fresh dataset straight from your decoupled ledger using your dashboard core tool
        const rawRows = await window.Dashboard.getFullTableData("Resell_Inventory");
        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === "Resell_Inventory");
        const statusIdx = sheetConfig.columns.findIndex(c => c.header === "Current Status");

        // Step 1: Filter the rows based on the dropdown selection
        const matchingRows = rawRows.filter(rowObj => {
            const cells = (rowObj.values && Array.isArray(rowObj.values[0])) ? rowObj.values[0] : rowObj.values;
            return String(cells[statusIdx]).trim() === selectedStatus;
        });

        // Save this specific subset to global memory so the print subsystem can see it seamlessly
        this.activeStatusPivotData = matchingRows;
        this.activeStatusPivotLabel = selectedStatus;

        if (matchingRows.length === 0) {
            tableMount.innerHTML = `<p style="padding:20px; text-align:center; font-style:italic; border:1px dashed #ccc; background:#f9f9f9;">Zero records currently match the status criteria: "${selectedStatus}".</p>`;
            UI.renderCommandBar(""); // Reset command bar layout
            return;
        }

        // Step 2: Normalize dates using your Data Flattening Bridge logic
        const flattenedSubset = matchingRows.map(rowObj => {
            const cells = (rowObj.values && Array.isArray(rowObj.values[0])) ? rowObj.values[0] : rowObj.values;
            const cleanValues = cells.map((val, idx) => {
                const colDef = sheetConfig.columns[idx];
                return (colDef && colDef.type === 'date') ? excelSerialToDate(val) : val;
            });
            return { ...rowObj, values: cleanValues };
        });

        // 🌟 Step 3: Run your Default Location_ID sorting engine over the filtered rows 🌟
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
                if (colDef.format && colDef.format.includes("$") && displayValue !== "") displayValue = UI.formatCurrency(displayValue);
                htmlTable += `<td class="locked-cell">${displayValue}</td>`;
            });
            htmlTable += `</tr>`;
        });
        htmlTable += `</tbody></table>`;
        tableMount.innerHTML = htmlTable;

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
    const tableElement = document.getElementById("main-data-table");
    if (!tableElement || !this.activeStatusPivotLabel) return;

    const customPrintTitle = `Resell Item with Current Status: ${this.activeStatusPivotLabel}`;
    
    // Leverage your clean detached tab document printing system context safely
    const printWindow = window.open('', '_blank');
    const styleLinks = Array.from(document.querySelectorAll('link[rel="stylesheet"]'))
        .map(link => `<link rel="stylesheet" href="${link.href}">`).join('');

    printWindow.document.write(`
        <html>
        <head>
            <title>${customPrintTitle}</title>
            ${styleLinks}
            <style>
                body { background: white !important; color: black !important; padding: 20px; font-family: sans-serif; }
                .edit-only-cell, button, .form-card { display: none !important; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td { border: 1px solid black !important; padding: 8px !important; text-align: left; }
                th { background-color: #eee !important; color: black !important; font-weight: bold; text-transform: uppercase; }
            </style>
        </head>
        <body>
            <h2 style="margin-bottom:5px; text-transform:uppercase; letter-spacing:1px; border-bottom:2px solid black; padding-bottom:5px;">MAE Workshop Inventory System</h2>
            <h4 style="margin-top:5px; color:#333;">${customPrintTitle}</h4>
            <div>${tableElement.outerHTML}</div>
            <script>
                window.onload = function() { window.print(); window.close(); };
            </script>
        </body>
        </html>
    `);
    printWindow.document.close();
}
//======  END Compiling and Printing the Filtered Subset (Resell Inventory)

};

window.UI = UI;

