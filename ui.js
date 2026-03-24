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
    renderMenu(activeWorksheets, onClickCallback) {
        const menu = document.getElementById("menu");
        menu.innerHTML = ""; // Clear existing
        
        activeWorksheets.forEach(sheet => {
            const li = document.createElement("li");
            const btn = document.createElement("button");
            btn.innerText = sheet.tabName;
            btn.className = "menu-btn";
            
            // When clicked, app.js logic will run
            btn.onclick = () => {
                // Set active visual state
                document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                onClickCallback(sheet.tableName);
            };

            li.appendChild(btn);
            menu.appendChild(li);
        });
    },

    // 3. TABLE RENDERING (The "Worker" logic refactored from app.js)
    // Practical: Uses the Config "Blueprint" to filter out hidden technical columns.
    // Rugged: Handles empty states and Microsoft Graph's row structure.
    renderTable(rows, tableName, sheetConfig) {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        
        if (!sheetConfig) {
            container.innerHTML = "Error: Worksheet configuration not found.";
            return;
        }

        title.innerText = `View: ${sheetConfig.tabName}`;

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

    setHealthStatus(isHealthy, firstTableName) {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");

        if (isHealthy) {
            title.innerText = "System Ready: Select a Category";
            container.innerHTML = `<p style="padding:20px;">Workbook verified. Use the sidebar to manage your workshop modules.</p>`;
        } else {
            title.innerText = "System Integrity Alert";
            container.innerHTML = `
                <div style="padding:20px; color: #c0392b;">
                    <p><strong>Warning:</strong> The spreadsheet structure has been modified outside the app.</p>
                    <p>Please ensure the table <b>${firstTableName}</b> exists in Excel.</p>
                </div>`;
        }
    },
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
printTable(tableName, sheetConfig) {
    // Target the main content area so the title stays aligned with the table
    const container = document.getElementById("app-content");
    if (!container) return;

    // Create the temporary print header
    const printHeader = document.createElement("div");
    printHeader.className = "print-only-title";
    
    // RUGGED: Simple, clear branding for the hardcopy
    printHeader.innerHTML = `
        <h1>MAE Workshop Inventory System: ${sheetConfig.tabName}</h1>
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
    const container = document.getElementById("table-container");
    const title = document.getElementById("current-view-title");
    
    // 1. Convert the raw Excel row array into a keyed object using Dashboard.parseSummary
    const dashboardData = Dashboard.parseSummary(row, config);

    title.innerText = "Workshop Master Dashboard";

    // 2. Inject the Responsive Grid
    container.innerHTML = `
        <div class="dashboard-grid">
            
            <!-- Snapshot A: Resell Inventory (Prioritizing Sales vs Investment) -->
            <div class="dash-card" onclick="loadTableData('Resell_Inventory')">
                <h4>Resell Inventory</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Actual Sales"])}</div>
                <p>Total Actual Sales</p>
                <small style="color: #7f8c8d;">Total Invested: ${formatCurrency(dashboardData["Total Resell Investment"])}</small>
            </div>

            <!-- Snapshot B: Total Asset Value -->
            <div class="dash-card" onclick="loadTableData('Shop_Machinery')">
                <h4>Total Shop Assets</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Shop Asset Value"])}</div>
                <p>Machinery, Tools & Supplies</p>
            </div>

            <!-- Snapshot C: Low Stock Alerts (Dynamic Alert Class) -->
            <div class="dash-card ${dashboardData["Low Stock Items Count"] > 0 ? 'alert' : ''}" 
                 onclick="loadTableData('Shop_Consumables')">
                <h4>Low Stock Alerts</h4>
                <div class="hero-num">${dashboardData["Low Stock Items Count"]}</div>
                <p>Items below reorder point</p>
            </div>

            <!-- Snapshot E: Monthly Overhead -->
            <div class="dash-card" onclick="loadTableData('Shop_Overhead')">
                <h4>Monthly Overhead</h4>
                <div class="hero-num">${formatCurrency(dashboardData["Total Monthly Overhead"])}</div>
                <p>Total Fixed Costs</p>
            </div>

            <!-- Snapshot F: Equipment Repairs (Warning Class) -->
            <div class="dash-card ${dashboardData["Equipment Needing Repair"] > 0 ? 'warning' : ''}" 
                 onclick="loadTableData('Shop_Machinery')">
                <h4>Repairs Needed</h4>
                <div class="hero-num">${dashboardData["Equipment Needing Repair"]}</div>
                <p>Out-of-Service Items</p>
            </div>

        </div>
    `;
}


//====END RENDER DASHBOARD =========

};


