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
        
        // Add "Delete" Header (Hidden by default via .edit-only-cell CSS)
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
            rows.forEach((row, rowIndex) => {
                html += `<tr data-row-index="${rowIndex}">`;

                // Add Delete Icon Cell
                html += `<td class="edit-only-cell">
                            <button class="delete-row-btn" onclick="requestDelete(${rowIndex})">🗑️</button>
                         </td>`;
                
                // Extract cell data (handles different Graph API response formats)
                const allCells = Array.isArray(row.values[0]) ? row.values[0] : row.values; 

                visibleIndices.forEach(idx => {
                    const colDef = sheetConfig.columns[idx];
                    const isEditable = !colDef.locked && colDef.type !== 'formula';
                    const isQuantity = colDef.header === "Quantity" || colDef.header === "Current Stock";
                    
                    // Get raw value
                    let displayValue = allCells[idx] ?? '';

                    // APPLY FORMATTING: If config shows currency format, use helper
                    if (colDef.format && colDef.format.includes("$")) {
                        displayValue = formatCurrency(displayValue);
                    }

                    // 3. Build the Cell
                    // We add 'col-type-qty' as a class to help the app.js Arrow Key logic
                    html += `<td 
                            class="${isEditable ? 'editable-cell' : 'locked-cell'} ${isQuantity ? 'col-type-qty' : ''}" 
                            data-col-index="${idx}">${displayValue}</td>`;
                });
                html += `</tr>`;
            });
        } else {
            // Span across all visible columns + the hidden delete column
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

    document.getElementById("submit-form-btn").onclick = () => {
        onSaveCallback(rowIndex, existingData); 
    };
},
//======= END RENDER ENTRY FORM ============

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
    window.print();
    
    // 2. Clean up
    printHeader.remove();
    table.classList.remove("manual-log-mode");
}

};


