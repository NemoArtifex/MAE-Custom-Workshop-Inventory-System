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
//=RENDER COMMAND BAR==========
    renderCommandBar(tableName) {
        const container = document.getElementById("action-bar-zone");

         // Define tables that should NOT show Add/Edit buttons
        const normalizedName = tableName.trim().toLowerCase();
        const dashboardTables = ["master_dashboard", "test_dashboard"];
        
        // Define buttons based on the current context
        let buttons = `
            <button class="action-btn" id="btn-print">Print Sheet</button>
            <button class="action-btn" id="btn-manual-print">Print Manual Log</button>
        `;

        // Only show "Add" and "Edit" if the current table is NOT in the dashboard list
        if (!dashboardTables.includes(normalizedName)) {
            buttons += `
                <button class="action-btn" id="btn-add">Add Item</button>
                <button class="action-btn" id="btn-edit">Edit Table</button>
                <button class="action-btn" id="btn-inventory-update">Quick Update</button>
            `;
        }

        container.innerHTML = `<div class="command-bar">${buttons}</div>`;
    },
// RENDER ENTRY FORM===============
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
            // If editing, pull the value from existingData using the column index
            const val = (isEdit && existingData) ? existingData[index] : "";

            formHtml += `<div class="input-group"><label>${col.header}</label>`;

            if (col.type === "dropdown") {
                formHtml += `
                    <select id="${fieldId}">
                        <option value="">-- Select ${col.header} --</option>
                        ${col.options.map(opt => 
                            `<option value="${opt}" ${opt == val ? 'selected' : ''}>${opt}</option>`
                        ).join('')}
                    </select>`;
            } else {
                let inputType = "text";
                if (col.type === "number") inputType = "number";
                if (col.type === "date") inputType = "date";

                const stepAttr = (col.type === "number" || (col.format && col.format.includes("$"))) ? 'step="0.01"' : '';

                formHtml += `
                    <input type="${inputType}" ${stepAttr} id="${fieldId}" value="${val}" placeholder="Enter ${col.header}...">`;
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

    // Attach the logic trigger
    document.getElementById("submit-form-btn").onclick = () => {
        onSaveCallback(rowIndex, existingData); 
    };
},
//=====PRINT TABLE ===========

printTable(tableName, sheetConfig) {
    const table = document.getElementById("main-data-table");
    if (!table) return;

    // Create a temporary print title for the top of the sheet
    const printTitle = document.createElement("div");
    printTitle.className = "print-only-title";
    printTitle.innerHTML = `<h1>MAE Workshop Inventory System: ${sheetConfig.tabName}</h1><hr>`;
    
    // Inject and trigger
    document.body.prepend(printTitle);
    window.print();
    printTitle.remove(); // Clean up after print dialog closes
}

};


