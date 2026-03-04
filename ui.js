/**
 * ui.js - MAE Custom Digital Solutions
 * Purpose: Handle all DOM manipulation and visual states.
 * Philosophy: Practical, Functional, Simple, Rugged.
 */

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
        
        title.innerText = `View: ${sheetConfig.tabName}`;

        if (!sheetConfig) {
            container.innerHTML = "Error: Worksheet configuration not found.";
            return;
        }

        // Identify which column indices are NOT hidden
        const visibleIndices = [];
        let html = `<table class="inventory-table"><thead><tr>`;
        
        sheetConfig.columns.forEach((col, index) => {
            if (col.hidden !== true) { 
                html += `<th>${col.header}</th>`;
                visibleIndices.push(index);
            }
        });
        html += `</tr></thead><tbody>`;

        // Render Rows using ONLY those visible indices
        if (rows && rows.length > 0) {
            rows.forEach((row) => {
                html += `<tr>`;
                
                // Microsoft Graph /rows returns values as an array [cell0, cell1...]
                // Note: Depending on your API call, it might be row.values[0] or just row.values
                const allCells = Array.isArray(row.values[0]) ? row.values[0] : row.values; 

                visibleIndices.forEach(idx => {
                    const value = allCells[idx];
                    html += `<td>${value !== null && value !== undefined ? value : ''}</td>`;
                });
                html += `</tr>`;
            });
        } else {
            html += `<tr><td colspan="${visibleIndices.length}" style="text-align:center; padding:20px;">No records found.</td></tr>`;
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

    renderCommandBar(tableName) {
        const container = document.getElementById("action-bar-zone");

         // Define tables that should NOT show Add/Edit buttons
        const dashboardTables = ["Master_Dashboard", "TEST_Dashboard"];
        
        // Define buttons based on the current context
        let buttons = `
            <button class="action-btn" id="btn-print">Print Sheet</button>
            <button class="action-btn" id="btn-manual-print">Print Manual Log</button>
        `;

        // Only show "Add" and "Edit" if the current table is NOT in the dashboard list
        if (!dashboardTables.includes(tableName)) {
            buttons += `
                <button class="action-btn" id="btn-add">Add Item</button>
                <button class="action-btn" id="btn-edit">Edit Table</button>
                <button class="action-btn" id="btn-inventory-update">Quick Update</button>
            `;
        }

        container.innerHTML = `<div class="command-bar">${buttons}</div>`;
    },

    // ui.js inside the export const UI = { ... }

renderAddForm(tableName, sheetConfig, onSaveCallback) {
    const container = document.getElementById("table-container");
    
    // 1. Create the Form Container
    let formHtml = `
        <div class="form-card" id="add-entry-form">
            <div class="form-header">
                <h3>Add New Entry: ${sheetConfig.tabName}</h3>
                <button class="close-x" onclick="document.getElementById('add-entry-form').remove()">×</button>
            </div>
            <div class="form-grid">`;

    // 2. Loop through columns and build inputs/dropdowns
    sheetConfig.columns.forEach(col => {
        // RUGGED: Skip ID, Hidden, and Formulas
        if (!col.hidden && col.type !== "formula") {
            formHtml += `<div class="input-group"><label>${col.header}</label>`;

            // HANDLE DROPDOWNS
            if (col.type === "dropdown") {
                formHtml += `
                    <select id="field-${col.header.replace(/\s+/g, '')}">
                        <option value="">-- Select ${col.header} --</option>
                        ${col.options.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                    </select>`;
            } 
            // HANDLE STANDARD INPUTS (Date, Number, Text)
            else {
                let inputType = "text";
                if (col.type === "number") inputType = "number";
                if (col.type === "date") inputType = "date";

                const stepAttr = (col.type === "number" || (col.format && col.format.includes("$"))) ? 'step="0.01"' : '';

                formHtml += `
                    <input type="${inputType}" 
                           ${stepAttr}
                           id="field-${col.header.replace(/\s+/g, '')}" 
                           placeholder="Enter ${col.header}...">`;
            }

            formHtml += `</div>`;
        }
    });

    formHtml += `</div>
        <div class="form-actions">
            <button class="save-btn" id="submit-new-row">Save to OneDrive</button>
            <button class="cancel-btn" onclick="document.getElementById('add-entry-form').remove()">Cancel</button>
        </div>
    </div>`;

    // 3. Inject the form (using 'afterbegin' so it appears at the top of the table container)
    // or 'beforebegin' if you want it outside the table box entirely.
    container.insertAdjacentHTML('beforebegin', formHtml);

    // 4. Attach the Logic Trigger
    document.getElementById("submit-new-row").onclick = () => {
        // Validation: Simple check to ensure form isn't empty
        onSaveCallback();
    };
}


};


