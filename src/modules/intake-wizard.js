/**
 * src/modules/intake-wizard.js - MAE Custom Digital Solutions
 * Capability: Central Item Intake Registration Portal (Stage 1 & Stage 2)
 * Responsibility: Renders onboarding forms, locks parameters contextually, and builds local session clipboards.
 * Philosophy: Practical, Functional, Simple, Rugged. Contains zero network fetch logic.
 */

export const IntakeWizard = {
    // 🌟 STAGE ONE: TOKEN IDENTIFICATION GATEWAY
    renderStageOne: function() {
        const container = document.getElementById("table-container");
        const title = document.getElementById("current-view-title");
        if (!container || !title) return;

        title.innerText = "Administrative: Centralized Item Intake Portal";
        window.currentTable = "inventory_registration"; // Lock the global router context state flag

        // Check for any pre-scanned tokens parked in our router mailbox
        const activeParkedToken = window.pendingScanValue || "";

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
                    <input type="text" id="field-Tag_ID" value="${activeParkedToken}" placeholder="Click here and scan physical label roll..." style="height:45px; border:2px solid var(--border); padding:0 12px; font-weight:bold; font-size:1rem; background: #fffde7;" autofocus>
                    <div id="wizard-tag-feedback" style="margin-top: 5px; font-size: 0.8rem; font-weight: bold;"></div>
                </div>
            </div>
            
            <div style="display: flex; gap: 15px;">
                <button class="action-btn" onclick="window.UI.processWizardStageOneScan()" style="background:var(--primary); height:45px; font-weight:bold; flex: 1;">⚡ Verify Scanned Tag</button>
                <button class="action-btn" onclick="window.UI.processWizardStageOneUntagged()" style="background:#7f8c8d; height:45px; font-weight:bold; flex: 1;">📦 Proceed as UNTAGGED</button>
                <button class="action-btn" onclick="window.IntakeWizard.resetWizardCanvas()" style="background:#c0392b; height:45px; font-weight:bold; flex: 1;">🔄 Clear / Reset Form</button>
            </div>
        </div>
        <div id="central-form-render-zone"></div>`;

        container.innerHTML = html;
        if (typeof window.UI.renderCommandBar === "function") window.UI.renderCommandBar("");

        // Setup real-time input highlights and listeners if values were auto-injected by router
        setTimeout(() => {
            const input = document.getElementById("field-Tag_ID");
            const tableSelect = document.getElementById("mae-central-table-selector");
            const feedback = document.getElementById("wizard-tag-feedback");

            if (input) {
                input.focus();
                if (input.value !== "") {
                    input.style.borderColor = "var(--accent)";
                    input.style.backgroundColor = "#e8f8f5"; // Operational mint shade highlight tint
                    if (feedback) {
                        feedback.style.color = "var(--primary)";
                        feedback.innerText = "⚡ Hardware Token Auto-Populated. Please select your Target Table dropdown to verify link uniqueness.";
                    }
                }
                input.onkeydown = (e) => {
                    if (e.key === 'Enter') {
                        e.preventDefault();
                        window.UI.processWizardStageOneScan();
                    }
                };
            }

            if (tableSelect && input) {
                tableSelect.onchange = () => {
                    if (tableSelect.value && input.value.trim() !== "") {
                        console.log(`MAE Intake System: Table destination choice [${tableSelect.value}] confirmed. Auto-triggering validation gate...`);
                        window.UI.processWizardStageOneScan();
                    }
                };
            }
        }, 100);
    },
    // 🌟 STAGE TWO: DISCIPLINARY INTER-LOCKING ENTRY FIELDS FORM
    renderStageTwo: function(targetTable, validatedTagId, tagType, isSubsequentEntry = false) {
        const formZone = document.getElementById("central-form-render-zone");
        if (!formZone) return;

        const sheetConfig = window.maeSystemConfig.worksheets.find(s => s.tableName === targetTable);
        console.log(`MAE Wizard Stage 2: Deploying form blueprint for sheet [${sheetConfig.tabName}] under profile [${window.maeWizardActiveCategory}].`);

        if (!isSubsequentEntry) {
            window.maeWizardSessionItems = []; // Initial startup flush for this batch running checklist
        }

        formZone.innerHTML = `
        <div id="mae-wizard-form-mount"></div>
        
        <!-- RUNNING LIVE SESSION CLIPBOARD PANEL -->
        <div id="mae-wizard-live-list-panel" style="margin-top: 30px; background: #ffffff; border: 1px solid var(--border); border-top: 4px solid var(--primary); padding: 20px; border-radius: 4px; display: ${window.maeWizardSessionItems.length > 0 ? 'block' : 'none'};">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 15px;">
                <h4 style="margin:0; color:var(--primary); text-transform:uppercase; font-weight:800; font-size:0.95rem;">📋 Items Registered in this Session</h4>
                <span style="background:var(--primary); color:white; padding:2px 8px; border-radius:10px; font-size:0.8rem; font-weight:bold;" id="mae-session-badge-count">${window.maeWizardSessionItems.length}</span>
            </div>
            <div id="mae-wizard-session-grid-mount"></div>
            
            <button class="action-btn" onclick="window.UI.finalizeWizardBatchSession()" style="width:100%; height:50px; background:var(--primary); font-weight:bold; font-size:1.1rem; margin-top:20px; text-transform:uppercase; letter-spacing:0.5px;">
                🏁 Finished Adding Items (Close Session)
            </button>
        </div>`;

        // 1. INJECT AND MAP THE FORM GRID VIEWS VIA THE BASE UI LAYER
        window.UI.renderEntryForm('add', targetTable, sheetConfig, async () => {
            // Re-enforce hardware definitions into DOM nodes right before harvesting values
            const tagField = document.getElementById("field-Tag_ID");
            const typeField = document.getElementById("field-Tag_Type");
            const categoryField = document.getElementById("field-Item_Category");
            const locationField = document.getElementById("field-Location_ID");

            if (tagField) tagField.value = validatedTagId;
            if (typeField) typeField.value = tagType;

            if (categoryField && window.maeWizardActiveCategory) {
                categoryField.disabled = false;
                categoryField.value = window.maeWizardActiveCategory;
            }
            if (locationField && window.maeWizardActiveCategory === "By_Location") {
                locationField.disabled = false;
            }

            // Capture strings for our local running checklist array BEFORE submission clears inputs
            const descCol = sheetConfig.columns.find(c => c.header.includes("Description") || c.header.includes("Name"));
            const descFieldId = `field-${descCol.header.replace(/\s+/g, '')}`;
            const enteredDescription = document.getElementById(descFieldId)?.value || "N/A";
            const enteredLocation = document.getElementById(`field-Location_ID`)?.value || "TBD";
            // Submit row asynchronously up to OneDrive data partitions
            const success = await window.submitNewRow(targetTable, sheetConfig);
            if (success) {
                window.maeWizardSessionItems.push({
                    description: enteredDescription,
                    location: enteredLocation,
                    timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
                });

                if (typeof window.warmInventoryCache === "function") {
                    await window.warmInventoryCache(); // Force background memory partitions to re-warm instantly
                }

                if (tagType === "MULTIPLE") {
                    // Continuous batch workflow: reload form keeping cluster variables intact
                    this.renderStageTwo(targetTable, validatedTagId, "MULTIPLE", true);
                    this.renderSessionGridList();
                } else {
                    // Unique item processing sequence sealed: clear out wizard parameters completely
                    window.maeWizardActiveCategory = null;
                    window.maeWizardSessionItems = [];
                    formZone.innerHTML = "";
                    window.loadTableData("Master_Dashboard");
                }
            }
        }, null, null, { tagId: validatedTagId, tagType: tagType });

        // 2. --- 🌟 DISCIPLINARY INTERLOCK ENFORCEMENTS OVER INPUT FIELD PARAMETERS 🌟 ---
        setTimeout(() => {
            const formCard = document.getElementById("entry-form");
            if (!formCard) return;

            // Restructure layouts: append form block inside our dedicated wizard render zone mount point
            document.getElementById("mae-wizard-form-mount").appendChild(formCard);
            const closeBtn = formCard.querySelector(".close-x");
            if (closeBtn) closeBtn.remove(); // Strip out exit button so users use the workflow wizard buttons

            const tagInputNode = document.getElementById("field-Tag_ID");
            const typeInputNode = document.getElementById("field-Tag_Type");
            const categoryInputNode = document.getElementById("field-Item_Category");
            const locationInputNode = document.getElementById("field-Location_ID");

            if (tagInputNode) {
                tagInputNode.value = validatedTagId;
                tagInputNode.readOnly = true;
                tagInputNode.style.cssText = "background-color:#e8f8f5; color:#27ae60; border:2px solid #27ae60; font-weight:bold; cursor:not-allowed;";
                
                let labelBadge = tagInputNode.parentNode.querySelector('.foundation-alert');
                if (!labelBadge) {
                    labelBadge = document.createElement("span");
                    labelBadge.className = "foundation-alert";
                    tagInputNode.parentNode.appendChild(labelBadge);
                }
                labelBadge.style.color = "#27ae60";
                labelBadge.innerText = "🔒 SCANNED TOKEN: Locked Against Manual Typing";
            }
             if (typeInputNode) {
                typeInputNode.value = tagType;
                typeInputNode.disabled = true;
                typeInputNode.style.cssText = "background-color:#eeeeee; color:#888888; cursor:not-allowed;";
            }

            // Enforce GEOGRAPHIC CONTAINER DISCIPLINE constraints strictly
            if (window.maeWizardActiveCategory === "By_Location") {
                if (categoryInputNode) {
                    categoryInputNode.value = "By_Location";
                    categoryInputNode.disabled = true;
                    categoryInputNode.style.cssText = "background-color:#eeeeee; color:#888888;";
                }
                if (locationInputNode) {
                    const hasPriorSessionEntries = window.maeWizardSessionItems && window.maeWizardSessionItems.length > 0;
                    if (hasPriorSessionEntries) {
                        // Force sub-items in this continuous loop to permanently inherit the location of your first entry!
                        const firstItemStorageSpotLoc = window.maeWizardSessionItems[0].location;
                        locationInputNode.value = firstItemStorageSpotLoc;
                        locationInputNode.disabled = true;
                        locationInputNode.style.cssText = "background-color:#eeeeee; color:#888888; cursor:not-allowed;";

                        const alertTextBadge = document.createElement("span");
                        alertTextBadge.style.cssText = "color:#c0392b; font-weight:bold; font-size:0.8rem; display:block; margin-top:5px;";
                        alertTextBadge.innerText = "🔒 GEOGRAPHIC CONTAINER DISCIPLINE: Locked to Shared Storage Spot";
                        locationInputNode.parentNode.appendChild(alertTextBadge);
                    }
                }
            } else if (window.maeWizardActiveCategory === "By_Topic") {
                if (categoryInputNode) {
                    categoryInputNode.value = "By_Topic";
                    categoryInputNode.disabled = true;
                    categoryInputNode.style.cssText = "background-color:#eeeeee; color:#888888;";
                }
            }

            // Autofocus text input description parameters instantly for typing speed
            const descriptiveTextInputBox = formCard.querySelector("input[type='text']:not(#field-Tag_ID)");
            if (descriptiveTextInputBox) {
                descriptiveTextInputBox.focus();
                descriptiveTextInputBox.style.backgroundColor = "#fffde7"; // Action yellow
            }

            if (isSubsequentEntry) {
                this.renderSessionGridList();
            }
        }, 150);
    },

    // 🌟 RENDER LIVE SESSION CHECKLIST SUB-GRID ROWS
    renderSessionGridList: function() {
        const gridMount = document.getElementById("mae-wizard-session-grid-mount");
        const panel = document.getElementById("mae-wizard-live-list-panel");
        const badge = document.getElementById("mae-session-badge-count");
        if (!gridMount || !window.maeWizardSessionItems || window.maeWizardSessionItems.length === 0) return;

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
            <tbody>`;

        // Reverse checklist records matrix array so the newest item always sits right on top
        [...window.maeWizardSessionItems].reverse().forEach(item => {
            htmlTable += `
            <tr>
                <td class="locked-cell" style="padding:8px; font-size:0.9rem; font-family:monospace; color:#666; text-align:center; vertical-align:middle;">${item.timestamp}</td>
                <td class="locked-cell" style="padding:8px; font-size:0.9rem; vertical-align:middle;"><b>${item.description}</b></td>
                <td class="locked-cell" style="padding:8px; font-size:0.9rem; color:var(--accent); font-weight:bold; vertical-align:middle;">${item.location}</td>
            </tr>`;
        });

        htmlTable += `</tbody></table>`;
        gridMount.innerHTML = htmlTable;
    },

    // 🌟 CLEAR AND RESET THE ONBOARDING PORTAL FRAME
    resetWizardCanvas: function() {
        window.activeScanTransactionId = null; // Kill lagging async loop checks instantly via the circuit breaker
        window.pendingScanValue = null;
        window.maeWizardActiveCategory = null;

        const tableSelect = document.getElementById("mae-central-table-selector");
        const tagInput = document.getElementById("field-Tag_ID");
        const feedback = document.getElementById("wizard-tag-feedback");
        const formZone = document.getElementById("central-form-render-zone");

        if (formZone) formZone.innerHTML = "";
        if (feedback) feedback.innerHTML = "";
        if (tableSelect) {
            tableSelect.value = "";
            tableSelect.disabled = false;
        }
        if (tagInput) {
            tagInput.value = "";
            tagInput.disabled = false;
            tagInput.style.borderColor = "var(--border)";
            tagInput.style.backgroundColor = "#fffde7"; // Clear yellow input hint focus
        }
        setTimeout(() => { if (tagInput) tagInput.focus(); }, 50);
        console.log("MAE Wizard Module: Intake portal canvas successfully sanitized and recycled.");
    }
};

// Bind execution endpoints directly onto the unified window context wrapper axis
window.IntakeWizard = IntakeWizard;
window.UI.renderCentralRegistrationWizard = IntakeWizard.renderStageOne;
window.UI.renderWizardSessionListGrid = IntakeWizard.renderSessionGridList;
window.renderCentralRegistrationWizardStageTwo = IntakeWizard.renderStageTwo;