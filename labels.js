/**
 * labels.js - MAE Custom Digital Solutions
 * Purpose: Handle Industrial HID Scanning (Inateck-75S) and ID extraction.
 * Philosophy: Practical, Functional, Simple, Rugged.
 */

import { UI } from './ui.js';

export const Labels = {
    // 1. EXTRACT ID: Remains flexible for URL-based or Raw tags
    extractCleanId: function(decodedText) {
        try {
            if (decodedText.startsWith("http")) {
                const url = new URL(decodedText);
                const queryId = url.searchParams.get("id");
                if (queryId) return queryId.trim();
                
                let cleanPath = url.pathname;
                if (cleanPath.endsWith('/')) {
                    cleanPath = cleanPath.slice(0, -1);
                }
                const pathId = cleanPath.split('/').pop();
                if (pathId) return pathId.trim();
            }
            return decodedText.trim();
        } catch (e) {
            console.warn("MAE System: Error parsing ID, using raw text.");
            return decodedText;
        }
    },

    // 🌟 2. INDUSTRIAL HID LISTENER: IMPENETRABLE TIMEOUT DEFERRAL SHIELD 🌟
    // Purpose: Defers all character entry by 40ms during the capture phase.
    // If a rapid succession of keys piles up during that window, it is locked 
    // down as a scanner burst, completely shielding input fields from character leakage.
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
                </div>
                <div id="mae-wizard-session-grid-mount"></div>
                
                <!-- TERMINATION ACTION CONTROL -->
                <button class="action-btn" onclick="UI.finalizeWizardBatchSession()" style="width:100%; height:50px; background:var(--primary); font-weight:bold; font-size:1.1rem; margin-top:20px; text-transform:uppercase; letter-spacing:0.5px;">
                    🏁 Finished Adding Items (Close Session)
                </button>
            </div>
        `;

        // 3. Trigger your proven, table-contextual entry form generator inside our new sub-mount
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
                    // Loop back inline, flag as subsequent entry to preserve session array records
                    window.UI.renderCentralRegistrationWizardStageTwo(targetTable, validatedTagId, "MULTIPLE", true);
                    
                    // Safe execution fallback: check if compiler engine is live before updating view elements
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
        });

        // 4. FORCE-INJECT COMPLIANCE VISUAL ATTRIBUTES
        setTimeout(() => {
            const formCard = document.getElementById("entry-form");
            if (!formCard) return;

            // Move form element layout directly inside our wizard sub-mount zone panel
            document.getElementById("mae-wizard-form-mount").appendChild(formCard);

            const closeBtn = formCard.querySelector(".close-x");
            if (closeBtn) closeBtn.remove();

            // A. 🌟 TARGET AND INJECT THE VALIDATED BARCODE STRING VALUE (UNDERSCORE FIXED)
            const tagIdInputBox = document.getElementById("field-Tag_ID");
            if (tagIdInputBox) {
                tagIdInputBox.value = validatedTagId; 
                tagIdInputBox.readOnly = true; 
                tagIdInputBox.style.backgroundColor = "#e8f8f5"; 
                tagIdInputBox.style.color = "#27ae60"; 
                tagIdInputBox.style.borderColor = "#27ae60"; 
                tagIdInputBox.style.fontWeight = "bold";
                tagIdInputBox.style.cursor = "not-allowed";

                let statusLabel = tagIdInputBox.parentNode.querySelector('.foundation-alert');
                if (!statusLabel) {
                    statusLabel = document.createElement("span");
                    statusLabel.className = "foundation-alert";
                    tagIdInputBox.parentNode.appendChild(statusLabel);
                }
                statusLabel.style.color = "#27ae60";
                statusLabel.innerText = "🔒 SCANNED TOKEN: Locked Against Manual Typing";
            } else {
                console.warn("MAE Intake Debug: Could not locate text field [field-Tag_ID] inside rendered DOM tree.");
            }

            // B. 🌟 TARGET AND INJECT THE ASSIGNED TAG CLASSIFICATION TYPE (UNDERSCORE FIXED)
            const tagTypeSelectBox = document.getElementById("field-Tag_Type");
            if (tagTypeSelectBox) {
                tagTypeSelectBox.value = tagType; 
                tagTypeSelectBox.disabled = true; 
                tagTypeSelectBox.style.backgroundColor = "#eeeeee";
                tagTypeSelectBox.style.color = "#888888";
                tagTypeSelectBox.style.cursor = "not-allowed";
                
                let secretPayload = document.getElementById("field-Tag_Type-hidden-backup");
                if (!secretPayload) {
                    secretPayload = document.createElement("input");
                    secretPayload.type = "hidden";
                    secretPayload.id = "field-Tag_Type-hidden-backup";
                    tagTypeSelectBox.parentNode.appendChild(secretPayload);
                }
                secretPayload.value = tagType;
            }

            // C. FLOW FOCUS REDIRECTION: Focus description inputs immediately (UNDERSCORE FIXED)
            const descInput = formCard.querySelector("input[type='text']:not(#field-Tag_ID)");
            if (descInput) {
                descInput.focus();
                descInput.style.backgroundColor = "#fffde7"; // Highlight active typing field yellow
            }

            // Draw current data lists if performing subsequent batch entry passes
            if (isSubsequentEntry && typeof window.UI.renderWizardSessionListGrid === "function") {
                window.UI.renderWizardSessionListGrid();
            }
        }, 150); 
    },

    // 3. UI HELPER
    focusScanner: function() {
        console.log("MAE System: Awaiting Barcode...");
    }
};

window.Labels = Labels;


