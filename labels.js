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
    initHIDScanner: function(onSuccessCallback) {
        let buffer = "";
        let scanTimer = null;
        let lastKeyTime = Date.now();
        
        window.isScannerActive = false;

        document.addEventListener('keydown', (e) => {
            // Ignore system navigation/modifier keys
            if (e.key.length !== 1 && e.key !== 'Enter') return;

            const currentTime = Date.now();
            const timeDelta = currentTime - lastKeyTime;
            lastKeyTime = currentTime;

            const focusedEl = document.activeElement;
            const isApprovedScannerField = focusedEl && (focusedEl.id === "field-Tag_ID" || focusedEl.id === "mae-bulk-container-input");

            // --- CONDITION 1: DATA PIPELINE ALREADY LOCKED BY SCANNER ---
            if (window.isScannerActive) {
                e.preventDefault(); // Lock out the DOM
                if (e.key === 'Enter') {
                    const cleanId = this.extractCleanId(buffer);
                    window.isScannerActive = false;
                    buffer = "";
                    onSuccessCallback(cleanId);
                } else if (e.key.length === 1) {
                    buffer += e.key;
                }
                return;
            }

            // --- CONDITION 2: ALLOW APPROVED SCANNER FIELD TO FLOW NATURALLY ---
            if (isApprovedScannerField) {
                // If the user clicked directly into the Tag_ID box, bypass the shield entirely
                if (e.key === 'Enter') {
                    e.preventDefault();
                    if (buffer.length > 0) {
                        const cleanId = this.extractCleanId(buffer);
                        onSuccessCallback(cleanId);
                        buffer = "";
                    } else {
                        // Extract from field value if typed manually
                        onSuccessCallback(this.extractCleanId(focusedEl.value));
                        focusedEl.value = "";
                    }
                } else if (e.key.length === 1) {
                    buffer += e.key;
                }
                return;
            }

            // --- CONDITION 3: UNAPPROVED FIELDS (THE SHIELD ACTIVE ZONE) ---
            // 1. Immediately cancel the default action so the character CANNOT print yet
            e.preventDefault();

            // 2. Accumulate character into our private evaluation string array
            if (e.key.length === 1) {
                buffer += e.key;
            }

            // 3. If a key arrives within 45ms of the previous key, confirm it is a hardware machine
            if (timeDelta < 45 && buffer.length > 1) {
                if (scanTimer) clearTimeout(scanTimer); // Kill human fallback timer
                window.isScannerActive = true;
                console.warn("MAE Absolute Shield: Hardware burst verified at capture layer. Freezing input.");
                
                if (focusedEl && (focusedEl.tagName === "INPUT" || focusedEl.tagName === "SELECT")) {
                    focusedEl.setAttribute("data-pre-scan-value", focusedEl.value || "");
                    focusedEl.disabled = true;
                    focusedEl.blur();
                }
                return;
            }

            // 4. Human Fallback Gate: If no fast keys hit, release the character to the input box after 40ms
            if (scanTimer) clearTimeout(scanTimer);
            
            if (e.key === 'Enter') {
                // Reset states if a human just hits enter on an unapproved field
                buffer = "";
                return;
            }

            const capturedChar = e.key;
            scanTimer = setTimeout(() => {
                // Timer expired without a collision! This is officially a human typing.
                if (!window.isScannerActive && focusedEl && document.activeElement === focusedEl) {
                    // Programmatically insert the character right where the user is typing
                    const start = focusedEl.selectionStart;
                    const end = focusedEl.selectionEnd;
                    const text = focusedEl.value;
                    focusedEl.value = text.slice(0, start) + capturedChar + text.slice(end);
                    focusedEl.setSelectionRange(start + 1, start + 1);
                    
                    // Trigger the snapshot input observer manually to keep app.js sync intact
                    focusedEl.dispatchEvent(new Event('input', { bubbles: true }));
                }
                buffer = ""; // Clear buffer for next key stroke
            }, 40);

        }, true); // TRUE ENFORCES THE CAPTURE GATE OVERRIDE
    },

    // 3. UI HELPER
    focusScanner: function() {
        console.log("MAE System: Awaiting Barcode...");
    }
};

window.Labels = Labels;


