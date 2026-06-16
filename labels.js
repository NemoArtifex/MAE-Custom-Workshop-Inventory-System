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

    // 🌟 2. INDUSTRIAL HID LISTENER: REPLACED WITH CAPTURE-PHASE SHIELDING 🌟
    // Purpose: Intercepts high-speed Inateck keystroke bursts during the capture phase,
    // freezing active input cells before characters can leak into currency or date fields.
    initHIDScanner: function(onSuccessCallback) {
        let buffer = "";
        let lastKeyTime = Date.now();
        let fastKeyCount = 0; // 🌟 Track consecutive fast keystrokes
        
        // Global tracking flag tells app.js what device sent the input stream
        window.isScannerActive = false;

        document.addEventListener('keydown', (e) => {
            const currentTime = Date.now();
            const timeDelta = currentTime - lastKeyTime;
            lastKeyTime = currentTime;

            // Ignore modifier keys like Shift or Control that don't add characters
            if (e.key.length === 1) {
                if (timeDelta < 30) {
                    fastKeyCount++;
                } else {
                    fastKeyCount = 0; // Reset if the user pauses normally
                }
            }

            // 🌟 RIGID BARRIER: Only lock if at least 3 consecutive characters arrive at hardware speeds
            if (fastKeyCount >= 3 && buffer.length >= 2) {
                if (!window.isScannerActive) {
                    window.isScannerActive = true;
                    console.warn("MAE Shield: Hardware scanner burst confirmed. Activating input lock.");
                    
                    const focusedEl = document.activeElement;
                    if (focusedEl && (focusedEl.tagName === "INPUT" || focusedEl.tagName === "SELECT")) {
                        if (focusedEl.id !== "field-Tag_ID" && focusedEl.id !== "mae-bulk-container-input") {
                            focusedEl.setAttribute("data-pre-scan-value", focusedEl.value || "");
                            focusedEl.disabled = true;
                            focusedEl.blur();
                        }
                    }
                }
            }

            // If the system flags an active scan, stop the browser from printing characters
            if (window.isScannerActive && e.key !== 'Enter') {
                if (e.key.length === 1) { buffer += e.key; }
                e.preventDefault(); 
                return;
            }

            // Standard human typing recovery path (Resets buffer if a long pause occurs)
            if (timeDelta > 150) {
                buffer = "";
                fastKeyCount = 0;
                window.isScannerActive = false;
            }

            // Processing the complete hardware barcode array
            if (e.key === 'Enter') {
                if (buffer.length > 2) {
                    const cleanId = this.extractCleanId(buffer);
                    window.isScannerActive = false;
                    fastKeyCount = 0;
                    onSuccessCallback(cleanId);
                    buffer = "";
                    e.preventDefault();
                }
            } else {
                if (e.key.length === 1) {
                    buffer += e.key;
                }
            }
        }, true); // 🌟 TRUE ACTIVATES THE CAPTURE PHASE SHIELD 🌟
    },

    // 3. UI HELPER
    focusScanner: function() {
        console.log("MAE System: Awaiting Barcode...");
    }
};

window.Labels = Labels;


