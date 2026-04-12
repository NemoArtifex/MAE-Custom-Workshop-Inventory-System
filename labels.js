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
                return url.pathname.split('/').pop() || url.searchParams.get("id");
            }
            return decodedText.trim();
        } catch (e) {
            console.warn("MAE System: Error parsing ID, using raw text.");
            return decodedText;
        }
    },

    // 2. INDUSTRIAL HID LISTENER
    // This listens for the "keystroke burst" from the Inateck-75S
    initHIDScanner: function(onSuccessCallback) {
        let buffer = "";
        let lastKeyTime = Date.now();

        console.log("MAE System: Industrial HID Listener Active.");

        document.addEventListener('keydown', (e) => {
            const currentTime = Date.now();
            
            // If typing is slow (>100ms between keys), it's a human, not the scanner.
            if (currentTime - lastKeyTime > 100) {
                buffer = ""; 
            }
            lastKeyTime = currentTime;

            // Inateck-75S sends 'Enter' at the end of a barcode scan
            if (e.key === 'Enter') {
                if (buffer.length > 2) { 
                    const cleanId = this.extractCleanId(buffer);
                    onSuccessCallback(cleanId);
                    buffer = ""; 
                    e.preventDefault(); 
                }
            } else {
                // Buffer the keys as they come in fast
                if (e.key.length === 1) {
                    buffer += e.key;
                }
            }
        });
    },

    // 3. UI HELPER
    focusScanner: function() {
        console.log("MAE System: Awaiting Barcode...");
    }
};

window.Labels = Labels;


