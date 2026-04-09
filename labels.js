/**
 * labels.js - MAE Custom Digital Solutions
 * Purpose: Handle QR scanning and ID extraction.
 * Philosophy: Practical & Rugged.
 */

import { UI } from './ui.js';

export const Labels = {
    // 1. EXTRACT ID: This makes the code flexible for any label (ToteScan, Metal, etc.)
    // It takes a raw string (like a URL) and returns just the ID.
    extractCleanId: function(decodedText) {
        try {
            // If the label is a URL (like ToteScan)
            if (decodedText.startsWith("http")) {
                const url = new URL(decodedText);
                // Grabs the ID from the end of the path or a query param
                return url.pathname.split('/').pop() || url.searchParams.get("id");
            }
            // If it's just a raw text ID (like your metallized tags)
            return decodedText.trim();
        } catch (e) {
            console.warn("MAE System: Error parsing ID, using raw text.");
            return decodedText;
        }
    },

    // 2. START SCANNER: Opens the phone camera
    startScanner: function(onSuccessCallback) {
        // We'll create this UI element in ui.js next
        UI.renderScannerUI(); 

        // Initialize the library you added to index.html
        const html5QrCode = new Html5Qrcode("reader");
        
        const config = { 
            fps: 10, 
            qrbox: { width: 250, height: 250 },
            aspectRatio: 1.0 
        };

        html5QrCode.start(
            { facingMode: "environment" }, // Use back camera
            config,
            (decodedText) => {
                const cleanId = this.extractCleanId(decodedText);
                
                // Stop camera to save battery/resources
                html5QrCode.stop().then(() => {
                    console.log(`MAE System: Scan successful. ID: ${cleanId}`);
                    onSuccessCallback(cleanId);
                });
            }
        ).catch(err => {
            console.error("MAE System: Camera access failed.", err);
            UI.showError("Camera error. Please ensure permissions are granted.");
        });
    }
};

window.Labels = Labels; 
