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
    // Check if the library loaded (now locally)
    if (typeof Html5Qrcode === 'undefined') {
        UI.showError("Scanner engine not loaded.");
        return;
    }

    // 1. Prepare the UI
    UI.renderScannerUI(); 

    // 2. Initialize the scanner instance
    const html5QrCode = new Html5Qrcode("reader");
    
    // RUGGED CONFIG: Keep it simple for the initial start to avoid Safari crashes
    const config = { 
        fps: 20,                       
        qrbox: { width: 250, height: 250 }, 
        aspectRatio: 1.0,
        disableFlip: true
    };

    // 3. Start the Camera
    html5QrCode.start(
        { facingMode: "environment" }, 
        config,
        (decodedText) => {  
            const cleanId = this.extractCleanId(decodedText);              
            html5QrCode.stop().then(() => onSuccessCallback(cleanId));
        }
    ).then(() => {
        // RUGGED IOS FIX: Prevent full-screen hijacking
        const videoElement = document.querySelector('#reader video');
        if (videoElement) {
            videoElement.setAttribute('playsinline', 'true');
            videoElement.setAttribute('webkit-playsinline', 'true');
            videoElement.style.display = "block";
        }

        // 4. THE HARDWARE UPGRADE (Resolution & Torch)
        // We wait a full 1000ms for Safari to stabilize before "tuning" the lens
        setTimeout(async () => {
            try {
                const track = html5QrCode.getRunningTrack();
                if (track) {
                    const capabilities = track.getCapabilities();
                    
                    const constraints = { advanced: [] };

                    // Enable Flashlight if supported (Essential for paper contrast)
                    if (capabilities.torch) {
                        constraints.advanced.push({ torch: true });
                    }

                    // Force HD for sharp paper scanning
                    constraints.advanced.push({ width: 1280, height: 720 });

                    await track.applyConstraints(constraints);
                    console.log("MAE System: Hardware optimized (HD + Torch).");
                }
            } catch (e) {
                console.warn("MAE System: Hardware optimization ignored by Safari.", e);
            }
        }, 1000); 

        console.log("MAE System: Camera feed active.");

    }).catch(err => {
        console.error("MAE System: Safari Handshake Failed:", err);
        // This clears the "Initialization" state and lets the user try again
        UI.showError("Camera error. Please refresh the page and ensure Safari permissions are granted.");
    });
}
};

window.Labels = Labels; 
