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

        // Check if the library loaded from the CDN
        if (typeof Html5Qrcode === 'undefined') {
            UI.showError("Scanner engine not loaded. Check internet connection.");
            console.error("MAE System: html5-qrcode library is missing.");
            return;
        }

        // 1. Prepare the UI
        UI.renderScannerUI(); 

        // 2. Initialize the scanner instance
        const html5QrCode = new Html5Qrcode("reader");
        
        const config = { 
            fps: 20, 
            qrbox: {width: 250, height: 250},
            aspectRatio: 1.0,
            disableFlip: true, //Ensures image is not reversed
            videoConstraints: {
                facingMode: "environment" 
            }
        };

        // 3. Start the Camera
        html5QrCode.start(
            { facingMode: "environment" }, 
            config,
            (decodedText) => {
                // SUCCESS: Found a code
                const cleanId = this.extractCleanId(decodedText);
                
                html5QrCode.stop().then(() => {
                    console.log(`MAE System: Scan successful. ID: ${cleanId}`);
                    onSuccessCallback(cleanId);
                }).catch(err => console.warn("MAE System: Error stopping scanner", err));
            }
        ).then(() => {
            // RUGGED IOS FIX: Once camera starts, force the video to play inline 
            // inside your Navy box instead of going full-screen or staying black.
            const videoElement = document.querySelector('#reader video');
            if (videoElement) {
                videoElement.setAttribute('playsinline', 'true');
                videoElement.setAttribute('webkit-playsinline', 'true');
                videoElement.style.display = "block";
                
                // Remove the "Initializing" message from ui.js once video is live
                const loader = document.getElementById("loading-message");
                if (loader) loader.style.display = "none";
            }
            console.log("MAE System: Camera feed active.");
        }).catch(err => {
            console.error("MAE System: Camera access failed.", err);
            UI.showError("Camera error. Please ensure Safari permissions are granted.");
        });
    }
};

window.Labels = Labels; 
