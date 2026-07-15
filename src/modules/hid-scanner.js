/**
 * src/modules/hid-scanner.js - MAE Custom Digital Solutions
 * Capability: Industrial HID Scanning (Inateck-75S) and ID extraction.
 * Responsibility: Listens to rapid hardware keystrokes and cleans URLs. Contains zero UI code.
 */

export const HidScanner = {
    // 1. EXTRACT ID: Extracts raw tags or IDs embedded within query parameters/URLs
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
            console.warn("MAE Scanner: Error parsing URL formatting, reverting to raw string data.");
            return decodedText;
        }
    },

    // 2. HARDWARE TIMEOUT DEFERRAL BUFFER: Isolates laser bursts from human typing
    initializeGlobalScanner: function() {
        let scanBuffer = "";
        let lastKeyTime = 0;
        let scanTimeout = null;

        console.log("MAE Scanner Axis: Global background HID window listener successfully engaged.");

        window.addEventListener("keydown", (event) => {
            // Safety Screen: Ignore modifier control keys
            if (event.key === "Shift" || event.key === "Control" || event.key === "Alt" || event.key === "Meta") {
                return;
            }

            // Form Pass-Through Check: Identify if operator is actively typing in a standard user field
            const activeElement = document.activeElement;
            const isUserTypingInOpenForm = activeElement && 
                                           activeElement.tagName === "INPUT" && 
                                           activeElement.id !== "field-Tag_ID" && 
                                           activeElement.id !== "mae-bulk-container-input";

            const currentTime = Date.now();
            const timeDiff = currentTime - lastKeyTime;
            lastKeyTime = currentTime;

            // Device Timing Burst Window: Scanners fire inputs with sub-30ms deltas
            if (timeDiff < 30 || scanBuffer === "") {
                
                // Block background collection if human typing cadence overflows into buffer
                if (isUserTypingInOpenForm && timeDiff >= 30) {
                    return;
                }

                // Freeze native browser actions to stop characters leaking into wrong text fields
                if (!isUserTypingInOpenForm && event.key !== "Enter") {
                    event.preventDefault();
                }

                if (event.key !== "Enter") {
                    scanBuffer += event.key;
                }

                // Defer character parsing until the 40ms stream quiet point settles cleanly
                clearTimeout(scanTimeout);
                scanTimeout = setTimeout(async () => {
                    if (scanBuffer.trim().length > 2) {
                        const cleanBarcode = this.extractCleanId(scanBuffer);
                        console.log(`MAE Scanner Matrix: Clear device burst captured: [${cleanBarcode}]. Dispatching to traffic controller.`);
                        
                        const finalPayload = cleanBarcode;
                        scanBuffer = ""; // Flush memory buffer instantly to protect against duplicate reads

                        // Safely forward clean data payload to the master router hook in app.js
                        if (typeof window.handleUniversalLookup === "function") {
                            await window.handleUniversalLookup(finalPayload);
                        } else {
                            console.warn("MAE Scanner Matrix Fault: The global window.handleUniversalLookup router engine is not initialized.");
                        }
                    } else {
                        scanBuffer = ""; // Flush accidental text hits or single click triggers
                    }
                }, 40);
            } else {
                scanBuffer = ""; // Time boundary split exceeded: treat as slow manual hand-typing
            }
        });
    }
};

// 🌟 THE ARRAYS ALIGNMENT BRIDGE 🌟
// Force-map to both names simultaneously so old calling scripts inside app.js/ui.js can't crash!
window.HidScanner = HidScanner;
window.Labels = { extractCleanId: HidScanner.extractCleanId };

// Ignite the hardware background thread
HidScanner.initializeGlobalScanner();