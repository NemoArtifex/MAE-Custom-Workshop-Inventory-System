// =============CONFIGURATION: The "Blueprint"  ======================
// Defines the configuration object for the Microsoft Authentication Library (MSAL)
// Used to integrate Microsoft's identity and sign-in features into web apps


import { maeSystemConfig } from './config.js';

const msalConfig = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://nemoartifex.github.io/MAE-Custom-Workshop-Inventory-System/",
        navigateToLoginRequestUrl: false 
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const myMSALObj = new window.msal.PublicClientApplication(msalConfig);

// =========================================================================
// ======= MAE SECURED HARDWARE INTERLOCK: AUTH BUTTON BINDING CORE ========
// =========================================================================
window.addEventListener("DOMContentLoaded", () => {
    const authButton = document.getElementById("auth-btn");
    if (authButton) {
        // 🔒 THE GLOBAL ROUTING BRIDGE: Hooks your physical click to your authentic signIn macro in app.js
        authButton.onclick = () => {
            if (typeof window.signIn === "function") {
                window.signIn();
            } else {
                alert("MAE Authentication Engine: Connecting with Microsoft Graph. Please standby.");
                // Fallback direct redirection handler if window mapping is still warming up
                if (window.myMSALObj) {
                    window.myMSALObj.loginRedirect({
                        scopes: ["User.Read", "Files.ReadWrite"],
                        prompt: "select_account"
                    });
                }
            }
        };
        console.log("MAE Auth Engine: Universal click shortcut successfully anchored to your master login macro.");
    } else {
        console.warn("MAE Auth Engine: Sidebar button target missing during DOM load check.");
    }
});

