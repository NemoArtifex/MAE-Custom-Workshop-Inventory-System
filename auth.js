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

