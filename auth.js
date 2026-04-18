// =============CONFIGURATION: The "Blueprint"  ======================
// Defines the configuration object for the Microsoft Authentication Library (MSAL)
// Used to integrate Microsoft's identity and sign-in features into web apps


import { maeSystemConfig } from './config.js';

const msalConfig = {
    auth: {
        clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
        authority: "https://microsoftonline.com/common",
        redirectUri: "https://nemoartifex.github.io/MAE-Custom-Workshop-Inventory-System/",
        navigateToLoginRequestUrl: false 
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const myMSALObj = new window.msal.PublicClientApplication(msalConfig);

//================================================================
//=================== from app.js ===============
//const msalConfig = {
 //   auth: {
 //       clientId: "1f9f1df5-e39b-4845-bb07-ba7a683cf999",
 //       authority: "https://login.microsoftonline.com/common",
 //       //redirectUri: "http://localhost:5500" ,
 //       redirectUri: "https://nemoartifex.github.io/MAE-Custom-Workshop-Inventory-System/",
 //       navigateToLoginRequestUrl: false 
 ///   },
    // Defines how and where the app stores security tokens after received
    // Tokens stored for duration of browser's tab life 
    // "false": tells MSAL NOT to store the auth state in browser cookies  
 //   cache: {
 //       cacheLocation: "sessionStorage", // Simple and effective for workshop environments
 //       storeAuthStateInCookie: false,
 //   }
    
//};

//==========================================
//===============================================