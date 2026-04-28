import { LogLevel } from '@azure/msal-browser';

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID}`,
    redirectUri: import.meta.env.VITE_REDIRECT_URI || 'http://localhost:5173',
    postLogoutRedirectUri: import.meta.env.VITE_REDIRECT_URI || 'http://localhost:5173',
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: true,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        switch (level) {
          case LogLevel.Error:   console.error(message); return;
          case LogLevel.Warning: console.warn(message);  return;
        }
      },
      logLevel: LogLevel.Warning,
    },
    allowNativeBroker: false,
  },
  telemetry: { enabled: false },
};

export const loginRequest = {
  scopes: ['User.Read', 'openid', 'profile'],
};

// ✅ CORRECT: SharePoint scope uses ROOT domain only — never a site path.
// ❌ WRONG:  'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/AllSites.Write'
// ✅ RIGHT:  'https://pmwgroupcom.sharepoint.com/AllSites.Write'
export const sharePointRequest = {
  scopes: [
    'https://pmwgroupcom.sharepoint.com/AllSites.Write',
    'https://pmwgroupcom.sharepoint.com/AllSites.Manage',
  ],
};