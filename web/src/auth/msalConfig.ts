import { LogLevel, type Configuration } from '@azure/msal-browser';

const clientId = import.meta.env.VITE_CLIENT_ID as string;

if (!clientId) {
  console.warn(
    '[msal] VITE_CLIENT_ID is not set. Create web/.env.local with VITE_CLIENT_ID=<app-id>.',
  );
}

export const msalConfig: Configuration = {
  auth: {
    clientId: clientId ?? '00000000-0000-0000-0000-000000000000',
    authority:
      (import.meta.env.VITE_AUTHORITY as string) ?? 'https://login.microsoftonline.com/common',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Warning,
      piiLoggingEnabled: false,
      loggerCallback: (level, message) => {
        if (level === LogLevel.Error) console.error(message);
      },
    },
  },
};

/**
 * Scope sent when calling our own backend. The backend uses OBO to swap this
 * for a Microsoft Graph token. The scope value is the App ID URI exposed by
 * the same app registration (api://<clientId>/access_as_user).
 */
export const apiScope = `api://${clientId}/access_as_user`;
