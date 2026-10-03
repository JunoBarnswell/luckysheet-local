import { createSpreadsheetSdk } from '@react-sheets/sdk';

/** Browser deployment configuration; SDK owns all authentication state and credentials. */
const authority = import.meta.env.VITE_OIDC_ISSUER?.trim();
const clientId = import.meta.env.VITE_OIDC_CLIENT_ID?.trim();
export const sdk = createSpreadsheetSdk({
  ...(authority && clientId ? { oidc: {
    authority, clientId,
    scope: import.meta.env.VITE_OIDC_SCOPE?.trim() || undefined,
    audience: import.meta.env.VITE_OIDC_AUDIENCE?.trim() || undefined,
    silentRedirectUri: import.meta.env.VITE_OIDC_SILENT_REDIRECT_URI?.trim() || undefined,
  } } : {}),
});
