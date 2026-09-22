/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_OIDC_AUDIENCE?: string;
  readonly VITE_OIDC_CLIENT_ID?: string;
  readonly VITE_OIDC_ISSUER?: string;
  readonly VITE_OIDC_SCOPE?: string;
  readonly VITE_OIDC_SILENT_REDIRECT_URI?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
