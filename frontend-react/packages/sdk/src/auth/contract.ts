import type { SdkError } from '../error';

export type AuthPhase = 'anonymous' | 'authenticated' | 'error' | 'loading' | 'unconfigured';
export interface AuthSnapshot {
  readonly mode?: 'local' | 'oidc';
  readonly bootstrapRequired: boolean;
  readonly displayName: string | null;
  readonly error: SdkError | null;
  readonly phase: AuthPhase;
  readonly subject: string | null;
  readonly capabilities: Readonly<{ canManageUsers: boolean }>;
}
export interface AuthSession {
  getSnapshot(): AuthSnapshot;
  subscribe(listener: () => void): () => void;
  initialize(): Promise<void>;
  authenticate(username: string, password: string): Promise<void>;
  bootstrap(token: string, username: string, password: string, displayName: string): Promise<void>;
  signIn(returnTo?: string): Promise<void>;
  signOut(): Promise<void>;
}
/** OIDC deployment settings are host configuration; credential lifecycle is owned by SDK. */
export interface OidcConfiguration {
  readonly authority: string;
  readonly clientId: string;
  readonly scope?: string;
  readonly audience?: string;
  readonly silentRedirectUri?: string;
}
export interface AuthOptions {
  readonly oidc?: OidcConfiguration;
  readonly baseUrl?: string;
  readonly fetch?: typeof globalThis.fetch;
}
