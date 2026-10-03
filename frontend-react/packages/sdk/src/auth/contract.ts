import type { SdkError } from '../error';

export type AuthPhase = 'anonymous' | 'guest' | 'authenticated' | 'error' | 'loading' | 'unconfigured';
/** Verified by the spreadsheet server; contains no credential material. */
export interface AuthContext {
  readonly authority: string;
  readonly subject: string;
  readonly principal: string;
  readonly scopeId: string;
  readonly sessionId: string;
  readonly tenantId: string | null;
  readonly appCode: string | null;
  readonly employmentId: string | null;
  readonly contextVersion: number;
  readonly contextId: string;
}
export interface AuthSnapshot {
  readonly mode?: 'local' | 'oidc' | 'delegated';
  readonly context: AuthContext | null;
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
  refresh(): Promise<void>;
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
  readonly redirectUri?: string;
}
export type CredentialEvent = 'credentials-changed' | 'context-changed' | 'signed-out';
export interface BearerCredential {
  readonly token: string;
  /** Unix milliseconds, supplied by the credential owner. Server still validates expiry. */
  readonly expiresAt: number;
}
export interface BearerCredentialSource {
  acquire(): Promise<BearerCredential | null>;
  subscribe(listener: (event: CredentialEvent) => void): () => void;
  signIn?(returnTo?: string): Promise<void>;
  signOut(): Promise<void>;
}
export interface HostSessionSource {
  getCsrfToken(): Promise<string>;
  subscribe(listener: (event: CredentialEvent) => void): () => void;
  signIn?(returnTo?: string): Promise<void>;
  signOut(): Promise<void>;
}
export type AuthSource =
  | Readonly<{ kind: 'oidc'; configuration: OidcConfiguration }>
  | Readonly<{ kind: 'external-bearer'; credentials: BearerCredentialSource }>
  | Readonly<{ kind: 'host-session'; session: HostSessionSource }>;
export interface AuthOptions {
  readonly source?: AuthSource;
  readonly baseUrl?: string;
  readonly collaborationUrl?: string;
  readonly fetch?: typeof globalThis.fetch;
}
