import { UserManager, WebStorageStateStore, type User } from 'oidc-client-ts';
import type { AuthTokenProvider } from '@react-sheets/protocol';

export type AuthPhase = 'anonymous' | 'authenticated' | 'error' | 'loading' | 'unconfigured';

export interface AuthSnapshot {
  accessToken: string | null;
  displayName: string | null;
  error: string | null;
  phase: AuthPhase;
  subject: string | null;
}

export interface AuthSession {
  getAccessToken: AuthTokenProvider;
  getSnapshot: () => AuthSnapshot;
  initialize: () => Promise<void>;
  signIn: (returnTo?: string) => Promise<void>;
  signOut: () => Promise<void>;
  subscribe: (listener: () => void) => () => void;
}

interface OidcConfiguration {
  audience?: string;
  authority: string;
  clientId: string;
  scope: string;
  silentRedirectUri: string;
}

const AUTH_RETURN_TO_KEY = 'react-sheets:oidc:return-to';
const initialSnapshot: AuthSnapshot = {
  accessToken: null,
  displayName: null,
  error: null,
  phase: 'loading',
  subject: null,
};

function callbackUri(): string {
  return new URL('/auth/callback', window.location.origin).toString();
}

function silentCallbackUri(): string {
  return new URL('/auth/silent-renew', window.location.origin).toString();
}

function readConfiguration(): OidcConfiguration | null {
  const authority = import.meta.env.VITE_OIDC_ISSUER?.trim();
  const clientId = import.meta.env.VITE_OIDC_CLIENT_ID?.trim();
  if (!authority || !clientId) return null;
  return {
    authority,
    clientId,
    scope: import.meta.env.VITE_OIDC_SCOPE?.trim() || 'openid profile email',
    audience: import.meta.env.VITE_OIDC_AUDIENCE?.trim() || undefined,
    silentRedirectUri: import.meta.env.VITE_OIDC_SILENT_REDIRECT_URI?.trim() || silentCallbackUri(),
  };
}

function toSnapshot(user: User | null, phase: AuthPhase, error: string | null = null): AuthSnapshot {
  if (!user || user.expired) {
    return {
      accessToken: null,
      displayName: null,
      error,
      phase,
      subject: null,
    };
  }
  const profile = user.profile as Record<string, unknown>;
  return {
    accessToken: user.access_token || null,
    displayName: typeof profile.name === 'string'
      ? profile.name
      : typeof profile.preferred_username === 'string'
        ? profile.preferred_username
        : typeof profile.sub === 'string'
          ? profile.sub
          : null,
    error,
    phase,
    subject: typeof profile.sub === 'string' ? profile.sub : null,
  };
}

class BrowserOidcSession implements AuthSession {
  private readonly configuration = readConfiguration();
  private readonly listeners = new Set<() => void>();
  private readonly manager: UserManager | null;
  private snapshot: AuthSnapshot = initialSnapshot;
  private initialized = false;

  constructor() {
    if (!this.configuration) {
      this.manager = null;
      this.snapshot = { ...initialSnapshot, phase: 'unconfigured' };
      return;
    }
    this.manager = new UserManager({
      authority: this.configuration.authority,
      client_id: this.configuration.clientId,
      redirect_uri: callbackUri(),
      silent_redirect_uri: this.configuration.silentRedirectUri,
      response_type: 'code',
      scope: this.configuration.scope,
      automaticSilentRenew: true,
      filterProtocolClaims: true,
      loadUserInfo: false,
      monitorSession: false,
      userStore: new WebStorageStateStore({ store: window.sessionStorage }),
      stateStore: new WebStorageStateStore({ store: window.sessionStorage }),
      extraQueryParams: this.configuration.audience ? { audience: this.configuration.audience } : undefined,
    });
    this.manager.events.addUserLoaded((user) => this.publish(toSnapshot(user, 'authenticated')));
    this.manager.events.addUserUnloaded(() => this.publish({ ...initialSnapshot, phase: 'anonymous' }));
    this.manager.events.addAccessTokenExpired(() => this.publish({ ...initialSnapshot, phase: 'anonymous' }));
    this.manager.events.addSilentRenewError((cause) => this.publish({
      ...toSnapshot(null, 'error', cause instanceof Error ? cause.message : 'Unable to renew the access token'),
    }));
  }

  getAccessToken: AuthTokenProvider = async () => {
    if (!this.manager) return null;
    const user = await this.manager.getUser();
    if (!user || user.expired) return null;
    return user.access_token || null;
  };

  getSnapshot = (): AuthSnapshot => this.snapshot;

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  async initialize(): Promise<void> {
    if (this.initialized) return;
    this.initialized = true;
    if (!this.manager) {
      this.publish({ ...initialSnapshot, phase: 'unconfigured' });
      return;
    }
    try {
      if (window.location.pathname === '/auth/callback' && new URLSearchParams(window.location.search).has('code')) {
        const user = await this.manager.signinRedirectCallback();
        const returnTo = window.sessionStorage.getItem(AUTH_RETURN_TO_KEY) || '/workbooks';
        window.sessionStorage.removeItem(AUTH_RETURN_TO_KEY);
        window.history.replaceState({}, '', returnTo);
        window.dispatchEvent(new PopStateEvent('popstate'));
        this.publish(toSnapshot(user, 'authenticated'));
        return;
      }
      if (window.location.pathname === '/auth/silent-renew') {
        await this.manager.signinSilentCallback();
        return;
      }
      const user = await this.manager.getUser();
      this.publish(toSnapshot(user, user?.expired ? 'anonymous' : user ? 'authenticated' : 'anonymous'));
    } catch (cause) {
      this.publish({
        ...initialSnapshot,
        error: cause instanceof Error ? cause.message : 'Unable to initialize authentication',
        phase: 'error',
      });
    }
  }

  async signIn(returnTo = `${window.location.pathname}${window.location.search}`): Promise<void> {
    if (!this.manager) {
      this.publish({
        ...initialSnapshot,
        error: 'Cloud workbooks require VITE_OIDC_ISSUER and VITE_OIDC_CLIENT_ID.',
        phase: 'unconfigured',
      });
      return;
    }
    window.sessionStorage.setItem(AUTH_RETURN_TO_KEY, returnTo.startsWith('/') ? returnTo : '/workbooks');
    await this.manager.signinRedirect();
  }

  async signOut(): Promise<void> {
    if (!this.manager) return;
    await this.manager.removeUser();
    this.publish({ ...initialSnapshot, phase: 'anonymous' });
  }

  private publish(snapshot: AuthSnapshot): void {
    this.snapshot = snapshot;
    for (const listener of this.listeners) listener();
  }
}

let session: AuthSession | null = null;

export function getAuthSession(): AuthSession {
  session ??= new BrowserOidcSession();
  return session;
}
