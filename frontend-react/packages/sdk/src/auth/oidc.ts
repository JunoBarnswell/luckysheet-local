import { UserManager, WebStorageStateStore, type User } from 'oidc-client-ts';
import type { AuthPhase, AuthSnapshot, OidcConfiguration } from './contract';
import { SdkError } from '../error';

const AUTH_RETURN_TO_KEY = 'react-sheets:oidc:return-to';
const initialSnapshot: AuthSnapshot = {
  bootstrapRequired: false,
  capabilities: Object.freeze({ canManageUsers: false }),
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

function toSnapshot(user: User | null, phase: AuthPhase, error: SdkError | null = null): AuthSnapshot {
  if (!user || user.expired) {
    return {
      bootstrapRequired: false,
      capabilities: Object.freeze({ canManageUsers: false }),
      displayName: null,
      error,
      phase,
      subject: null,
    };
  }
  const profile = user.profile as Record<string, unknown>;
  return {
    bootstrapRequired: false,
    capabilities: Object.freeze({ canManageUsers: false }),
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

export class BrowserOidcSession {
  private readonly configuration: OidcConfiguration | undefined;
  private readonly listeners = new Set<() => void>();
  private readonly manager: UserManager | null;
  private snapshot: AuthSnapshot = initialSnapshot;
  private initialized = false;

  constructor(configuration?: OidcConfiguration) {
    this.configuration = configuration?.authority.trim() && configuration.clientId.trim() ? configuration : undefined;
    if (!this.configuration) {
      this.manager = null;
      this.snapshot = { ...initialSnapshot, phase: 'unconfigured' };
      return;
    }
    this.manager = new UserManager({
      authority: this.configuration.authority,
      client_id: this.configuration.clientId,
      redirect_uri: callbackUri(),
      silent_redirect_uri: this.configuration.silentRedirectUri ?? silentCallbackUri(),
      response_type: 'code',
      scope: this.configuration.scope ?? 'openid profile email',
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
      ...toSnapshot(null, 'error', new SdkError('UNAUTHENTICATED', 'auth.oidc.renew', '登录凭证续期失败。', '请重新登录。', { cause })),
    }));
  }

  getAccessToken = async (): Promise<string | null> => {
    if (!this.manager) return null;
    const user = await this.manager.getUser();
    if (!user || user.expired) { this.publish({ ...initialSnapshot, phase: 'anonymous' }); return null; }
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
        const storedReturnTo = window.sessionStorage.getItem(AUTH_RETURN_TO_KEY);
        const returnTo = storedReturnTo?.startsWith('/') && !storedReturnTo.startsWith('//') ? storedReturnTo : '/workbooks';
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
        error: new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.oidc.initialize', 'OIDC 登录初始化失败。', '请检查 OIDC 部署配置后重新登录。', { cause }),
        phase: 'error',
      });
    }
  }

  async signIn(returnTo = `${window.location.pathname}${window.location.search}`): Promise<void> {
    if (!this.manager) {
      this.publish({
        ...initialSnapshot,
        error: new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.oidc.signIn', '缺少 OIDC issuer 和 client ID。', '请提供 OIDC 部署配置。'),
        phase: 'unconfigured',
      });
      return;
    }
    window.sessionStorage.setItem(AUTH_RETURN_TO_KEY, returnTo.startsWith('/') && !returnTo.startsWith('//') ? returnTo : '/workbooks');
    await this.manager.signinRedirect();
  }

  async signOut(): Promise<void> {
    if (!this.manager) return;
    await this.manager.removeUser();
    this.publish({ ...initialSnapshot, phase: 'anonymous' });
  }

  dispose(): void {
    void this.manager?.stopSilentRenew();
    this.listeners.clear();
    this.snapshot = initialSnapshot;
  }

  private publish(snapshot: AuthSnapshot): void {
    this.snapshot = Object.freeze({ ...snapshot });
    for (const listener of this.listeners) listener();
  }
}
