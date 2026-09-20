import { BrowserOidcSession, type AuthSession, type AuthSnapshot } from './oidc';

interface LocalSessionResponse {
  authenticated: boolean;
  subject: string | null;
  displayName: string | null;
  admin: boolean;
  bootstrapRequired: boolean;
  csrfToken: string;
}

export interface LocalUser {
  id: string;
  username: string;
  displayName: string;
  enabled: boolean;
  admin: boolean;
}

class ApplicationAuthSession implements AuthSession {
  private snapshot: AuthSnapshot = { phase: 'loading', accessToken: null, subject: null, displayName: null, error: null };
  private readonly listeners = new Set<() => void>();
  private oidc: BrowserOidcSession | null = null;
  private initialization: Promise<void> | null = null;
  private csrfToken: string | null = null;

  getSnapshot = (): AuthSnapshot => this.snapshot;
  getAccessToken = async (): Promise<string | null> => this.oidc?.getAccessToken() ?? null;
  getCsrfToken = (): string | null => this.csrfToken;
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };
  private publish(snapshot: AuthSnapshot): void {
    this.snapshot = snapshot;
    for (const listener of this.listeners) listener();
  }
  initialize = (): Promise<void> => {
    this.initialization ??= this.initializeMode();
    return this.initialization;
  };
  private async initializeMode(): Promise<void> {
    try {
      const response = await fetch('/api/auth/config', { credentials: 'same-origin', cache: 'no-store' });
      if (!response.ok) throw new Error(`无法读取认证配置 (${response.status})`);
      const configuration: { mode: string } = await response.json();
      if (configuration.mode === 'oidc') {
        this.oidc = new BrowserOidcSession();
        this.oidc.subscribe(() => this.publish({ ...this.oidc!.getSnapshot(), mode: 'oidc' }));
        await this.oidc.initialize();
      } else if (configuration.mode === 'local') {
        await this.refresh();
      } else throw new Error('服务器返回了不支持的认证模式');
    } catch (error) {
      this.publish({ ...this.snapshot, phase: 'error', error: error instanceof Error ? error.message : '认证服务不可用' });
      this.initialization = null;
    }
  }
  async refresh(): Promise<void> {
    const response = await fetch('/api/auth/session', { credentials: 'same-origin', cache: 'no-store' });
    if (!response.ok) throw new Error(`无法读取会话 (${response.status})`);
    const session: LocalSessionResponse = await response.json();
    if (typeof session.authenticated !== 'boolean' || typeof session.csrfToken !== 'string') throw new Error('会话响应契约无效');
    this.csrfToken = session.csrfToken;
    this.publish({ mode: 'local', phase: session.authenticated ? 'authenticated' : 'anonymous', accessToken: null,
      subject: session.subject, displayName: session.displayName, admin: session.admin, bootstrapRequired: session.bootstrapRequired, error: null });
  }
  async request(path: string, method: string, body?: object): Promise<Response> {
    const headers = new Headers({ 'Content-Type': 'application/json' });
    if (this.csrfToken) headers.set('X-CSRF-TOKEN', this.csrfToken);
    const response = await fetch(path, { method, credentials: 'same-origin', headers, body: body === undefined ? undefined : JSON.stringify(body) });
    if (!response.ok) {
      const message = await response.text();
      throw new Error(`请求失败 (${response.status}): ${message.slice(0, 500)}`);
    }
    return response;
  }
  authenticate = async (username: string, password: string): Promise<void> => {
    await this.request('/api/auth/login', 'POST', { username, password });
    await this.refresh();
  };
  bootstrap = async (token: string, username: string, password: string, displayName: string): Promise<void> => {
    await this.request('/api/auth/bootstrap', 'POST', { token, username, password, displayName });
    await this.refresh();
  };
  signIn = async (returnTo?: string): Promise<void> => {
    if (this.oidc) await this.oidc.signIn(returnTo);
    else await this.refresh();
  };
  signOut = async (): Promise<void> => {
    if (this.oidc) await this.oidc.signOut();
    else {
      await this.request('/api/auth/logout', 'POST');
      await this.refresh();
    }
  };
}

const session = new ApplicationAuthSession();
export function getAuthSession(): ApplicationAuthSession { return session; }
