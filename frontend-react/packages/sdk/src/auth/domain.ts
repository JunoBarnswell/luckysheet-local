import { SdkError } from '../error';
import type { IdentityActions, LocalUser } from '../identity/contract';
import type { AuthOptions, AuthSession, AuthSnapshot } from './contract';
import { BrowserOidcSession } from './oidc';

const emptySnapshot = (): AuthSnapshot => Object.freeze({
  phase: 'loading', bootstrapRequired: false, subject: null, displayName: null, error: null,
  capabilities: Object.freeze({ canManageUsers: false }),
});
function record(value: unknown): value is Record<string, unknown> { return Boolean(value) && typeof value === 'object' && !Array.isArray(value); }

/** Only SDK runtime code holds the credential owner; the public actions contain no HTTP or token methods. */
export class AuthDomain {
  private snapshot = emptySnapshot();
  private readonly listeners = new Set<() => void>();
  private readonly fetchPort: typeof globalThis.fetch;
  private oidc: BrowserOidcSession | null = null;
  private oidcUnsubscribe: (() => void) | null = null;
  private initialization: Promise<void> | null = null;
  private csrfToken: string | null = null;
  private disposed = false;
  readonly session: AuthSession;
  readonly identity: IdentityActions;

  constructor(private readonly options: AuthOptions = {}) {
    this.fetchPort = options.fetch ?? ((input, init) => globalThis.fetch(input, init));
    this.session = Object.freeze({
      getSnapshot: () => this.snapshot,
      subscribe: (listener: () => void) => { this.assertActive(); this.listeners.add(listener); return () => { this.listeners.delete(listener); }; },
      initialize: () => this.initialize(),
      authenticate: (username: string, password: string) => this.authenticate(username, password),
      bootstrap: (token: string, username: string, password: string, displayName: string) => this.bootstrap(token, username, password, displayName),
      signIn: (returnTo?: string) => this.signIn(returnTo),
      signOut: () => this.signOut(),
    });
    this.identity = Object.freeze({
      listUsers: () => this.listUsers(),
      createUser: (input: { username: string; displayName: string; password: string }) => this.adminRequest('/api/admin/users', 'POST', input).then(() => undefined),
      setUserEnabled: (userId: string, enabled: boolean) => this.adminRequest(`/api/admin/users/${encodeURIComponent(userId)}`, 'PATCH', { enabled }).then(() => undefined),
      resetPassword: (userId: string, password: string) => this.adminRequest(`/api/admin/users/${encodeURIComponent(userId)}/password`, 'POST', { password }).then(() => undefined),
    });
  }
  // Internal credential providers; never exported on the SDK public AuthSession.
  getAccessToken = async (): Promise<string | null> => { this.assertActive(); return this.oidc?.getAccessToken() ?? null; };
  getCsrfToken = (): string | null => { this.assertActive(); return this.csrfToken; };
  private assertActive(): void {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'auth', 'SDK 认证会话已释放。', '请创建新的 SDK 实例。');
  }
  private publish(snapshot: AuthSnapshot): void {
    if (this.disposed) return;
    this.snapshot = Object.freeze({ ...snapshot, capabilities: Object.freeze({ ...snapshot.capabilities }) });
    for (const listener of this.listeners) listener();
  }
  private rejectIdentity(cause: unknown, operation: string): SdkError {
    const error = cause instanceof SdkError ? cause : new SdkError('SERVICE_UNAVAILABLE', operation, '认证服务不可用。', '请检查服务连接后重试。', { cause });
    this.csrfToken = null;
    this.initialization = null;
    this.publish({ ...emptySnapshot(), mode: this.snapshot.mode, phase: 'error', error });
    return error;
  }
  private async request(path: string, method = 'GET', body?: object): Promise<Response> {
    this.assertActive();
    const headers = new Headers();
    if (body !== undefined) headers.set('Content-Type', 'application/json');
    if (this.csrfToken && method !== 'GET') headers.set('X-CSRF-TOKEN', this.csrfToken);
    const bearer = await this.getAccessToken();
    if (bearer) headers.set('Authorization', `Bearer ${bearer}`);
    let response: Response;
    try {
      response = await this.fetchPort(this.options.baseUrl ? new URL(path, this.options.baseUrl).toString() : path,
        { method, credentials: 'same-origin', cache: 'no-store', headers, ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
    } catch (cause) {
      throw new SdkError('SERVICE_UNAVAILABLE', path, '认证服务无法连接。', '请检查服务器和网络后重试。', { cause });
    }
    this.assertActive();
    if (!response.ok) {
      const code = response.status === 401 ? 'UNAUTHENTICATED' : response.status === 403 ? 'FORBIDDEN' : response.status >= 500 ? 'SERVICE_UNAVAILABLE' : 'REQUEST_REJECTED';
      throw new SdkError(code, path, `认证请求失败 (${response.status})。`, response.status === 401 ? '请确认账号和密码，重新登录。' : '请检查权限或输入后重试。', { status: response.status });
    }
    return response;
  }
  private async json(response: Response, operation: string): Promise<unknown> {
    try { return await response.json(); }
    catch (cause) { throw new SdkError('CONTRACT_INVALID', operation, '服务器响应不是有效 JSON。', '请修复服务器响应契约后重试。', { cause }); }
  }
  private initialize(): Promise<void> {
    this.assertActive();
    if (this.initialization) return this.initialization;
    const pending = this.initializeMode();
    this.initialization = pending;
    void pending.then(() => { if (this.snapshot.phase === 'error') this.initialization = null; }, () => { this.initialization = null; });
    return pending;
  }
  private async initializeMode(): Promise<void> {
    this.publish({ ...emptySnapshot() });
    try {
      const configuration = await this.json(await this.request('/api/auth/config'), 'auth.config');
      if (!record(configuration) || (configuration.mode !== 'oidc' && configuration.mode !== 'local')) {
        throw new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.config', '服务器返回了无效认证模式。', '请配置 local 或 oidc 认证模式。');
      }
      if (configuration.mode === 'oidc') {
        this.oidcUnsubscribe?.();
        this.oidc?.dispose();
        this.oidc = new BrowserOidcSession(this.options.oidc);
        this.oidcUnsubscribe = this.oidc.subscribe(() => this.publish({ ...this.oidc!.getSnapshot(), mode: 'oidc' }));
        await this.oidc.initialize();
      } else await this.refresh();
    } catch (cause) { this.rejectIdentity(cause, 'auth.initialize'); }
  }
  private async refresh(): Promise<void> {
    try {
      const value = await this.json(await this.request('/api/auth/session'), 'auth.session');
      if (!record(value) || typeof value.authenticated !== 'boolean' || typeof value.csrfToken !== 'string' || !value.csrfToken
        || typeof value.admin !== 'boolean' || typeof value.bootstrapRequired !== 'boolean'
        || (value.subject !== null && typeof value.subject !== 'string')
        || (value.displayName !== null && typeof value.displayName !== 'string')
        || (value.authenticated && (typeof value.subject !== 'string' || !value.subject.trim()))
        || (!value.authenticated && (value.subject !== null || value.admin))) {
        throw new SdkError('CONTRACT_INVALID', 'auth.session', '认证会话响应契约无效。', '请修复服务器会话契约后重新登录。');
      }
      this.csrfToken = value.csrfToken;
      this.publish({ mode: 'local', phase: value.authenticated ? 'authenticated' : 'anonymous', subject: value.subject as string | null,
        displayName: value.displayName as string | null, bootstrapRequired: value.bootstrapRequired, error: null,
        capabilities: { canManageUsers: value.authenticated && value.admin } });
    } catch (cause) { throw this.rejectIdentity(cause, 'auth.session'); }
  }
  private async requireLocalMode(): Promise<void> {
    await this.initialize();
    if (this.snapshot.mode !== 'local' || this.snapshot.phase === 'error') throw this.snapshot.error ?? new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.local', '本地认证不可用。', '请使用配置的认证方式。');
  }
  private async authenticate(username: string, password: string): Promise<void> {
    await this.requireLocalMode();
    await this.request('/api/auth/login', 'POST', { username, password });
    await this.refresh();
  }
  private async bootstrap(token: string, username: string, password: string, displayName: string): Promise<void> {
    await this.requireLocalMode();
    await this.request('/api/auth/bootstrap', 'POST', { token, username, password, displayName });
    await this.refresh();
  }
  private async signIn(returnTo?: string): Promise<void> {
    await this.initialize();
    if (this.oidc) await this.oidc.signIn(returnTo);
    else await this.refresh();
  }
  private async signOut(): Promise<void> {
    this.assertActive();
    try {
      if (this.oidc) await this.oidc.signOut();
      else { await this.request('/api/auth/logout', 'POST'); await this.refresh(); }
    } catch (cause) { throw this.rejectIdentity(cause, 'auth.logout'); }
  }
  private async adminRequest(path: string, method: string, body?: object): Promise<Response> {
    this.assertActive();
    if (!this.snapshot.capabilities.canManageUsers) throw new SdkError('FORBIDDEN', 'identity.users', '需要用户管理权限。', '请使用具有用户管理权限的账号登录。');
    try { return await this.request(path, method, body); }
    catch (cause) {
      if (cause instanceof SdkError && cause.code === 'UNAUTHENTICATED') this.rejectIdentity(cause, 'identity.users');
      throw cause;
    }
  }
  private async listUsers(): Promise<readonly LocalUser[]> {
    const response = await this.adminRequest('/api/admin/users', 'GET');
    const value = await this.json(response, 'identity.listUsers');
    if (!Array.isArray(value) || value.some(user => !record(user) || typeof user.id !== 'string' || !user.id
      || typeof user.username !== 'string' || typeof user.displayName !== 'string' || typeof user.enabled !== 'boolean' || typeof user.admin !== 'boolean')) {
      throw new SdkError('CONTRACT_INVALID', 'identity.listUsers', '用户列表契约无效。', '请修复服务器用户列表后重试。');
    }
    return Object.freeze(value.map(user => Object.freeze({ id: user.id, username: user.username, displayName: user.displayName, enabled: user.enabled, admin: user.admin })));
  }
  dispose(): void {
    this.disposed = true;
    this.csrfToken = null;
    this.oidcUnsubscribe?.();
    this.oidc?.dispose();
    this.oidc = null;
    this.listeners.clear();
    this.snapshot = emptySnapshot();
  }
}
