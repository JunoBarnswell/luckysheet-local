import { SdkError } from '../error';
import type { IdentityService } from '../identity/contract';
import { IdentityDomain } from '../identity/domain';
import type { UserAdministrationService } from '../users/contract';
import { UserAdministrationDomain } from '../users/domain';
import { GuestCapabilityDomain } from './guest-capability';
import type { AuthOptions, AuthSession, AuthSnapshot, BearerCredential, CredentialEvent } from './contract';
import { authContext } from './context';
import { BrowserOidcSession } from './oidc';

const emptySnapshot = (): AuthSnapshot => Object.freeze({
  phase: 'loading', context: null, bootstrapRequired: false, subject: null, displayName: null, error: null,
  capabilities: Object.freeze({ canManageUsers: false }),
});
function record(value: unknown): value is Record<string, unknown> { return Boolean(value) && typeof value === 'object' && !Array.isArray(value); }

/** Credential owner and verified-context boundary, private to the SDK composition root. */
export class AuthDomain {
  private snapshot = emptySnapshot();
  private readonly listeners = new Set<() => void>();
  private readonly fetchPort: typeof globalThis.fetch;
  private oidc: BrowserOidcSession | null = null;
  private sourceUnsubscribe: (() => void) | null = null;
  private initialization: Promise<void> | null = null;
  private acquisition: Promise<string | null> | null = null;
  private verification: Promise<void> | null = null;
  private csrfToken: string | null = null;
  private bearer: string | null = null;
  private mode: 'local' | 'oidc' | 'delegated' | undefined;
  private sourceVersion = 0;
  private generation = 0;
  private disposed = false;
  readonly session: AuthSession;
  private readonly identityDomain = new IdentityDomain();
  private readonly guest: GuestCapabilityDomain;
  private readonly usersDomain: UserAdministrationDomain;
  readonly identity: IdentityService;
  readonly users: UserAdministrationService;

  constructor(private readonly options: AuthOptions = {}) {
    if (options.baseUrl && (!/^https?:\/\//.test(options.baseUrl) || /[?#]/.test(options.baseUrl))) {
      throw new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.endpoint', 'API baseUrl 必须是固定 HTTP(S) 地址且不含 query/hash。', '请配置可信 API 来源及路径前缀。');
    }
    if (options.collaborationUrl) {
      const url = new URL(options.collaborationUrl);
      const api = new URL(options.baseUrl ?? this.origin());
      if (!['ws:', 'wss:'].includes(url.protocol) || url.host !== api.host || url.search || url.hash
        || (api.protocol === 'https:' && url.protocol !== 'wss:')) {
        throw new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.endpoint', '协作端点必须与可信 API 同一来源并使用对应安全协议。', '请配置 Gateway 的正式 WebSocket 路由。');
      }
    }
    this.fetchPort = options.fetch ?? ((input, init) => globalThis.fetch(input, init));
    this.session = Object.freeze({
      getSnapshot: () => this.snapshot,
      subscribe: (listener: () => void) => { this.assertActive(); this.listeners.add(listener); return () => { this.listeners.delete(listener); }; },
      initialize: () => this.initialize(), refresh: () => this.refresh(),
      authenticate: (username: string, password: string) => this.authenticate(username, password),
      bootstrap: (token: string, username: string, password: string, displayName: string) => this.bootstrap(token, username, password, displayName),
      signIn: (returnTo?: string) => this.signIn(returnTo), signOut: () => this.signOut(),
    });
    this.identity = this.identityDomain.service;
    this.guest = new GuestCapabilityDomain(() => this.retireGuestContext());
    this.usersDomain = new UserAdministrationDomain({ snapshot: () => this.snapshot, request: (path, method, body) => this.administrationRequest(path, method, body) });
    this.users = this.usersDomain.service;
    const source = options.source;
    if (source?.kind === 'external-bearer') this.sourceUnsubscribe = source.credentials.subscribe(event => this.sourceChanged(event));
    if (source?.kind === 'host-session') this.sourceUnsubscribe = source.session.subscribe(event => this.sourceChanged(event));
  }
  get accessContextKey(): string | null {
    return this.snapshot.context?.contextId ?? (this.snapshot.phase === 'guest' ? `guest:${this.guest.scope}` : null);
  }
  private retireGuestContext(): void {
    this.generation++;
    if (!this.disposed && this.snapshot.phase === 'guest') this.publish({ ...emptySnapshot(), mode: this.mode, phase: 'anonymous' });
  }
  private origin(): string { return typeof window === 'undefined' ? 'http://sdk.invalid' : window.location.origin; }
  private endpoint(path: string): string { return `${this.options.baseUrl?.replace(/\/$/, '') ?? ''}${path}`; }
  private credentialsMode(): RequestCredentials {
    return this.options.source?.kind === 'host-session' ? 'include' : this.mode === 'local' ? 'same-origin' : 'omit';
  }
  get collaborationUrl(): string | undefined {
    const configured = this.options.collaborationUrl ?? (typeof window === 'undefined' ? undefined : `${window.location.protocol === 'https:' ? 'wss:' : 'ws:'}//${window.location.host}/ws`);
    if (!configured) return undefined;
    const url = new URL(configured);
    if (this.snapshot.context) url.searchParams.set('context', this.snapshot.context.contextId);
    return url.toString();
  }
  private assertActive(): void {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'auth', 'SDK 认证会话已释放。', '请创建新的 SDK 实例。');
  }
  private assertGeneration(generation: number): void {
    this.assertActive();
    if (generation !== this.generation) throw new SdkError('STALE_OPERATION', 'auth.context', '请求所属身份上下文已退休。', '请在当前工作区重新打开对象；不要重试旧写入。');
  }
  private publish(snapshot: AuthSnapshot): void {
    if (this.disposed) return;
    if (snapshot.context?.contextId !== this.snapshot.context?.contextId) this.generation++;
    this.identityDomain.publish(snapshot.context, snapshot.displayName);
    this.snapshot = Object.freeze({ ...snapshot, capabilities: Object.freeze({ ...snapshot.capabilities }) });
    for (const listener of this.listeners) listener();
  }
  private rejectIdentity(cause: unknown, operation: string): SdkError {
    const error = cause instanceof SdkError ? cause : new SdkError('SERVICE_UNAVAILABLE', operation, '认证服务不可用。', '请检查服务连接后重试。', { cause });
    this.csrfToken = null; this.bearer = null; this.initialization = null;
    this.publish({ ...emptySnapshot(), mode: this.mode, phase: 'error', error });
    return error;
  }
  private sourceChanged(event: CredentialEvent): void {
    if (this.disposed) return;
    this.sourceVersion++; this.acquisition = null; this.verification = null;
    if (event !== 'credentials-changed') {
      this.bearer = null; this.csrfToken = null;
      this.publish({ ...emptySnapshot(), mode: this.mode, phase: event === 'signed-out' ? 'anonymous' : 'loading' });
    }
    if (event !== 'signed-out' && this.mode) void this.refresh().catch(() => undefined);
  }
  // Providers are bound to a context generation before a runtime or socket can use them.
  createTransport(): { baseUrl?: string; fetchImpl: typeof fetch; authTokenProvider: () => Promise<string | null>; csrfTokenProvider: () => string | null; shareTokenProvider: () => string | null } {
    const generation = this.generation;
    const expected = this.snapshot.context?.contextId ?? null;
    const assert = () => this.assertGeneration(generation);
    return {
      baseUrl: this.options.baseUrl?.replace(/\/$/, ''),
      authTokenProvider: async () => { assert(); const token = await this.getAccessToken(); assert(); return token; },
      csrfTokenProvider: () => { assert(); return this.csrfToken; },
      shareTokenProvider: () => { assert(); const token = this.snapshot.phase === 'authenticated' ? null : this.guest.read(); assert(); return token; },
      fetchImpl: async (input, init) => {
        assert();
        const url = new URL(input instanceof Request ? input.url : String(input), this.origin());
        const base = new URL(this.options.baseUrl ?? this.origin());
        const prefix = base.pathname.replace(/\/$/, '');
        if (url.origin !== base.origin || (prefix && !url.pathname.startsWith(`${prefix}/`))) {
          throw new SdkError('FORBIDDEN', 'auth.endpoint', '请求不属于可信 API 来源。', '请使用 SDK 配置的 API 端点。');
        }
        const headers = new Headers(init?.headers);
        if (expected) headers.set('X-Spreadsheet-Context', expected);
        const response = await this.fetchPort(input, { ...init, headers, credentials: this.credentialsMode() });
        assert();
        for (const method of ['json', 'text', 'arrayBuffer', 'blob', 'formData'] as const) {
          const read = response[method].bind(response);
          Object.defineProperty(response, method, { value: async () => { assert(); const value = await read(); assert(); return value; } });
        }
        if (response.status === 401) throw this.rejectIdentity(new SdkError('UNAUTHENTICATED', 'auth.context', '服务端拒绝当前凭证。', '请重新验证身份并打开工作簿。', { status: 401 }), 'auth.context');
        return response;
      },
    };
  }
  getAccessToken = async (): Promise<string | null> => {
    this.assertActive();
    if (this.mode === 'local' || this.options.source?.kind === 'host-session') return null;
    if (this.acquisition) return this.acquisition;
    const version = this.sourceVersion;
    const pending = (async () => {
      try {
        let credential: BearerCredential | null = null;
        if (this.options.source?.kind === 'external-bearer') credential = await this.options.source.credentials.acquire();
        else if (this.oidc) credential = await this.oidc.getCredential();
        this.assertActive();
        if (version !== this.sourceVersion) throw new SdkError('STALE_OPERATION', 'auth.credential', '凭证来源已变化。', '请使用当前身份重新发起操作。');
        if (credential && (typeof credential.token !== 'string' || !credential.token.trim()
          || !Number.isFinite(credential.expiresAt) || credential.expiresAt <= Date.now())) {
          throw new SdkError('UNAUTHENTICATED', 'auth.credential', '来源凭证无效或已过期。', '请通过凭证 owner 重新登录。');
        }
        const token = credential?.token ?? null;
        if (token !== this.bearer || (token && !this.snapshot.context)) await this.verifySession(token);
        return token;
      } catch (cause) {
        if (version !== this.sourceVersion || this.disposed) throw cause;
        throw this.rejectIdentity(cause, 'auth.credential');
      }
    })();
    this.acquisition = pending;
    void pending.finally(() => { if (this.acquisition === pending) this.acquisition = null; }).catch(() => undefined);
    return pending;
  };
  getCsrfToken = (): string | null => { this.assertActive(); return this.csrfToken; };
  private async request(path: string, method = 'GET', body?: object, bearer: string | null = this.bearer): Promise<Response> {
    this.assertActive();
    const version = this.sourceVersion;
    const headers = new Headers();
    if (body !== undefined) headers.set('Content-Type', 'application/json');
    if (this.csrfToken && method !== 'GET') headers.set('X-CSRF-TOKEN', this.csrfToken);
    if (bearer) headers.set('Authorization', `Bearer ${bearer}`);
    let response: Response;
    try {
      response = await this.fetchPort(this.endpoint(path), { method, credentials: this.credentialsMode(), cache: 'no-store', headers,
        ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
    } catch (cause) { throw new SdkError('SERVICE_UNAVAILABLE', path, '认证服务无法连接。', '请检查服务器和网络后重试。', { cause }); }
    this.assertActive();
    if (version !== this.sourceVersion) throw new SdkError('STALE_OPERATION', path, '认证响应属于旧凭证上下文。', '请在当前工作区重新验证身份。');
    if (!response.ok) {
      const code = response.status === 401 ? 'UNAUTHENTICATED' : response.status === 403 ? 'FORBIDDEN' : response.status >= 500 ? 'SERVICE_UNAVAILABLE' : 'REQUEST_REJECTED';
      throw new SdkError(code, path, `认证请求失败 (${response.status})。`, response.status === 401 ? '请确认凭证并重新登录。' : '请检查权限或输入后重试。', { status: response.status });
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
    const pending = this.initializeMode(); this.initialization = pending;
    void pending.then(() => { if (this.snapshot.phase === 'error') this.initialization = null; }, () => { this.initialization = null; });
    return pending;
  }
  private async initializeMode(): Promise<void> {
    this.publish(emptySnapshot());
    try {
      const configuration = await this.json(await this.request('/api/auth/config'), 'auth.config');
      if (!record(configuration) || !['local', 'oidc', 'delegated'].includes(configuration.mode as string)) {
        throw new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.config', '服务器返回了无效认证模式。', '请配置 local、oidc 或 delegated 认证模式。');
      }
      this.mode = configuration.mode as 'local' | 'oidc' | 'delegated';
      const source = this.options.source;
      if ((this.mode === 'local' && source) || (this.mode === 'delegated' && source?.kind !== 'host-session')
        || (this.mode === 'oidc' && source?.kind === 'host-session')) {
        throw new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.source', '凭证来源与服务端认证模式不一致。', '请配置正式 API 对应的凭证来源。');
      }
      if (source?.kind === 'oidc') {
        this.oidc?.dispose(); this.sourceUnsubscribe?.();
        this.oidc = new BrowserOidcSession(source.configuration);
        await this.oidc.initialize();
        this.sourceUnsubscribe = this.oidc.subscribe(() => {
          const state = this.oidc!.getSnapshot();
          if (state.phase === 'error') this.rejectIdentity(state.error, 'auth.oidc');
          else this.sourceChanged(state.phase === 'authenticated' ? 'credentials-changed' : 'signed-out');
        });
      }
      if (this.mode === 'oidc' && !source) {
        this.publish({ ...emptySnapshot(), mode: this.mode, phase: this.guest.scope ? 'guest' : 'unconfigured' }); return;
      }
      await this.refresh();
    } catch (cause) { this.rejectIdentity(cause, 'auth.initialize'); }
  }
  private refresh(): Promise<void> {
    this.assertActive();
    if (!this.mode) return this.initialize();
    if (this.mode === 'oidc' && !this.options.source) return Promise.reject(new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.source', '缺少 OIDC 或现有 bearer 凭证来源。', '请配置凭证 owner。'));
    if (this.verification) return this.verification;
    const version = this.sourceVersion;
    const pending = (async () => {
      try {
        if (this.mode === 'oidc') {
          const previous = this.bearer;
          const previouslyVerified = this.snapshot.context !== null;
          const token = await this.getAccessToken();
          // New credentials were already verified by the acquisition boundary.
          if (token === previous && (token === null || previouslyVerified)) await this.verifySession(token);
        } else await this.verifySession(null);
      } catch (cause) {
        if (version !== this.sourceVersion || this.disposed) throw cause;
        throw this.rejectIdentity(cause, 'auth.session');
      }
    })();
    this.verification = pending;
    void pending.finally(() => { if (this.verification === pending) this.verification = null; }).catch(() => undefined);
    return pending;
  }
  private async verifySession(token: string | null): Promise<void> {
    const version = this.sourceVersion;
    const value = await this.json(await this.request('/api/auth/session', 'GET', undefined, token), 'auth.session');
    if (version !== this.sourceVersion) throw new SdkError('STALE_OPERATION', 'auth.session', '旧会话响应已退休。', '请重新验证当前工作区。');
    if (!record(value) || typeof value.authenticated !== 'boolean' || typeof value.csrfToken !== 'string'
      || (this.mode === 'local' && !value.csrfToken) || typeof value.admin !== 'boolean' || typeof value.bootstrapRequired !== 'boolean'
      || (value.subject !== null && typeof value.subject !== 'string') || (value.displayName !== null && typeof value.displayName !== 'string')
      || (value.authenticated && (typeof value.subject !== 'string' || !value.subject.trim()))
      || (!value.authenticated && (value.subject !== null || value.admin || value.context !== null))
      || (token !== null && !value.authenticated)) {
      throw new SdkError('CONTRACT_INVALID', 'auth.session', '认证会话响应契约无效。', '请修复服务器会话契约后重新登录。');
    }
    const context = value.authenticated ? authContext(value.context, value.subject as string) : null;
    let csrf = value.csrfToken;
    if (this.options.source?.kind === 'host-session') {
      csrf = await this.options.source.session.getCsrfToken();
      if (typeof csrf !== 'string' || !csrf.trim()) throw new SdkError('CONTRACT_INVALID', 'auth.csrf', '宿主未提供有效 CSRF。', '请修复 Gateway 会话契约。');
    }
    if (version !== this.sourceVersion) throw new SdkError('STALE_OPERATION', 'auth.session', '会话验证期间工作区已变化。', '请使用当前工作区重新打开对象。');
    this.csrfToken = csrf || null; this.bearer = token;
    if (value.authenticated) this.guest.revoke();
    this.publish({ mode: this.mode, context, phase: value.authenticated ? 'authenticated' : this.guest.scope ? 'guest' : 'anonymous', subject: value.subject as string | null,
      displayName: value.displayName as string | null, bootstrapRequired: value.bootstrapRequired, error: null,
      capabilities: { canManageUsers: this.mode === 'local' && value.authenticated && value.admin } });
  }
  private async requireLocalMode(): Promise<void> {
    await this.initialize();
    if (this.mode !== 'local' || this.options.source || this.snapshot.phase === 'error') throw this.snapshot.error ?? new SdkError('AUTH_CONFIGURATION_ERROR', 'auth.local', '本地认证不可用。', '请使用配置的凭证来源。');
  }
  private beginLocalTransition(): void {
    this.sourceVersion++; this.verification = null; this.acquisition = null;
    if (this.snapshot.context) this.publish({ ...emptySnapshot(), mode: this.mode, phase: 'anonymous' });
  }
  private async authenticate(username: string, password: string): Promise<void> {
    await this.requireLocalMode(); this.beginLocalTransition(); await this.request('/api/auth/login', 'POST', { username, password }); await this.refresh();
  }
  private async bootstrap(token: string, username: string, password: string, displayName: string): Promise<void> {
    await this.requireLocalMode(); this.beginLocalTransition(); await this.request('/api/auth/bootstrap', 'POST', { token, username, password, displayName }); await this.refresh();
  }
  private async signIn(returnTo?: string): Promise<void> {
    await this.initialize();
    const source = this.options.source;
    if (this.oidc) await this.oidc.signIn(returnTo);
    else if (source?.kind === 'external-bearer' && source.credentials.signIn) await source.credentials.signIn(returnTo);
    else if (source?.kind === 'host-session' && source.session.signIn) await source.session.signIn(returnTo);
    else await this.refresh();
  }
  private async signOut(): Promise<void> {
    this.assertActive();
    const source = this.options.source;
    try {
      this.guest.revoke();
      if (this.oidc) await this.oidc.signOut();
      else if (source?.kind === 'external-bearer') await source.credentials.signOut();
      else if (source?.kind === 'host-session') await source.session.signOut();
      else { this.beginLocalTransition(); await this.request('/api/auth/logout', 'POST'); await this.refresh(); }
      if (source) this.sourceChanged('signed-out');
    } catch (cause) { throw this.rejectIdentity(cause, 'auth.logout'); }
  }
  private async administrationRequest(path: string, method: string, body?: object): Promise<Response> {
    this.assertActive();
    const transport = this.createTransport();
    const headers = new Headers();
    const token = await transport.authTokenProvider();
    if (token) headers.set('Authorization', `Bearer ${token}`);
    if (body) headers.set('Content-Type', 'application/json');
    if (method !== 'GET' && this.csrfToken) headers.set('X-CSRF-TOKEN', this.csrfToken);
    const response = await transport.fetchImpl(this.endpoint(path), { method, headers, ...(body ? { body: JSON.stringify(body) } : {}) });
    if (!response.ok) throw new SdkError(response.status === 403 ? 'FORBIDDEN' : response.status >= 500 ? 'SERVICE_UNAVAILABLE' : 'REQUEST_REJECTED', path, `用户管理请求被拒绝 (${response.status})。`, '请检查管理权限后重试。', { status: response.status });
    return response;
  }
  dispose(): void {
    this.disposed = true; this.generation++; this.sourceVersion++; this.csrfToken = null; this.bearer = null;
    this.usersDomain.dispose(); this.identityDomain.dispose(); this.guest.dispose();
    this.sourceUnsubscribe?.(); this.oidc?.dispose(); this.oidc = null; this.listeners.clear(); this.snapshot = emptySnapshot();
  }
}
