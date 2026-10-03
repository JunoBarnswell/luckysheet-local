import { WorkbookApiClient } from '@react-sheets/protocol';
import { RemoteAssetStore, WorkspacePersistence, WorkspaceStorageError, isWorkspaceStorageError, resolveShareToken, WorkbookSession,
  type WorkbookResolution, type WorkspacePersistenceState } from '@react-sheets/spreadsheet-app';
import type { AuthDomain } from '../auth/domain';
import { WorkbooksDomain } from '../workbooks/domain';
import { SdkError } from '../error';

export interface StorageReadiness {
  readonly state: WorkspacePersistenceState | 'warming' | 'failed';
  readonly error: WorkspaceStorageError | null;
}
/** The catalog public surface excludes protocol clients and persistence owners. */
export type WorkbooksActions = Omit<WorkbooksDomain, 'retire'>;

export class ApplicationRuntime {
  private catalogDomain: WorkbooksDomain;
  get catalog(): WorkbooksActions { return this.catalogDomain; }
  private readonly api: WorkbookApiClient;
  private persistence = new WorkspacePersistence();
  private snapshot: StorageReadiness = Object.freeze({ state: 'warming', error: null });
  private readonly listeners = new Set<() => void>();
  private readonly sessions = new Set<WorkbookSession>();
  private readiness: Promise<void> | null = null;
  private disposed = false;
  private users = 0;
  private owner: string | null = null;
  private readonly unsubscribeAuth: () => void;
  private releaseTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly shareTokenProvider = () => resolveShareToken();

  constructor(private readonly auth: AuthDomain, baseUrl?: string, fetchImpl?: typeof globalThis.fetch) {
    this.api = new WorkbookApiClient({ baseUrl, fetchImpl, authTokenProvider: auth.getAccessToken, csrfTokenProvider: auth.getCsrfToken, shareTokenProvider: this.shareTokenProvider });
    this.catalogDomain = this.createCatalog();
    this.unsubscribeAuth = auth.session.subscribe(() => {
      const subject = auth.session.getSnapshot().subject;
      if (subject !== this.owner) { this.owner = subject; this.resetWorkspace(); }
    });
  }
  private createCatalog(): WorkbooksDomain {
    return new WorkbooksDomain({
      persistence: this.persistence, remote: this.api,
      remoteAvailable: () => !this.disposed && (this.auth.session.getSnapshot().phase === 'authenticated' || Boolean(this.shareTokenProvider())),
      shareTokenProvider: this.shareTokenProvider,
    });
  }
  getSnapshot = (): StorageReadiness => this.snapshot;
  subscribe = (listener: () => void): (() => void) => { this.listeners.add(listener); return () => { this.listeners.delete(listener); }; };
  private publish(snapshot: StorageReadiness): void {
    if (this.disposed) return;
    this.snapshot = Object.freeze(snapshot);
    for (const listener of this.listeners) listener();
  }
  ensureStorageReady = (): Promise<void> => {
    if (this.disposed) return Promise.reject(new SdkError('RUNTIME_DISPOSED', 'workbooks.storage', 'SDK Runtime 已释放。', '请创建新的 SDK。'));
    if (this.readiness) return this.readiness;
    this.publish({ state: 'warming', error: null });
    const owner = this.persistence;
    const pending = owner.ensureReady().then(() => { if (owner === this.persistence) this.publish({ state: owner.state, error: null }); }, cause => {
      if (owner !== this.persistence) throw cause;
      const error = isWorkspaceStorageError(cause) ? cause : new WorkspaceStorageError({ code: 'STORAGE_MEMORY_TRANSACTION_FAILED', operation: 'open',
        message: cause instanceof Error ? cause.message : '工作簿缓存不可用。', recovery: '请重新建立缓存会话后重试。', cause });
      this.publish({ state: 'failed', error });
      this.readiness = null;
      throw error;
    });
    this.readiness = pending;
    return pending;
  };
  retryStorage = (): Promise<void> => { this.readiness = null; return this.ensureStorageReady(); };
  acquire(): () => void {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'runtime.acquire', 'SDK Runtime 已释放。', '请创建新的 SDK。');
    this.users++;
    if (this.releaseTimer) { clearTimeout(this.releaseTimer); this.releaseTimer = null; }
    let released = false;
    return () => {
      if (released) return;
      released = true;
      this.users--;
      if (!this.users) this.releaseTimer = setTimeout(() => {
        this.releaseTimer = null;
        if (this.users || this.disposed) return;
        this.resetWorkspace();
      }, 0);
    };
  }
  private resetWorkspace(): void {
    if (this.disposed) return;
    for (const session of this.sessions) session.dispose();
    this.sessions.clear();
    const previous = this.persistence;
    this.catalogDomain.retire();
    this.persistence = new WorkspacePersistence();
    this.catalogDomain = this.createCatalog();
    this.readiness = null;
    this.publish({ state: 'warming', error: null });
    void previous.disposeAsync();
  }
  createSession(resolution: WorkbookResolution): WorkbookSession {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'workbook.open', 'SDK Runtime 已释放。', '请创建新的 SDK。');
    if (typeof Worker === 'undefined') throw new SdkError('UNSUPPORTED_FEATURE', 'workbook.open', '此宿主缺少工作簿运行所需的 browser Worker。', '请提供支持 browser Worker 的浏览器宿主。');
    const session = new WorkbookSession({
      unitId: resolution.unitId, initialPhase: 'loading', resolution, api: this.api, workspacePersistence: this.persistence,
      authTokenProvider: this.auth.getAccessToken, shareTokenProvider: this.shareTokenProvider,
      recoverySubject: this.auth.session.getSnapshot().subject ?? undefined,
      pivotExecution: 'worker', assetStore: new RemoteAssetStore(resolution.unitId, this.api),
      onReady: () => this.catalog.markOpened(resolution).then(() => undefined),
    });
    this.sessions.add(session);
    return session;
  }
  closeSession(session: WorkbookSession): void { session.dispose(); this.sessions.delete(session); }
  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    this.unsubscribeAuth();
    this.catalogDomain.retire();
    if (this.releaseTimer) clearTimeout(this.releaseTimer);
    for (const session of this.sessions) session.dispose();
    this.sessions.clear();
    this.listeners.clear();
    await this.persistence.disposeAsync();
  }
}
