import { WorkbookApiClient } from '@react-sheets/protocol';
import { RemoteAssetStore, WorkspacePersistence, WorkspaceStorageError, isWorkspaceStorageError, type WorkspacePersistenceState } from '../../../spreadsheet-app/src/features/persistence';
import { WorkbookSession } from '../../../spreadsheet-app/src/workbook-session';
import type { WorkbookResolution } from '../../../spreadsheet-app/src/features/workbook-catalog';
import type { AuthDomain } from '../auth/domain';
import { WorkbooksDomain } from '../workbooks/domain';
import { DataDomain } from '../data/domain';
import { DimensionsDomain } from '../dimensions/domain';
import { getWorkbookObjectPort } from '../../../spreadsheet-app/src/workbook-object-port';
import { Workbook } from '../workbook/workbook';
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
  private api: WorkbookApiClient;
  private persistence = new WorkspacePersistence();
  private snapshot: StorageReadiness = Object.freeze({ state: 'warming', error: null });
  private readonly listeners = new Set<() => void>();
  private readonly sessions = new Set<WorkbookSession>();
  private readonly sessionByUnit = new Map<string, WorkbookSession>();
  private readonly sessionLeases = new Map<WorkbookSession, number>();
  private readonly workbookObjects = new Map<string, Workbook>();
  private readonly workbookOpens = new Map<string, Promise<Workbook>>();
  private resetting = false;
  private readonly assetStores = new Map<string, RemoteAssetStore>();
  private readonly data = new Map<WorkbookSession, DataDomain>();
  private readonly dimensions = new Map<WorkbookSession, DimensionsDomain>();
  private readiness: Promise<void> | null = null;
  private disposed = false;
  private users = 0;
  private owner: string | null = null;
  private readonly unsubscribeAuth: () => void;
  private releaseTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(private readonly auth: AuthDomain) {
    this.api = this.createApi();
    this.catalogDomain = this.createCatalog();
    this.unsubscribeAuth = auth.session.subscribe(() => {
      const subject = auth.accessContextKey;
      if (subject !== this.owner) { this.owner = subject; this.resetWorkspace(); }
    });
  }
  private createApi(): WorkbookApiClient {
    return new WorkbookApiClient(this.auth.createTransport());
  }
  private createCatalog(): WorkbooksDomain {
    return new WorkbooksDomain({
      openWorkbook: (resolution) => this.openWorkbook(resolution),
      persistence: this.persistence, remote: this.api, assetStoreFor: (unitId) => this.assetStoreFor(unitId),
      remoteAvailable: () => !this.disposed && (['authenticated', 'guest'].includes(this.auth.session.getSnapshot().phase)),
      shareTokenProvider: this.auth.createTransport().shareTokenProvider,
    });
  }
  private assetStoreFor(unitId: string): RemoteAssetStore {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'assets', 'SDK Runtime 已释放。', '请创建新的 SDK。');
    let store = this.assetStores.get(unitId);
    if (!store) { store = new RemoteAssetStore(unitId, this.api); this.assetStores.set(unitId, store); }
    return store;
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
      if (!this.users && !this.disposed && !this.resetting) this.releaseTimer = setTimeout(() => {
        this.releaseTimer = null;
        if (this.users || this.disposed) return;
        this.resetWorkspace();
      }, 0);
    };
  }
  private resetWorkspace(): void {
    if (this.disposed) return;
    this.resetting = true;
    for (const workbook of [...this.workbookObjects.values()]) workbook.close();
    this.workbookObjects.clear(); this.workbookOpens.clear();
    this.sessionByUnit.clear(); this.sessionLeases.clear();
    for (const domain of this.data.values()) domain.dispose();
    this.data.clear();
    for (const domain of this.dimensions.values()) domain.dispose();
    this.dimensions.clear();
    for (const session of this.sessions) session.dispose();
    this.sessions.clear();
    this.assetStores.clear();
    const previous = this.persistence;
    this.catalogDomain.retire();
    this.persistence = new WorkspacePersistence();
    this.api = this.createApi();
    this.catalogDomain = this.createCatalog();
    this.readiness = null;
    this.publish({ state: 'warming', error: null });
    void previous.disposeAsync();
    this.resetting = false;
  }
  createSession(resolution: WorkbookResolution): WorkbookSession {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'workbook.open', 'SDK Runtime 已释放。', '请创建新的 SDK。');
    if (typeof Worker === 'undefined') throw new SdkError('UNSUPPORTED_FEATURE', 'workbook.open', '此宿主缺少工作簿运行所需的 browser Worker。', '请提供支持 browser Worker 的浏览器宿主。');
    const existing = this.sessionByUnit.get(resolution.unitId);
    if (existing) return existing;
    const session = new WorkbookSession({
      unitId: resolution.unitId, initialPhase: 'loading', resolution, api: this.api, workspacePersistence: this.persistence,
      authTokenProvider: this.auth.createTransport().authTokenProvider, shareTokenProvider: this.auth.createTransport().shareTokenProvider,
      collaborationUrl: this.auth.collaborationUrl,
      recoverySubject: this.auth.session.getSnapshot().context?.contextId,
      pivotExecution: 'worker', assetStore: this.assetStoreFor(resolution.unitId),
      onReady: resolution.source === 'shared' ? undefined : () => this.catalog.markOpened(resolution).then(() => undefined),
    });
    this.sessions.add(session);
    this.sessionByUnit.set(resolution.unitId, session);
    this.data.set(session, new DataDomain(session));
    this.dimensions.set(session, new DimensionsDomain(session, () => session.getSelectedSheet()));
    return session;
  }
  retainSession(session: WorkbookSession): () => void {
    if (!this.sessions.has(session) || this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'workbook.retain', '会话已释放。', '请重新打开工作簿。');
    this.sessionLeases.set(session, (this.sessionLeases.get(session) ?? 0) + 1);
    let released = false;
    return () => {
      if (released) return;
      released = true;
      const count = this.sessionLeases.get(session);
      if (count === undefined) return;
      if (count > 1) this.sessionLeases.set(session, count - 1);
      else { this.sessionLeases.delete(session); this.closeSession(session); }
    };
  }
  private openWorkbook(resolution: WorkbookResolution): Promise<Workbook> {
    const existing = this.workbookOpens.get(resolution.unitId);
    if (existing) return existing;
    const releaseRuntime = this.acquire();
    let session: WorkbookSession;
    try { session = this.createSession(resolution); }
    catch (cause) { releaseRuntime(); throw cause; }
    const releaseSession = this.retainSession(session);
    let workbook!: Workbook;
    workbook = new Workbook(getWorkbookObjectPort(session), this, () => {
      if (this.workbookObjects.get(resolution.unitId) === workbook) {
        this.workbookObjects.delete(resolution.unitId); this.workbookOpens.delete(resolution.unitId);
      }
      releaseSession(); releaseRuntime();
    });
    this.workbookObjects.set(resolution.unitId, workbook);
    const opening = workbook.ready().catch(cause => { workbook.close(); throw cause; });
    this.workbookOpens.set(resolution.unitId, opening);
    session.start();
    return opening;
  }
  dataActions(session: WorkbookSession) {
    const domain = this.data.get(session);
    if (!domain) throw new SdkError('RUNTIME_DISPOSED', 'data', '数据会话不可用。', '请重新打开工作簿。');
    return domain.actions;
  }
  dimensionActions(session: WorkbookSession) {
    const domain = this.dimensions.get(session);
    if (!domain) throw new SdkError('RUNTIME_DISPOSED', 'dimensions', '行列尺寸会话不可用。', '请重新打开工作簿。');
    return domain.actions;
  }
  closeSession(session: WorkbookSession): void { this.sessionByUnit.delete(getWorkbookObjectPort(session).unitId); this.sessionLeases.delete(session); this.data.get(session)?.dispose(); this.data.delete(session); this.dimensions.get(session)?.dispose(); this.dimensions.delete(session); session.dispose(); this.sessions.delete(session); }
  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    for (const workbook of [...this.workbookObjects.values()]) workbook.close();
    this.workbookObjects.clear(); this.workbookOpens.clear();
    this.sessionByUnit.clear(); this.sessionLeases.clear();
    this.unsubscribeAuth();
    this.catalogDomain.retire();
    if (this.releaseTimer) clearTimeout(this.releaseTimer);
    for (const domain of this.data.values()) domain.dispose();
    this.data.clear();
    for (const domain of this.dimensions.values()) domain.dispose();
    this.dimensions.clear();
    for (const session of this.sessions) session.dispose();
    this.sessions.clear();
    this.assetStores.clear();
    this.listeners.clear();
    await this.persistence.disposeAsync();
  }
}
