/**
 * Shared IndexedDB contract for all browser-persistent workbook data.
 *
 * Every store opens this database through this module so schema creation does
 * not depend on which feature happens to initialize first.  The source bytes
 * and runtime overlays remain separate object stores; neither is part of a
 * WorkbookSnapshot or an operation envelope.
 */
export const WORKSPACE_DATABASE_NAME = 'react-sheets-workspaces';
export const WORKSPACE_DATABASE_VERSION = 7;

export const WORKSPACE_STORE_NAME = 'workspaces';
export const DATA_BLOCK_STORE_NAME = 'dataBlocks';
export const NATIVE_PACKAGE_STORE_NAME = 'nativePackages';
export const OVERLAY_STORE_NAME = 'sparseOverlays';
export const ASSET_STORE_NAME = 'assets';

/** Structural factory shape also accepts the lightweight test doubles used by callers. */
export interface IndexedDbFactory {
  open(name: string, version?: number): unknown;
}

export type IndexedDbFactoryLike = IDBFactory | IndexedDbFactory;

export type WorkspaceStorageErrorCode =
  | 'STORAGE_UNAVAILABLE'
  | 'STORAGE_UPGRADE_BLOCKED'
  | 'STORAGE_OPEN_TIMEOUT'
  | 'STORAGE_SCHEMA_INVALID'
  | 'STORAGE_TRANSACTION_FAILED';

export class WorkspaceStorageError extends Error {
  readonly code: WorkspaceStorageErrorCode;
  readonly databaseName: string;
  readonly targetVersion: number;
  readonly operation: string;
  readonly recovery: string;

  constructor(input: {
    code: WorkspaceStorageErrorCode;
    databaseName: string;
    operation: string;
    message: string;
    recovery: string;
    cause?: unknown;
  }) {
    super(input.message, input.cause === undefined ? undefined : { cause: input.cause });
    this.name = 'WorkspaceStorageError';
    this.code = input.code;
    this.databaseName = input.databaseName;
    this.targetVersion = WORKSPACE_DATABASE_VERSION;
    this.operation = input.operation;
    this.recovery = input.recovery;
  }
}

export interface IndexedDbStoreOptions {
  databaseName?: string;
  indexedDB?: IndexedDbFactoryLike | null;
  coordinator?: WorkspaceDatabaseCoordinator;
  /** Optional owner namespace for source blocks/overlays. */
  unitId?: string | (() => string);
}

/**
 * Browser-persistent source records are keyed by workbook as well as their
 * source id.  This prevents two workbooks that happen to use the same source
 * id from sharing bytes or overlays.  Existing callers without a namespace
 * retain their legacy key so old focused stores remain readable.
 */
export function resolveWorkspaceUnitId(options: Pick<IndexedDbStoreOptions, 'unitId'> = {}): string | null {
  const value = typeof options.unitId === 'function' ? options.unitId() : options.unitId;
  return value?.trim() || null;
}

export function namespaceWorkspaceSourceId(unitId: string | null | undefined, sourceId: string): string {
  const normalizedUnitId = unitId?.trim();
  if (!normalizedUnitId || sourceId.startsWith(`unit:${normalizedUnitId}:source:`)) return sourceId;
  return `unit:${normalizedUnitId}:source:${sourceId}`;
}

export function unnamespaceWorkspaceSourceId(sourceId: string): string {
  const marker = ':source:';
  const markerIndex = sourceId.indexOf(marker);
  return markerIndex >= 0 && sourceId.startsWith('unit:')
    ? sourceId.slice(markerIndex + marker.length)
    : sourceId;
}

function resolveFactory(explicit: IndexedDbFactoryLike | null | undefined): IndexedDbFactoryLike | null {
  if (explicit !== undefined) return explicit;
  return typeof globalThis !== 'undefined' && 'indexedDB' in globalThis
    ? (globalThis as typeof globalThis & { indexedDB?: IDBFactory }).indexedDB ?? null
    : null;
}

export function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'));
  });
}

export function transactionComplete(transaction: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onComplete = () => resolve();
    const onError = () => reject(transaction.error ?? new Error('IndexedDB transaction failed'));
    const onAbort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted'));
    if ('addEventListener' in transaction && typeof transaction.addEventListener === 'function') {
      transaction.addEventListener('complete', onComplete, { once: true });
      transaction.addEventListener('error', onError, { once: true });
      transaction.addEventListener('abort', onAbort, { once: true });
    } else {
      transaction.oncomplete = onComplete;
      transaction.onerror = onError;
      transaction.onabort = onAbort;
    }
  });
}

function createSourceBlockStore(database: IDBDatabase): void {
  if (database.objectStoreNames.contains(DATA_BLOCK_STORE_NAME)) return;
  const store = database.createObjectStore(DATA_BLOCK_STORE_NAME, { keyPath: ['sourceId', 'blockId'] }) as IDBObjectStore | undefined;
  store?.createIndex('sourceId', 'sourceId', { unique: false });
}

function createOverlayStore(database: IDBDatabase): void {
  if (database.objectStoreNames.contains(OVERLAY_STORE_NAME)) return;
  const store = database.createObjectStore(OVERLAY_STORE_NAME, { keyPath: ['sourceId', 'blockId', 'revision'] }) as IDBObjectStore | undefined;
  store?.createIndex('sourceBlock', ['sourceId', 'blockId'], { unique: false });
  store?.createIndex('sourceId', 'sourceId', { unique: false });
}

function createAssetStore(database: IDBDatabase): void {
  if (database.objectStoreNames.contains(ASSET_STORE_NAME)) return;
  const store = database.createObjectStore(ASSET_STORE_NAME, { keyPath: ['unitId', 'assetId'] }) as IDBObjectStore | undefined;
  store?.createIndex('contentHash', ['unitId', 'contentHash'], { unique: false });
}

/** Creates every store and index in one upgrade callback. */
export function ensureWorkspaceStores(database: IDBDatabase): void {
  if (!database.objectStoreNames.contains(WORKSPACE_STORE_NAME)) {
    database.createObjectStore(WORKSPACE_STORE_NAME, { keyPath: 'unitId' });
  }
  createSourceBlockStore(database);
  if (!database.objectStoreNames.contains(NATIVE_PACKAGE_STORE_NAME)) {
    database.createObjectStore(NATIVE_PACKAGE_STORE_NAME, { keyPath: 'unitId' });
  }
  createOverlayStore(database);
  createAssetStore(database);
}

interface IndexedDbOpenRequest {
  result: IDBDatabase;
  onupgradeneeded: (() => void) | null;
  onsuccess: (() => void) | null;
  onerror: (() => void) | null;
  onblocked: (() => void) | null;
  error?: DOMException | null;
}

const WORKSPACE_DATABASE_OPEN_TIMEOUT_MS = 15_000;
const WORKSPACE_DATABASE_CHANNEL_PREFIX = 'react-sheets-workspace-db';

export type WorkspaceDatabaseState = 'idle' | 'opening' | 'ready' | 'closing' | 'failed';

interface WorkspaceDatabaseMessage {
  type: 'close-request' | 'closed';
  databaseName: string;
  targetVersion: number;
  instanceId: string;
}

function storageError(
  code: WorkspaceStorageErrorCode,
  databaseName: string,
  operation: string,
  message: string,
  recovery: string,
  cause?: unknown,
): WorkspaceStorageError {
  return new WorkspaceStorageError({ code, databaseName, operation, message, recovery, cause });
}

function assertWorkspaceSchema(database: IDBDatabase, databaseName: string): void {
  const required = [WORKSPACE_STORE_NAME, DATA_BLOCK_STORE_NAME, NATIVE_PACKAGE_STORE_NAME, OVERLAY_STORE_NAME, ASSET_STORE_NAME];
  const missing = required.filter((store) => !database.objectStoreNames.contains(store));
  if (missing.length > 0) {
    throw storageError(
      'STORAGE_SCHEMA_INVALID',
      databaseName,
      'open',
      `本地工作簿存储结构不完整：${missing.join(', ')}`,
      '请保留当前页面并联系管理员检查数据库迁移；不要清除浏览器数据。',
    );
  }
}

/**
 * The only owner of the browser workspace database connection. All durable
 * stores share this lifecycle and must acquire transactions through it.
 */
export class WorkspaceDatabaseCoordinator {
  readonly databaseName: string;
  readonly instanceId: string;
  private readonly factory: IndexedDbFactoryLike | null;
  private readonly channel: BroadcastChannel | null;
  private readonly openTimeoutMs: number;
  private database: IDBDatabase | null = null;
  private opening: Promise<IDBDatabase> | null = null;
  private closing: Promise<void> | null = null;
  private cancelOpening: ((error: WorkspaceStorageError) => void) | null = null;
  private stateValue: WorkspaceDatabaseState = 'idle';
  private activeTransactions = 0;
  private drainWaiters: Array<() => void> = [];

  constructor(options: Omit<IndexedDbStoreOptions, 'coordinator' | 'unitId'> & { openTimeoutMs?: number; broadcast?: boolean } = {}) {
    this.databaseName = options.databaseName ?? WORKSPACE_DATABASE_NAME;
    this.factory = resolveFactory(options.indexedDB);
    this.openTimeoutMs = options.openTimeoutMs ?? WORKSPACE_DATABASE_OPEN_TIMEOUT_MS;
    this.instanceId = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    this.channel = options.broadcast !== false && typeof BroadcastChannel === 'function'
      ? new BroadcastChannel(`${WORKSPACE_DATABASE_CHANNEL_PREFIX}:${this.databaseName}`)
      : null;
    if (this.channel) this.channel.onmessage = (event: MessageEvent<WorkspaceDatabaseMessage>) => this.receive(event.data);
    if (typeof globalThis.addEventListener === 'function') {
      globalThis.addEventListener('pagehide', this.handlePageHide, { capture: true });
    }
  }

  get state(): WorkspaceDatabaseState { return this.stateValue; }

  open(): Promise<IDBDatabase> {
    if (this.closing) return this.closing.then(() => this.open());
    if (this.database && this.stateValue === 'ready') return Promise.resolve(this.database);
    if (this.opening) return this.opening;
    if (!this.factory) {
      this.stateValue = 'failed';
      return Promise.reject(storageError(
        'STORAGE_UNAVAILABLE',
        this.databaseName,
        'open',
        '当前环境不支持本地工作簿存储。',
        '请使用支持 IndexedDB 的浏览器并允许本地站点存储。',
      ));
    }

    this.stateValue = 'opening';
    this.opening = new Promise<IDBDatabase>((resolve, reject) => {
      let settled = false;
      let blocked = false;
      const request = (this.factory as IndexedDbFactory).open(this.databaseName, WORKSPACE_DATABASE_VERSION) as IndexedDbOpenRequest;
      const finishFailure = (error: WorkspaceStorageError): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timeoutId);
        if (this.cancelOpening === finishFailure) this.cancelOpening = null;
        this.opening = null;
        this.stateValue = 'failed';
        reject(error);
      };
      const timeoutId = setTimeout(() => finishFailure(blocked
        ? storageError(
          'STORAGE_UPGRADE_BLOCKED',
          this.databaseName,
          'upgrade',
          '本地工作簿存储升级仍被旧页面占用。',
          '请保存其他工作簿页面中的内容并完整刷新或关闭旧页面，然后重试；现有数据库不会被删除。',
        )
        : storageError(
          'STORAGE_OPEN_TIMEOUT',
          this.databaseName,
          'open',
          '打开本地工作簿存储超时。',
          '请完整刷新当前页面后重试；若问题持续，请检查浏览器站点存储权限。',
        )), this.openTimeoutMs);
      this.cancelOpening = finishFailure;

      request.onupgradeneeded = () => ensureWorkspaceStores(request.result);
      request.onblocked = () => {
        blocked = true;
        this.channel?.postMessage({
          type: 'close-request',
          databaseName: this.databaseName,
          targetVersion: WORKSPACE_DATABASE_VERSION,
          instanceId: this.instanceId,
        } satisfies WorkspaceDatabaseMessage);
      };
      request.onerror = () => finishFailure(storageError(
        'STORAGE_TRANSACTION_FAILED',
        this.databaseName,
        'open',
        '无法打开本地工作簿存储。',
        '请保留浏览器数据并重试；错误详情可用于管理员诊断。',
        request.error,
      ));
      request.onsuccess = () => {
        if (settled) {
          request.result.close();
          return;
        }
        try {
          assertWorkspaceSchema(request.result, this.databaseName);
        } catch (error) {
          request.result.close();
          finishFailure(error instanceof WorkspaceStorageError ? error : storageError(
            'STORAGE_SCHEMA_INVALID', this.databaseName, 'open', '本地工作簿存储结构无效。', '请联系管理员检查数据库迁移。', error,
          ));
          return;
        }
        settled = true;
        clearTimeout(timeoutId);
        if (this.cancelOpening === finishFailure) this.cancelOpening = null;
        this.database = request.result;
        this.database.onversionchange = () => { void this.close('versionchange'); };
        this.opening = null;
        this.stateValue = 'ready';
        resolve(request.result);
      };
    });
    return this.opening;
  }

  async transaction(storeNames: string | string[], mode: IDBTransactionMode): Promise<IDBTransaction> {
    const database = await this.open();
    if (this.stateValue !== 'ready') {
      throw storageError('STORAGE_TRANSACTION_FAILED', this.databaseName, 'transaction', '本地工作簿存储尚未就绪。', '请等待存储就绪后重试。');
    }
    let transaction: IDBTransaction;
    try {
      transaction = database.transaction(storeNames, mode);
    } catch (error) {
      throw storageError('STORAGE_TRANSACTION_FAILED', this.databaseName, 'transaction', '无法开始本地工作簿事务。', '请重试当前操作；若问题持续，请保留错误详情。', error);
    }
    this.activeTransactions += 1;
    let released = false;
    const release = (): void => {
      if (released) return;
      released = true;
      this.activeTransactions -= 1;
      if (this.activeTransactions === 0) {
        const waiters = this.drainWaiters;
        this.drainWaiters = [];
        waiters.forEach((resolve) => resolve());
      }
    };
    if ('addEventListener' in transaction && typeof transaction.addEventListener === 'function') {
      transaction.addEventListener('complete', release, { once: true });
      transaction.addEventListener('abort', release, { once: true });
      transaction.addEventListener('error', release, { once: true });
    }
    return transaction;
  }

  close(reason: 'dispose' | 'pagehide' | 'versionchange' | 'upgrade-request' = 'dispose'): Promise<void> {
    if (this.closing) return this.closing;
    this.cancelOpening?.(storageError(
      'STORAGE_UNAVAILABLE',
      this.databaseName,
      'close',
      '本地工作簿存储在打开完成前被页面生命周期关闭。',
      '请重新进入工作簿页面后重试当前操作。',
    ));
    this.cancelOpening = null;
    this.stateValue = 'closing';
    const closing = Promise.resolve().then(async () => {
      if (this.activeTransactions > 0) await new Promise<void>((resolve) => this.drainWaiters.push(resolve));
      this.database?.close();
      this.database = null;
      this.opening = null;
      this.stateValue = 'idle';
      if (reason === 'upgrade-request') {
        this.channel?.postMessage({
          type: 'closed', databaseName: this.databaseName, targetVersion: WORKSPACE_DATABASE_VERSION, instanceId: this.instanceId,
        } satisfies WorkspaceDatabaseMessage);
      }
    }).finally(() => {
      if (this.closing === closing) this.closing = null;
    });
    this.closing = closing;
    return closing;
  }

  dispose(): void {
    void this.close('dispose');
    this.channel?.close();
    if (typeof globalThis.removeEventListener === 'function') {
      globalThis.removeEventListener('pagehide', this.handlePageHide, { capture: true });
    }
  }

  private readonly handlePageHide = (): void => { void this.close('pagehide'); };

  private receive(message: WorkspaceDatabaseMessage): void {
    if (!message || message.databaseName !== this.databaseName || message.instanceId === this.instanceId) return;
    if (message.type === 'close-request' && message.targetVersion >= WORKSPACE_DATABASE_VERSION) {
      void this.close('upgrade-request');
    }
  }
}

type CoordinatorMap = Map<string, WorkspaceDatabaseCoordinator>;
const coordinatorByFactory = new WeakMap<object, CoordinatorMap>();
const ownedCoordinators = new Set<WorkspaceDatabaseCoordinator>();

export function resolveWorkspaceDatabaseCoordinator(options: IndexedDbStoreOptions = {}): WorkspaceDatabaseCoordinator {
  if (options.coordinator) return options.coordinator;
  const factory = resolveFactory(options.indexedDB);
  if (!factory) return new WorkspaceDatabaseCoordinator(options);
  let byName = coordinatorByFactory.get(factory as object);
  if (!byName) {
    byName = new Map();
    coordinatorByFactory.set(factory as object, byName);
  }
  const name = options.databaseName ?? WORKSPACE_DATABASE_NAME;
  let coordinator = byName.get(name);
  if (!coordinator) {
    coordinator = new WorkspaceDatabaseCoordinator({ databaseName: name, indexedDB: factory });
    byName.set(name, coordinator);
    ownedCoordinators.add(coordinator);
  }
  return coordinator;
}

const hotModule = (import.meta as ImportMeta & { hot?: { dispose(callback: () => void): void } }).hot;
hotModule?.dispose(() => {
  for (const coordinator of ownedCoordinators) coordinator.dispose();
  ownedCoordinators.clear();
});

export function resolveIndexedDbFactory(explicit: IndexedDbFactoryLike | null | undefined): IndexedDbFactoryLike | null {
  return resolveFactory(explicit);
}
