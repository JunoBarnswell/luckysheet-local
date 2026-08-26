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

export type WorkspaceStorageMode = 'persistent' | 'memory' | 'unavailable';

interface FactorySelection {
  factory: IndexedDbFactoryLike | null;
  mode: WorkspaceStorageMode;
  allowMemoryFallback: boolean;
}

let inMemoryIndexedDb: IDBFactory | null = null;

function memoryFactory(): IDBFactory {
  inMemoryIndexedDb ??= new InMemoryIndexedDbFactory();
  return inMemoryIndexedDb;
}

function isBrowserRuntime(): boolean {
  return typeof window !== 'undefined' && typeof document !== 'undefined';
}

function resolveFactorySelection(explicit: IndexedDbFactoryLike | null | undefined): FactorySelection {
  if (explicit !== undefined) return { factory: explicit, mode: explicit ? 'persistent' : 'unavailable', allowMemoryFallback: false };
  if (!isBrowserRuntime()) return { factory: null, mode: 'unavailable', allowMemoryFallback: false };
  try {
    const nativeFactory = (globalThis as typeof globalThis & { indexedDB?: IDBFactory }).indexedDB;
    if (nativeFactory) return { factory: nativeFactory, mode: 'persistent', allowMemoryFallback: true };
  } catch {
    // Access itself may throw SecurityError in a restricted browser context.
  }
  return { factory: memoryFactory(), mode: 'memory', allowMemoryFallback: false };
}

function resolveFactory(explicit: IndexedDbFactoryLike | null | undefined): IndexedDbFactoryLike | null {
  return resolveFactorySelection(explicit).factory;
}

function isUnsupportedIndexedDbError(cause: unknown): boolean {
  if (!cause || typeof cause !== 'object') return false;
  const name = 'name' in cause ? String(cause.name) : '';
  return name === 'NotSupportedError' || name === 'SecurityError';
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
const WORKSPACE_DATABASE_CLOSE_DRAIN_TIMEOUT_MS = 5_000;
const WORKSPACE_DATABASE_CHANNEL_PREFIX = 'react-sheets-workspace-db';

export type WorkspaceDatabaseState = 'idle' | 'opening' | 'ready' | 'closing' | 'failed';

export interface WorkspaceDatabaseMessage {
  type: 'close-request' | 'closed';
  databaseName: string;
  targetVersion: number;
  instanceId: string;
}

interface OpeningAttempt {
  generation: number;
  blocked: boolean;
  peerClosed: boolean;
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
  private factory: IndexedDbFactoryLike | null;
  private storageModeValue: WorkspaceStorageMode;
  private allowMemoryFallback: boolean;
  private readonly channel: BroadcastChannel | null;
  private readonly openTimeoutMs: number;
  private readonly closeDrainTimeoutMs: number;
  private database: IDBDatabase | null = null;
  private opening: Promise<IDBDatabase> | null = null;
  private closing: Promise<void> | null = null;
  private cancelOpening: ((error: WorkspaceStorageError) => void) | null = null;
  private stateValue: WorkspaceDatabaseState = 'idle';
  private activeTransactions = 0;
  private drainWaiters: Array<() => void> = [];
  private openGeneration = 0;
  private openingAttempt: OpeningAttempt | null = null;

  constructor(options: Omit<IndexedDbStoreOptions, 'coordinator' | 'unitId'> & {
    openTimeoutMs?: number;
    closeDrainTimeoutMs?: number;
    broadcast?: boolean;
  } = {}) {
    this.databaseName = options.databaseName ?? WORKSPACE_DATABASE_NAME;
    const selection = resolveFactorySelection(options.indexedDB);
    this.factory = selection.factory;
    this.storageModeValue = selection.mode;
    this.allowMemoryFallback = selection.allowMemoryFallback;
    this.openTimeoutMs = options.openTimeoutMs ?? WORKSPACE_DATABASE_OPEN_TIMEOUT_MS;
    this.closeDrainTimeoutMs = options.closeDrainTimeoutMs ?? WORKSPACE_DATABASE_CLOSE_DRAIN_TIMEOUT_MS;
    this.instanceId = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    this.channel = options.broadcast !== false && selection.mode === 'persistent' && typeof BroadcastChannel === 'function'
      ? new BroadcastChannel(`${WORKSPACE_DATABASE_CHANNEL_PREFIX}:${this.databaseName}`)
      : null;
    if (this.channel) this.channel.onmessage = (event: MessageEvent<WorkspaceDatabaseMessage>) => this.receive(event.data);
    if (typeof globalThis.addEventListener === 'function') {
      globalThis.addEventListener('pagehide', this.handlePageHide, { capture: true });
    }
  }

  get state(): WorkspaceDatabaseState { return this.stateValue; }
  get storageMode(): WorkspaceStorageMode { return this.storageModeValue; }

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

    const generation = ++this.openGeneration;
    const attempt: OpeningAttempt = { generation, blocked: false, peerClosed: false };
    this.openingAttempt = attempt;
    this.stateValue = 'opening';
    this.opening = new Promise<IDBDatabase>((resolve, reject) => {
      let settled = false;
      let request: IndexedDbOpenRequest;
      try {
        request = (this.factory as IndexedDbFactory).open(this.databaseName, WORKSPACE_DATABASE_VERSION) as IndexedDbOpenRequest;
      } catch (error) {
        if (this.isCurrentGeneration(generation) && this.switchToMemoryIfUnsupported(error)) {
          this.clearOpening(generation);
          this.stateValue = 'idle';
          void this.open().then(resolve, reject);
          return;
        }
        this.failOpening(generation, storageError(
          'STORAGE_UNAVAILABLE',
          this.databaseName,
          'open',
          '当前浏览器拒绝本地工作簿存储。',
          '请允许站点存储后重试。',
          error,
        ), reject);
        return;
      }

      const finishFailure = (error: WorkspaceStorageError): void => {
        if (settled || !this.isCurrentGeneration(generation)) return;
        settled = true;
        clearTimeout(timeoutId);
        this.failOpening(generation, error, reject);
      };

      const timeoutId = setTimeout(() => finishFailure(this.openTimeoutError(attempt)), this.openTimeoutMs);
      this.cancelOpening = finishFailure;

      request.onupgradeneeded = () => {
        if (!this.isCurrentGeneration(generation)) {
          try { request.result.close(); } catch { /* ignore zombie upgrade handle */ }
          return;
        }
        ensureWorkspaceStores(request.result);
      };
      request.onblocked = () => {
        if (!this.isCurrentGeneration(generation)) return;
        attempt.blocked = true;
        this.channel?.postMessage({
          type: 'close-request',
          databaseName: this.databaseName,
          targetVersion: WORKSPACE_DATABASE_VERSION,
          instanceId: this.instanceId,
        } satisfies WorkspaceDatabaseMessage);
      };
      request.onerror = () => {
        if (!this.isCurrentGeneration(generation)) return;
        if (this.switchToMemoryIfUnsupported(request.error)) {
          if (settled) return;
          settled = true;
          clearTimeout(timeoutId);
          this.clearOpening(generation);
          this.stateValue = 'idle';
          void this.open().then(resolve, reject);
          return;
        }
        finishFailure(storageError(
          'STORAGE_TRANSACTION_FAILED',
          this.databaseName,
          'open',
          '无法打开本地工作簿存储。',
          '请保留浏览器数据并重试；错误详情可用于管理员诊断。',
          request.error,
        ));
      };
      request.onsuccess = () => {
        if (settled || !this.isCurrentGeneration(generation)) {
          try { request.result.close(); } catch { /* ignore zombie connection */ }
          return;
        }
        try {
          assertWorkspaceSchema(request.result, this.databaseName);
        } catch (error) {
          try { request.result.close(); } catch { /* ignore invalid schema handle */ }
          finishFailure(error instanceof WorkspaceStorageError ? error : storageError(
            'STORAGE_SCHEMA_INVALID', this.databaseName, 'open', '本地工作簿存储结构无效。', '请联系管理员检查数据库迁移。', error,
          ));
          return;
        }
        settled = true;
        clearTimeout(timeoutId);
        if (this.cancelOpening === finishFailure) this.cancelOpening = null;
        this.clearOpening(generation);
        this.database = request.result;
        this.database.onversionchange = () => { void this.close('versionchange'); };
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
    } else {
      const previousComplete = transaction.oncomplete;
      const previousAbort = transaction.onabort;
      const previousError = transaction.onerror;
      transaction.oncomplete = (event) => {
        release();
        if (typeof previousComplete === 'function') previousComplete.call(transaction, event);
      };
      transaction.onabort = (event) => {
        release();
        if (typeof previousAbort === 'function') previousAbort.call(transaction, event);
      };
      transaction.onerror = (event) => {
        release();
        if (typeof previousError === 'function') previousError.call(transaction, event);
      };
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
    this.openGeneration += 1;
    this.openingAttempt = null;
    this.stateValue = 'closing';
    const closing = Promise.resolve().then(async () => {
      if (this.activeTransactions > 0) {
        await Promise.race([
          new Promise<void>((resolve) => this.drainWaiters.push(resolve)),
          new Promise<void>((resolve) => {
            setTimeout(resolve, this.closeDrainTimeoutMs);
          }),
        ]);
        if (this.activeTransactions > 0) {
          this.activeTransactions = 0;
          const waiters = this.drainWaiters;
          this.drainWaiters = [];
          waiters.forEach((resolve) => resolve());
        }
      }
      try { this.database?.close(); } catch { /* ignore already-closed handle */ }
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
    void this.disposeAsync();
  }

  async disposeAsync(): Promise<void> {
    try {
      await this.close('dispose');
    } finally {
      try { this.channel?.close(); } catch { /* ignore channel teardown */ }
      if (typeof globalThis.removeEventListener === 'function') {
        globalThis.removeEventListener('pagehide', this.handlePageHide, { capture: true });
      }
    }
  }

  private readonly handlePageHide = (): void => { void this.close('pagehide'); };

  private receive(message: WorkspaceDatabaseMessage): void {
    if (!message || message.databaseName !== this.databaseName || message.instanceId === this.instanceId) return;
    if (message.type === 'close-request' && message.targetVersion >= WORKSPACE_DATABASE_VERSION) {
      void this.close('upgrade-request');
      return;
    }
    if (message.type === 'closed') {
      const attempt = this.openingAttempt;
      if (!attempt || attempt.generation !== this.openGeneration) return;
      attempt.peerClosed = true;
      // Same-generation open continues waiting for onsuccess; peer close only records protocol progress.
    }
  }

  private isCurrentGeneration(generation: number): boolean {
    return generation === this.openGeneration && this.openingAttempt?.generation === generation;
  }

  private clearOpening(generation: number): void {
    if (this.openingAttempt?.generation === generation) this.openingAttempt = null;
    this.opening = null;
    if (this.cancelOpening) this.cancelOpening = null;
  }

  private failOpening(generation: number, error: WorkspaceStorageError, reject: (error: WorkspaceStorageError) => void): void {
    if (this.isCurrentGeneration(generation)) {
      if (this.cancelOpening) this.cancelOpening = null;
      this.clearOpening(generation);
      this.stateValue = 'failed';
    }
    reject(error);
  }

  /** Test seam: deliver a peer BroadcastChannel message without a live channel. */
  deliverPeerMessageForTest(message: WorkspaceDatabaseMessage): void {
    this.receive(message);
  }

  private openTimeoutError(attempt: OpeningAttempt): WorkspaceStorageError {
    if (attempt.blocked) {
      return storageError(
        'STORAGE_UPGRADE_BLOCKED',
        this.databaseName,
        'upgrade',
        '本地工作簿存储升级仍被旧页面占用。',
        '请保存其他工作簿页面中的内容并完整刷新或关闭旧页面，然后重试；现有数据库不会被删除。',
      );
    }
    return storageError(
      'STORAGE_OPEN_TIMEOUT',
      this.databaseName,
      'open',
      '打开本地工作簿存储超时。',
      '请关闭其它标签页中的本应用后完整刷新当前页面并重试；若问题持续，请检查浏览器站点存储权限。',
    );
  }

  private switchToMemoryIfUnsupported(cause: unknown): boolean {
    if (!this.allowMemoryFallback || !isUnsupportedIndexedDbError(cause)) return false;
    this.channel?.close();
    this.factory = memoryFactory();
    this.storageModeValue = 'memory';
    this.allowMemoryFallback = false;
    return true;
  }
}

type CoordinatorMap = Map<string, WorkspaceDatabaseCoordinator>;
const coordinatorByFactory = new WeakMap<object, CoordinatorMap>();
const ownedCoordinators = new Set<WorkspaceDatabaseCoordinator>();

export function resolveWorkspaceDatabaseCoordinator(options: IndexedDbStoreOptions = {}): WorkspaceDatabaseCoordinator {
  if (options.coordinator) return options.coordinator;
  const selection = resolveFactorySelection(options.indexedDB);
  const factory = selection.factory;
  if (!factory) return new WorkspaceDatabaseCoordinator(options);
  let byName = coordinatorByFactory.get(factory as object);
  if (!byName) {
    byName = new Map();
    coordinatorByFactory.set(factory as object, byName);
  }
  const name = options.databaseName ?? WORKSPACE_DATABASE_NAME;
  let coordinator = byName.get(name);
  if (!coordinator) {
    coordinator = selection.mode === 'memory'
      ? new WorkspaceDatabaseCoordinator({ databaseName: name, broadcast: false })
      : new WorkspaceDatabaseCoordinator({ databaseName: name, indexedDB: factory, broadcast: selection.mode === 'persistent' });
    byName.set(name, coordinator);
    ownedCoordinators.add(coordinator);
  }
  return coordinator;
}

/** Used by Vite HMR dispose and tests to close every owned coordinator before a new module opens the DB. */
export async function disposeOwnedWorkspaceCoordinators(): Promise<void> {
  const pending = [...ownedCoordinators];
  ownedCoordinators.clear();
  await Promise.all(pending.map((coordinator) => coordinator.disposeAsync()));
}

const hotModule = (import.meta as ImportMeta & { hot?: { dispose(callback: () => void | Promise<void>): void } }).hot;
hotModule?.dispose(() => disposeOwnedWorkspaceCoordinators());

export function resolveIndexedDbFactory(explicit: IndexedDbFactoryLike | null | undefined): IndexedDbFactoryLike | null {
  return resolveFactory(explicit);
}
import { IDBFactory as InMemoryIndexedDbFactory } from 'fake-indexeddb';
