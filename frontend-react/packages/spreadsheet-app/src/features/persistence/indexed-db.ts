/**
 * Shared IndexedDB contract for all browser-persistent workbook data.
 *
 * Every store opens this database through this module so schema creation does
 * not depend on which feature happens to initialize first.  The source bytes
 * and runtime overlays remain separate object stores; neither is part of a
 * WorkbookSnapshot or an operation envelope.
 */
export const WORKSPACE_DATABASE_NAME = 'react-sheets-workspaces';
export const WORKSPACE_DATABASE_VERSION = 4;

export const WORKSPACE_STORE_NAME = 'workspaces';
export const DATA_BLOCK_STORE_NAME = 'dataBlocks';
export const XLSX_ARTIFACT_STORE_NAME = 'xlsxArtifacts';
export const OVERLAY_STORE_NAME = 'sparseOverlays';

/** Structural factory shape also accepts the lightweight test doubles used by callers. */
export interface IndexedDbFactory {
  open(name: string, version?: number): unknown;
}

export type IndexedDbFactoryLike = IDBFactory | IndexedDbFactory;

export interface IndexedDbStoreOptions {
  databaseName?: string;
  indexedDB?: IndexedDbFactoryLike | null;
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
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB transaction failed'));
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted'));
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

/** Creates every store and index in one upgrade callback. */
export function ensureWorkspaceStores(database: IDBDatabase): void {
  if (!database.objectStoreNames.contains(WORKSPACE_STORE_NAME)) {
    database.createObjectStore(WORKSPACE_STORE_NAME, { keyPath: 'unitId' });
  }
  createSourceBlockStore(database);
  if (!database.objectStoreNames.contains(XLSX_ARTIFACT_STORE_NAME)) {
    database.createObjectStore(XLSX_ARTIFACT_STORE_NAME, { keyPath: 'unitId' });
  }
  createOverlayStore(database);
}

type DatabasePromiseMap = Map<string, Promise<IDBDatabase>>;
const databasePromises = new WeakMap<object, DatabasePromiseMap>();

/**
 * Opens the shared workbook database and guarantees the complete current
 * schema during the upgrade transaction.  Null means the runtime has no
 * IndexedDB implementation and callers may use their explicit memory path.
 */
export function openWorkspaceDatabase(options: IndexedDbStoreOptions = {}): Promise<IDBDatabase | null> {
  const factory = resolveFactory(options.indexedDB);
  if (!factory) return Promise.resolve(null);

  const name = options.databaseName ?? WORKSPACE_DATABASE_NAME;
  const factoryKey = factory as object;
  let byName = databasePromises.get(factoryKey);
  if (!byName) {
    byName = new Map<string, Promise<IDBDatabase>>();
    databasePromises.set(factoryKey, byName);
  }
  const existing = byName.get(name);
  if (existing) return existing;

  const promise = new Promise<IDBDatabase>((resolve, reject) => {
    const request = (factory as IndexedDbFactory).open(name, WORKSPACE_DATABASE_VERSION) as {
      result: IDBDatabase;
      onupgradeneeded: (() => void) | null;
      onsuccess: (() => void) | null;
      onerror: (() => void) | null;
      error?: DOMException | null;
    };
    request.onupgradeneeded = () => ensureWorkspaceStores(request.result);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB open failed'));
  });
  byName.set(name, promise);
  return promise;
}

export function resolveIndexedDbFactory(explicit: IndexedDbFactoryLike | null | undefined): IndexedDbFactoryLike | null {
  return resolveFactory(explicit);
}
