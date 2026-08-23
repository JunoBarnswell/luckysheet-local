import { computeChecksum } from '../persistence/checksum';
import {
  openWorkspaceDatabase,
  requestResult,
  resolveIndexedDbFactory,
  transactionComplete,
  OVERLAY_STORE_NAME,
  WORKSPACE_DATABASE_NAME,
  type IndexedDbFactoryLike,
} from '../persistence/indexed-db';
import type { SparseCellOverlayMetadata, SparseCellOverlayMetadataCell } from './import';

const DEFAULT_DATABASE_NAME = WORKSPACE_DATABASE_NAME;

export interface SparseOverlayRecord {
  schema: 'SparseCellOverlayRecord';
  sourceId: string;
  blockId: string;
  revision: number;
  overlay: SparseCellOverlayMetadata;
  checksum: string;
  updatedAt: string;
}

export interface SparseOverlayStore {
  put(sourceId: string, blockId: string, overlay: SparseCellOverlayMetadata): Promise<SparseOverlayRecord>;
  get(sourceId: string, blockId: string, revision: number): Promise<SparseOverlayRecord | null>;
  remove(sourceId: string, blockId: string, revision: number): Promise<void>;
  removeBlock(sourceId: string, blockId: string): Promise<void>;
  removeSource(sourceId: string): Promise<void>;
}

export interface SparseOverlayStoreOptions {
  databaseName?: string;
  indexedDB?: IndexedDbFactoryLike | null;
}

const memoryDatabases = new Map<string, Map<string, SparseOverlayRecord>>();

function fail(message: string): never {
  throw new Error(`Sparse overlay store ${message}`);
}

function isSafeRevision(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function recordKey(sourceId: string, blockId: string, revision: number): string {
  return `${sourceId}\u0000${blockId}\u0000${String(revision)}`;
}

function blockPrefix(sourceId: string, blockId: string): string {
  return `${sourceId}\u0000${blockId}\u0000`;
}

function sourcePrefix(sourceId: string): string {
  return `${sourceId}\u0000`;
}

function clone<T>(value: T): T {
  return structuredClone(value);
}

function overlayPayload(overlay: SparseCellOverlayMetadata): string {
  return JSON.stringify({
    schema: overlay.schema,
    revision: overlay.revision,
    cells: overlay.cells,
  });
}

function validateCell(cell: SparseCellOverlayMetadataCell, index: number): void {
  if (!isRecord(cell)
    || !Number.isSafeInteger(cell.row)
    || Number(cell.row) < 0
    || !Number.isSafeInteger(cell.column)
    || Number(cell.column) < 0) {
    fail(`contains an invalid cell coordinate at index ${String(index)}`);
  }
  if (cell.formula !== undefined && typeof cell.formula !== 'string') {
    fail(`contains an invalid formula at index ${String(index)}`);
  }
  if (cell.style !== undefined && !isRecord(cell.style)) {
    fail(`contains an invalid style at index ${String(index)}`);
  }
  if (cell.comment !== undefined && !isRecord(cell.comment)) {
    fail(`contains an invalid comment at index ${String(index)}`);
  }
  if (cell.formula === undefined && cell.style === undefined && cell.comment === undefined) {
    fail(`contains an empty metadata cell at index ${String(index)}`);
  }
}

function normalizeOverlay(overlay: SparseCellOverlayMetadata, revision: number): SparseCellOverlayMetadata {
  if (!isRecord(overlay) || overlay.schema !== 'SparseCellOverlayMetadata' || overlay.revision !== revision || !isSafeRevision(overlay.revision) || !Array.isArray(overlay.cells)) {
    fail('has an invalid schema or revision');
  }
  const coordinates = new Set<string>();
  const cells = overlay.cells.map((cell, index) => {
    validateCell(cell, index);
    const key = `${String(cell.row)}:${String(cell.column)}`;
    if (coordinates.has(key)) fail(`contains duplicate cell ${key}`);
    coordinates.add(key);
    return clone(cell);
  }).sort((left, right) => left.row - right.row || left.column - right.column);
  return {
    schema: 'SparseCellOverlayMetadata',
    revision,
    cells,
  };
}

function validateIdentity(sourceId: string, blockId: string, revision: number): void {
  if (!sourceId.trim()) fail('requires a source id');
  if (!blockId.trim()) fail('requires a block id');
  if (!isSafeRevision(revision)) fail('revision must be a non-negative safe integer');
}

function buildRecord(
  sourceId: string,
  blockId: string,
  overlay: SparseCellOverlayMetadata,
): SparseOverlayRecord {
  const normalized = normalizeOverlay(overlay, overlay.revision);
  const record: SparseOverlayRecord = {
    schema: 'SparseCellOverlayRecord',
    sourceId,
    blockId,
    revision: normalized.revision,
    overlay: normalized,
    checksum: computeChecksum(overlayPayload(normalized)),
    updatedAt: new Date().toISOString(),
  };
  return clone(record);
}

function validateRecord(record: SparseOverlayRecord): SparseOverlayRecord {
  if (!isRecord(record)
    || record.schema !== 'SparseCellOverlayRecord'
    || typeof record.sourceId !== 'string'
    || !record.sourceId.trim()
    || typeof record.blockId !== 'string'
    || !record.blockId.trim()
    || !isSafeRevision(record.revision)
    || typeof record.updatedAt !== 'string'
    || !record.updatedAt
    || typeof record.checksum !== 'string'
    || !record.checksum) {
    fail('contains an invalid record');
  }
  const overlay = normalizeOverlay(record.overlay, record.revision);
  if (computeChecksum(overlayPayload(overlay)) !== record.checksum) fail(`checksum mismatch for ${record.blockId}`);
  return {
    ...clone(record),
    overlay,
  };
}

function memoryRecords(databaseName: string): Map<string, SparseOverlayRecord> {
  let records = memoryDatabases.get(databaseName);
  if (!records) {
    records = new Map<string, SparseOverlayRecord>();
    memoryDatabases.set(databaseName, records);
  }
  return records;
}

/** In-memory implementation used by tests and non-browser runtimes. */
export class MemorySparseOverlayStore implements SparseOverlayStore {
  private readonly records = new Map<string, SparseOverlayRecord>();

  async put(sourceId: string, blockId: string, overlay: SparseCellOverlayMetadata): Promise<SparseOverlayRecord> {
    validateIdentity(sourceId, blockId, overlay?.revision);
    const record = buildRecord(sourceId, blockId, overlay);
    this.records.set(recordKey(sourceId, blockId, record.revision), clone(record));
    return clone(record);
  }

  async get(sourceId: string, blockId: string, revision: number): Promise<SparseOverlayRecord | null> {
    validateIdentity(sourceId, blockId, revision);
    const value = this.records.get(recordKey(sourceId, blockId, revision));
    if (!value) return null;
    try {
      return validateRecord(clone(value));
    } catch {
      this.records.delete(recordKey(sourceId, blockId, revision));
      return null;
    }
  }

  async remove(sourceId: string, blockId: string, revision: number): Promise<void> {
    validateIdentity(sourceId, blockId, revision);
    this.records.delete(recordKey(sourceId, blockId, revision));
  }

  async removeBlock(sourceId: string, blockId: string): Promise<void> {
    if (!sourceId.trim()) fail('requires a source id');
    if (!blockId.trim()) fail('requires a block id');
    const prefix = blockPrefix(sourceId, blockId);
    for (const key of this.records.keys()) if (key.startsWith(prefix)) this.records.delete(key);
  }

  async removeSource(sourceId: string): Promise<void> {
    if (!sourceId.trim()) fail('requires a source id');
    const prefix = sourcePrefix(sourceId);
    for (const key of this.records.keys()) if (key.startsWith(prefix)) this.records.delete(key);
  }
}

/** IndexedDB implementation; records are mirrored in memory only when unavailable. */
export class IndexedDbSparseOverlayStore implements SparseOverlayStore {
  private readonly databaseName: string;
  private readonly factory: IndexedDbFactoryLike | null;
  private databasePromise: Promise<IDBDatabase> | null = null;

  constructor(options: SparseOverlayStoreOptions = {}) {
    this.databaseName = options.databaseName ?? DEFAULT_DATABASE_NAME;
    if (!this.databaseName.trim()) fail('database name cannot be empty');
    this.factory = resolveIndexedDbFactory(options.indexedDB);
    if (!this.factory) fail('requires IndexedDB');
  }

  private database(): Promise<IDBDatabase> {
    this.databasePromise ??= openWorkspaceDatabase({ databaseName: this.databaseName, indexedDB: this.factory })
      .then((database) => {
        if (!database) fail('requires IndexedDB');
        return database;
      });
    return this.databasePromise;
  }

  async put(sourceId: string, blockId: string, overlay: SparseCellOverlayMetadata): Promise<SparseOverlayRecord> {
    validateIdentity(sourceId, blockId, overlay?.revision);
    const record = buildRecord(sourceId, blockId, overlay);
    const database = await this.database();
    const transaction = database.transaction(OVERLAY_STORE_NAME, 'readwrite');
    transaction.objectStore(OVERLAY_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return clone(record);
  }

  async get(sourceId: string, blockId: string, revision: number): Promise<SparseOverlayRecord | null> {
    validateIdentity(sourceId, blockId, revision);
    const database = await this.database();
    const transaction = database.transaction(OVERLAY_STORE_NAME, 'readonly');
    const complete = transactionComplete(transaction);
    const value = await requestResult(transaction.objectStore(OVERLAY_STORE_NAME).get([sourceId, blockId, revision])) as SparseOverlayRecord | undefined;
    await complete;
    if (!value) return null;
    try {
      if (value.sourceId !== sourceId || value.blockId !== blockId || value.revision !== revision) {
        fail('record key does not match the requested source, block, and revision');
      }
      return validateRecord(value);
    } catch {
      await this.remove(sourceId, blockId, revision);
      return null;
    }
  }

  async remove(sourceId: string, blockId: string, revision: number): Promise<void> {
    validateIdentity(sourceId, blockId, revision);
    const database = await this.database();
    const transaction = database.transaction(OVERLAY_STORE_NAME, 'readwrite');
    transaction.objectStore(OVERLAY_STORE_NAME).delete([sourceId, blockId, revision]);
    await transactionComplete(transaction);
  }

  async removeBlock(sourceId: string, blockId: string): Promise<void> {
    if (!sourceId.trim()) fail('requires a source id');
    if (!blockId.trim()) fail('requires a block id');
    const database = await this.database();
    const transaction = database.transaction(OVERLAY_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(OVERLAY_STORE_NAME);
    const complete = transactionComplete(transaction);
    const rows = await requestResult(store.index('sourceBlock').getAll([sourceId, blockId])) as SparseOverlayRecord[];
    for (const row of rows) store.delete([row.sourceId, row.blockId, row.revision]);
    await complete;
  }

  async removeSource(sourceId: string): Promise<void> {
    if (!sourceId.trim()) fail('requires a source id');
    const database = await this.database();
    const transaction = database.transaction(OVERLAY_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(OVERLAY_STORE_NAME);
    const complete = transactionComplete(transaction);
    const rows = await requestResult(store.index('sourceId').getAll(IDBKeyRange.only(sourceId))) as SparseOverlayRecord[];
    for (const row of rows) store.delete([row.sourceId, row.blockId, row.revision]);
    await complete;
  }
}

/** Browser-first store with a durable memory namespace when IndexedDB is absent. */
export class LocalSparseOverlayStore implements SparseOverlayStore {
  private readonly indexedDb: IndexedDbSparseOverlayStore | null;
  private readonly records: Map<string, SparseOverlayRecord>;

  constructor(options: SparseOverlayStoreOptions = {}) {
    const databaseName = options.databaseName ?? DEFAULT_DATABASE_NAME;
    if (!databaseName.trim()) fail('database name cannot be empty');
    this.records = memoryRecords(databaseName);
    this.indexedDb = resolveIndexedDbFactory(options.indexedDB) === null
      ? null
      : new IndexedDbSparseOverlayStore({ ...options, databaseName });
  }

  put(sourceId: string, blockId: string, overlay: SparseCellOverlayMetadata): Promise<SparseOverlayRecord> {
    if (this.indexedDb) return this.indexedDb.put(sourceId, blockId, overlay);
    validateIdentity(sourceId, blockId, overlay?.revision);
    const record = buildRecord(sourceId, blockId, overlay);
    this.records.set(recordKey(sourceId, blockId, record.revision), clone(record));
    return Promise.resolve(clone(record));
  }

  async get(sourceId: string, blockId: string, revision: number): Promise<SparseOverlayRecord | null> {
    if (this.indexedDb) return this.indexedDb.get(sourceId, blockId, revision);
    validateIdentity(sourceId, blockId, revision);
    const value = this.records.get(recordKey(sourceId, blockId, revision));
    if (!value) return null;
    try {
      return validateRecord(clone(value));
    } catch {
      this.records.delete(recordKey(sourceId, blockId, revision));
      return null;
    }
  }

  remove(sourceId: string, blockId: string, revision: number): Promise<void> {
    if (this.indexedDb) return this.indexedDb.remove(sourceId, blockId, revision);
    validateIdentity(sourceId, blockId, revision);
    this.records.delete(recordKey(sourceId, blockId, revision));
    return Promise.resolve();
  }

  removeBlock(sourceId: string, blockId: string): Promise<void> {
    if (this.indexedDb) return this.indexedDb.removeBlock(sourceId, blockId);
    if (!sourceId.trim()) fail('requires a source id');
    if (!blockId.trim()) fail('requires a block id');
    const prefix = blockPrefix(sourceId, blockId);
    for (const key of this.records.keys()) if (key.startsWith(prefix)) this.records.delete(key);
    return Promise.resolve();
  }

  removeSource(sourceId: string): Promise<void> {
    if (this.indexedDb) return this.indexedDb.removeSource(sourceId);
    if (!sourceId.trim()) fail('requires a source id');
    const prefix = sourcePrefix(sourceId);
    for (const key of this.records.keys()) if (key.startsWith(prefix)) this.records.delete(key);
    return Promise.resolve();
  }
}
