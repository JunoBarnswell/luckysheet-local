import { computeChecksum } from '../persistence/checksum';
import {
  namespaceWorkspaceSourceId,
  requestResult,
  resolveWorkspaceDatabaseCoordinator,
  resolveWorkspaceUnitId,
  transactionComplete,
  OVERLAY_STORE_NAME,
  WORKSPACE_DATABASE_NAME,
  type IndexedDbFactoryLike,
  type WorkspaceDatabaseCoordinator,
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
  coordinator?: WorkspaceDatabaseCoordinator;
  unitId?: string | (() => string);
}

function storageSourceId(options: Pick<SparseOverlayStoreOptions, 'unitId'>, sourceId: string): string {
  return namespaceWorkspaceSourceId(resolveWorkspaceUnitId(options), sourceId);
}

function fail(message: string): never {
  throw new Error(`Sparse overlay store ${message}`);
}

function isSafeRevision(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
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

/** IndexedDB implementation owned by the shared workspace coordinator. */
export class IndexedDbSparseOverlayStore implements SparseOverlayStore {
  private readonly coordinator: WorkspaceDatabaseCoordinator;
  private readonly options: SparseOverlayStoreOptions;

  constructor(options: SparseOverlayStoreOptions = {}) {
    this.options = options;
    const databaseName = options.databaseName ?? DEFAULT_DATABASE_NAME;
    if (!databaseName.trim()) fail('database name cannot be empty');
    this.coordinator = resolveWorkspaceDatabaseCoordinator(options);
  }

  async put(sourceId: string, blockId: string, overlay: SparseCellOverlayMetadata): Promise<SparseOverlayRecord> {
    validateIdentity(sourceId, blockId, overlay?.revision);
    const storageId = storageSourceId(this.options, sourceId);
    const record = buildRecord(storageId, blockId, overlay);
    const transaction = await this.coordinator.transaction(OVERLAY_STORE_NAME, 'readwrite');
    transaction.objectStore(OVERLAY_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return clone(record);
  }

  async get(sourceId: string, blockId: string, revision: number): Promise<SparseOverlayRecord | null> {
    validateIdentity(sourceId, blockId, revision);
    const transaction = await this.coordinator.transaction(OVERLAY_STORE_NAME, 'readonly');
    const complete = transactionComplete(transaction);
    const store = transaction.objectStore(OVERLAY_STORE_NAME);
    const storageId = storageSourceId(this.options, sourceId);
    const value = await requestResult(store.get([storageId, blockId, revision])) as SparseOverlayRecord | undefined;
    await complete;
    if (!value) return null;
    if (value.sourceId !== storageId
      || value.blockId !== blockId
      || value.revision !== revision) {
      fail('record key does not match the requested source, block, and revision');
    }
    return validateRecord(value);
  }

  async remove(sourceId: string, blockId: string, revision: number): Promise<void> {
    validateIdentity(sourceId, blockId, revision);
    const transaction = await this.coordinator.transaction(OVERLAY_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(OVERLAY_STORE_NAME);
    const storageId = storageSourceId(this.options, sourceId);
    store.delete([storageId, blockId, revision]);
    await transactionComplete(transaction);
  }

  async removeBlock(sourceId: string, blockId: string): Promise<void> {
    if (!sourceId.trim()) fail('requires a source id');
    if (!blockId.trim()) fail('requires a block id');
    const transaction = await this.coordinator.transaction(OVERLAY_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(OVERLAY_STORE_NAME);
    const storageId = storageSourceId(this.options, sourceId);
    const complete = transactionComplete(transaction);
    const rows = await requestResult(store.index('sourceBlock').getAll([storageId, blockId])) as SparseOverlayRecord[];
    for (const row of rows) store.delete([row.sourceId, row.blockId, row.revision]);
    await complete;
  }

  async removeSource(sourceId: string): Promise<void> {
    if (!sourceId.trim()) fail('requires a source id');
    const transaction = await this.coordinator.transaction(OVERLAY_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(OVERLAY_STORE_NAME);
    const storageId = storageSourceId(this.options, sourceId);
    const complete = transactionComplete(transaction);
    const rows = await requestResult(store.index('sourceId').getAll(IDBKeyRange.only(storageId))) as SparseOverlayRecord[];
    for (const row of rows) store.delete([row.sourceId, row.blockId, row.revision]);
    await complete;
  }
}

/** Canonical local overlay store. */
export class LocalSparseOverlayStore implements SparseOverlayStore {
  private readonly indexedDb: IndexedDbSparseOverlayStore;

  constructor(options: SparseOverlayStoreOptions = {}) {
    const databaseName = options.databaseName ?? DEFAULT_DATABASE_NAME;
    if (!databaseName.trim()) fail('database name cannot be empty');
    this.indexedDb = new IndexedDbSparseOverlayStore({ ...options, databaseName });
  }

  put(sourceId: string, blockId: string, overlay: SparseCellOverlayMetadata): Promise<SparseOverlayRecord> {
    return this.indexedDb.put(sourceId, blockId, overlay);
  }

  get(sourceId: string, blockId: string, revision: number): Promise<SparseOverlayRecord | null> {
    return this.indexedDb.get(sourceId, blockId, revision);
  }

  remove(sourceId: string, blockId: string, revision: number): Promise<void> {
    return this.indexedDb.remove(sourceId, blockId, revision);
  }

  removeBlock(sourceId: string, blockId: string): Promise<void> {
    return this.indexedDb.removeBlock(sourceId, blockId);
  }

  removeSource(sourceId: string): Promise<void> {
    return this.indexedDb.removeSource(sourceId);
  }
}
