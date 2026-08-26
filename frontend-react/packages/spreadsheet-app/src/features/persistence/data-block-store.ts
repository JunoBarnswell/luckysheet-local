import type { DataBlockRef } from '@react-sheets/core-model';
import { computeBinaryChecksum } from './checksum';
import {
  DATA_BLOCK_STORE_NAME,
  namespaceWorkspaceSourceId,
  requestResult,
  resolveWorkspaceDatabaseCoordinator,
  resolveWorkspaceUnitId,
  transactionComplete,
  type IndexedDbStoreOptions,
  type WorkspaceDatabaseCoordinator,
} from './indexed-db';

export interface DataBlockRecord {
  schema: 'DataBlockRecord';
  sourceId: string;
  blockId: string;
  checksum: string;
  bytes: ArrayBuffer;
  updatedAt: string;
}

function cloneBytes(bytes: ArrayBuffer): ArrayBuffer {
  return bytes.slice(0);
}

function cloneRecord(record: DataBlockRecord): DataBlockRecord {
  return { ...record, bytes: cloneBytes(record.bytes) };
}

async function assertRecord(record: DataBlockRecord): Promise<void> {
  if (record.schema !== 'DataBlockRecord' || !record.sourceId || !record.blockId || !record.checksum) {
    throw new Error('Invalid data block record');
  }
  if (!(record.bytes instanceof ArrayBuffer)) throw new Error('Data block bytes must be an ArrayBuffer');
  if (await computeBinaryChecksum(record.bytes) !== record.checksum) {
    throw new Error(`Data block checksum mismatch: ${record.blockId}`);
  }
}

/**
 * Persistent bytes for block-backed sources. Snapshot/presence/history retain
 * only DataBlockRef metadata so a remote operation never serializes source
 * data into a JSON changeset.
 */
export class LocalDataBlockStore {
  private readonly coordinator: WorkspaceDatabaseCoordinator;
  private readonly unitId: string | (() => string) | undefined;

  constructor(options: IndexedDbStoreOptions = {}) {
    this.coordinator = resolveWorkspaceDatabaseCoordinator(options);
    this.unitId = options.unitId;
  }

  private storageSourceId(sourceId: string): string {
    return namespaceWorkspaceSourceId(resolveWorkspaceUnitId({ unitId: this.unitId }), sourceId);
  }

  async put(ref: DataBlockRef, bytes: ArrayBuffer): Promise<DataBlockRecord> {
    const checksum = await computeBinaryChecksum(bytes);
    if (checksum !== ref.checksum) throw new Error(`Data block checksum does not match manifest: ${ref.id}`);
    const record: DataBlockRecord = {
      schema: 'DataBlockRecord',
      sourceId: this.storageSourceId(ref.dataSourceId),
      blockId: ref.id,
      checksum,
      bytes: cloneBytes(bytes),
      updatedAt: new Date().toISOString(),
    };
    await assertRecord(record);
    const transaction = await this.coordinator.transaction(DATA_BLOCK_STORE_NAME, 'readwrite');
    transaction.objectStore(DATA_BLOCK_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return cloneRecord(record);
  }

  async get(ref: Pick<DataBlockRef, 'dataSourceId' | 'id' | 'checksum'>): Promise<DataBlockRecord | null> {
    const transaction = await this.coordinator.transaction(DATA_BLOCK_STORE_NAME, 'readonly');
    const store = transaction.objectStore(DATA_BLOCK_STORE_NAME);
    const storageSourceId = this.storageSourceId(ref.dataSourceId);
    const record = await requestResult(store.get([storageSourceId, ref.id])) as DataBlockRecord | undefined;
    await transactionComplete(transaction);
    if (!record) return null;
    await assertRecord(record);
    if (record.checksum !== ref.checksum) throw new Error(`Data block manifest checksum mismatch: ${ref.id}`);
    return cloneRecord(record);
  }

  async remove(sourceId: string, blockId: string): Promise<void> {
    const storageSourceId = this.storageSourceId(sourceId);
    const transaction = await this.coordinator.transaction(DATA_BLOCK_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(DATA_BLOCK_STORE_NAME);
    store.delete([storageSourceId, blockId]);
    await transactionComplete(transaction);
  }

  async removeSource(sourceId: string): Promise<void> {
    const storageSourceId = this.storageSourceId(sourceId);
    const transaction = await this.coordinator.transaction(DATA_BLOCK_STORE_NAME, 'readwrite');
    const index = transaction.objectStore(DATA_BLOCK_STORE_NAME).index('sourceId');
    const rows = await requestResult(index.getAll(IDBKeyRange.only(storageSourceId))) as DataBlockRecord[];
    for (const row of rows) transaction.objectStore(DATA_BLOCK_STORE_NAME).delete([row.sourceId, row.blockId]);
    await transactionComplete(transaction);
  }
}
