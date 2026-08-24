import type { DataBlockRef } from '@react-sheets/core-model';
import { computeBinaryChecksum } from './checksum';
import {
  DATA_BLOCK_STORE_NAME,
  namespaceWorkspaceSourceId,
  openWorkspaceDatabase,
  requestResult,
  resolveWorkspaceUnitId,
  transactionComplete,
  type IndexedDbStoreOptions,
  WORKSPACE_DATABASE_NAME,
} from './indexed-db';

export interface DataBlockRecord {
  schema: 'DataBlockRecord';
  sourceId: string;
  blockId: string;
  checksum: string;
  bytes: ArrayBuffer;
  updatedAt: string;
}

const memoryDatabases = new Map<string, Map<string, DataBlockRecord>>();

function memoryRecords(databaseName: string): Map<string, DataBlockRecord> {
  let records = memoryDatabases.get(databaseName);
  if (!records) {
    records = new Map<string, DataBlockRecord>();
    memoryDatabases.set(databaseName, records);
  }
  return records;
}

function recordKey(sourceId: string, blockId: string): string {
  return `${sourceId}:${blockId}`;
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
  private readonly options: IndexedDbStoreOptions;
  private readonly databaseName: string;
  private readonly unitId: string | (() => string) | undefined;
  private database: Promise<IDBDatabase | null> | null = null;

  constructor(options: IndexedDbStoreOptions = {}) {
    this.options = options;
    this.databaseName = options.databaseName ?? WORKSPACE_DATABASE_NAME;
    this.unitId = options.unitId;
  }

  private storageSourceId(sourceId: string): string {
    return namespaceWorkspaceSourceId(resolveWorkspaceUnitId({ unitId: this.unitId }), sourceId);
  }

  private async getDatabase(): Promise<IDBDatabase | null> {
    this.database ??= openWorkspaceDatabase(this.options);
    return this.database;
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
    const database = await this.getDatabase();
    if (!database) {
      memoryRecords(this.databaseName).set(recordKey(record.sourceId, record.blockId), cloneRecord(record));
      return cloneRecord(record);
    }
    const transaction = database.transaction(DATA_BLOCK_STORE_NAME, 'readwrite');
    transaction.objectStore(DATA_BLOCK_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return cloneRecord(record);
  }

  async get(ref: Pick<DataBlockRef, 'dataSourceId' | 'id' | 'checksum'>): Promise<DataBlockRecord | null> {
    const database = await this.getDatabase();
    let record: DataBlockRecord | undefined;
    if (!database) {
      const storageSourceId = this.storageSourceId(ref.dataSourceId);
      record = memoryRecords(this.databaseName).get(recordKey(storageSourceId, ref.id));
      if (!record && storageSourceId !== ref.dataSourceId) {
        // Read the pre-namespace record once so existing workbooks are not
        // stranded after the key-space migration.
        record = memoryRecords(this.databaseName).get(recordKey(ref.dataSourceId, ref.id));
      }
    } else {
      const transaction = database.transaction(DATA_BLOCK_STORE_NAME, 'readonly');
      const store = transaction.objectStore(DATA_BLOCK_STORE_NAME);
      const storageSourceId = this.storageSourceId(ref.dataSourceId);
      record = await requestResult(store.get([storageSourceId, ref.id])) as DataBlockRecord | undefined;
      if (!record && storageSourceId !== ref.dataSourceId) {
        record = await requestResult(store.get([ref.dataSourceId, ref.id])) as DataBlockRecord | undefined;
      }
      await transactionComplete(transaction);
    }
    if (!record) return null;
    try {
      await assertRecord(record);
    } catch {
      await this.remove(ref.dataSourceId, ref.id);
      return null;
    }
    if (record.checksum !== ref.checksum) return null;
    return cloneRecord(record);
  }

  async remove(sourceId: string, blockId: string): Promise<void> {
    const database = await this.getDatabase();
    const storageSourceId = this.storageSourceId(sourceId);
    if (!database) {
      const records = memoryRecords(this.databaseName);
      records.delete(recordKey(storageSourceId, blockId));
      if (storageSourceId !== sourceId) records.delete(recordKey(sourceId, blockId));
      return;
    }
    const transaction = database.transaction(DATA_BLOCK_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(DATA_BLOCK_STORE_NAME);
    store.delete([storageSourceId, blockId]);
    if (storageSourceId !== sourceId) store.delete([sourceId, blockId]);
    await transactionComplete(transaction);
  }

  async removeSource(sourceId: string): Promise<void> {
    const database = await this.getDatabase();
    const storageSourceId = this.storageSourceId(sourceId);
    if (!database) {
      const records = memoryRecords(this.databaseName);
      for (const key of records.keys()) {
        if (key.startsWith(`${storageSourceId}:`) || (storageSourceId !== sourceId && key.startsWith(`${sourceId}:`))) records.delete(key);
      }
      return;
    }
    const transaction = database.transaction(DATA_BLOCK_STORE_NAME, 'readwrite');
    const index = transaction.objectStore(DATA_BLOCK_STORE_NAME).index('sourceId');
    const rows = [
      ...await requestResult(index.getAll(IDBKeyRange.only(storageSourceId))) as DataBlockRecord[],
      ...(storageSourceId === sourceId ? [] : await requestResult(index.getAll(IDBKeyRange.only(sourceId))) as DataBlockRecord[]),
    ];
    for (const row of rows) transaction.objectStore(DATA_BLOCK_STORE_NAME).delete([row.sourceId, row.blockId]);
    await transactionComplete(transaction);
  }
}
