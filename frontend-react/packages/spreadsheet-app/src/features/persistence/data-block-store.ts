import type { DataBlockRef } from '@react-sheets/core-model';
import { computeBinaryChecksum } from './checksum';

const DATABASE_NAME = 'react-sheets-workspaces';
const DATABASE_SCHEMA_REVISION = 3;
const DATA_BLOCK_STORE = 'dataBlocks';
export const XLSX_ARTIFACT_STORE = 'xlsxArtifacts';

export interface DataBlockRecord {
  schema: 'DataBlockRecord';
  sourceId: string;
  blockId: string;
  checksum: string;
  bytes: ArrayBuffer;
  updatedAt: string;
}

const memoryRecords = new Map<string, DataBlockRecord>();

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

function resolveFactory(): IDBFactory | null {
  return typeof indexedDB === 'undefined' ? null : indexedDB;
}

function request<T>(value: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    value.onsuccess = () => resolve(value.result);
    value.onerror = () => reject(value.error ?? new Error('IndexedDB request failed'));
  });
}

function transactionComplete(transaction: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB transaction failed'));
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted'));
  });
}

async function openDatabase(factory: IDBFactory): Promise<IDBDatabase> {
  const open = factory.open(DATABASE_NAME, DATABASE_SCHEMA_REVISION);
  open.onupgradeneeded = () => {
    const database = open.result;
    if (!database.objectStoreNames.contains(DATA_BLOCK_STORE)) {
      const store = database.createObjectStore(DATA_BLOCK_STORE, { keyPath: ['sourceId', 'blockId'] });
      store.createIndex('sourceId', 'sourceId', { unique: false });
    }
    if (!database.objectStoreNames.contains(XLSX_ARTIFACT_STORE)) {
      database.createObjectStore(XLSX_ARTIFACT_STORE, { keyPath: 'unitId' });
    }
  };
  return request(open);
}

/**
 * Persistent bytes for block-backed sources. Snapshot/presence/history retain
 * only DataBlockRef metadata so a remote operation never serializes source
 * data into a JSON changeset.
 */
export class LocalDataBlockStore {
  private database: Promise<IDBDatabase> | null = null;

  private get factory(): IDBFactory | null {
    return resolveFactory();
  }

  private async getDatabase(): Promise<IDBDatabase | null> {
    const factory = this.factory;
    if (!factory) return null;
    this.database ??= openDatabase(factory);
    return this.database;
  }

  async put(ref: DataBlockRef, bytes: ArrayBuffer): Promise<DataBlockRecord> {
    const checksum = await computeBinaryChecksum(bytes);
    if (checksum !== ref.checksum) throw new Error(`Data block checksum does not match manifest: ${ref.id}`);
    const record: DataBlockRecord = {
      schema: 'DataBlockRecord',
      sourceId: ref.dataSourceId,
      blockId: ref.id,
      checksum,
      bytes: cloneBytes(bytes),
      updatedAt: new Date().toISOString(),
    };
    await assertRecord(record);
    const database = await this.getDatabase();
    if (!database) {
      memoryRecords.set(recordKey(record.sourceId, record.blockId), cloneRecord(record));
      return cloneRecord(record);
    }
    const transaction = database.transaction(DATA_BLOCK_STORE, 'readwrite');
    transaction.objectStore(DATA_BLOCK_STORE).put(record);
    await transactionComplete(transaction);
    return cloneRecord(record);
  }

  async get(ref: Pick<DataBlockRef, 'dataSourceId' | 'id' | 'checksum'>): Promise<DataBlockRecord | null> {
    const database = await this.getDatabase();
    let record: DataBlockRecord | undefined;
    if (!database) {
      record = memoryRecords.get(recordKey(ref.dataSourceId, ref.id));
    } else {
      const transaction = database.transaction(DATA_BLOCK_STORE, 'readonly');
      record = await request(transaction.objectStore(DATA_BLOCK_STORE).get([ref.dataSourceId, ref.id])) as DataBlockRecord | undefined;
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
    if (!database) {
      memoryRecords.delete(recordKey(sourceId, blockId));
      return;
    }
    const transaction = database.transaction(DATA_BLOCK_STORE, 'readwrite');
    transaction.objectStore(DATA_BLOCK_STORE).delete([sourceId, blockId]);
    await transactionComplete(transaction);
  }

  async removeSource(sourceId: string): Promise<void> {
    const database = await this.getDatabase();
    if (!database) {
      for (const key of memoryRecords.keys()) {
        if (key.startsWith(`${sourceId}:`)) memoryRecords.delete(key);
      }
      return;
    }
    const transaction = database.transaction(DATA_BLOCK_STORE, 'readwrite');
    const index = transaction.objectStore(DATA_BLOCK_STORE).index('sourceId');
    const rows = await request(index.getAll(IDBKeyRange.only(sourceId))) as DataBlockRecord[];
    for (const row of rows) transaction.objectStore(DATA_BLOCK_STORE).delete([row.sourceId, row.blockId]);
    await transactionComplete(transaction);
  }
}
