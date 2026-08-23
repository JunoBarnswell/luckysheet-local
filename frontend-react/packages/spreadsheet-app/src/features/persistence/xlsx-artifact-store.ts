import type { XlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import { verifyXlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import { XLSX_ARTIFACT_STORE } from './data-block-store';

const DATABASE_NAME = 'react-sheets-workspaces';
const DATABASE_SCHEMA_REVISION = 3;

export interface XlsxArtifactRecord {
  schema: 'XlsxArtifactRecord';
  version: 1;
  unitId: string;
  artifact: XlsxSourceArtifact;
  updatedAt: string;
}

const memoryArtifacts = new Map<string, XlsxArtifactRecord>();

function copyArtifact(artifact: XlsxSourceArtifact): XlsxSourceArtifact {
  return { ...artifact, buffer: artifact.buffer.slice(0), detectedFeatures: [...artifact.detectedFeatures] };
}

function copyRecord(record: XlsxArtifactRecord): XlsxArtifactRecord {
  return { ...record, artifact: copyArtifact(record.artifact) };
}

function factory(): IDBFactory | null {
  return typeof indexedDB === 'undefined' ? null : indexedDB;
}

function request<T>(value: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    value.onsuccess = () => resolve(value.result);
    value.onerror = () => reject(value.error ?? new Error('IndexedDB request failed'));
  });
}

function complete(transaction: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB transaction failed'));
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted'));
  });
}

async function database(): Promise<IDBDatabase | null> {
  const value = factory();
  if (!value) return null;
  const open = value.open(DATABASE_NAME, DATABASE_SCHEMA_REVISION);
  open.onupgradeneeded = () => {
    if (!open.result.objectStoreNames.contains(XLSX_ARTIFACT_STORE)) {
      open.result.createObjectStore(XLSX_ARTIFACT_STORE, { keyPath: 'unitId' });
    }
    if (!open.result.objectStoreNames.contains('dataBlocks')) {
      const store = open.result.createObjectStore('dataBlocks', { keyPath: ['sourceId', 'blockId'] });
      store.createIndex('sourceId', 'sourceId', { unique: false });
    }
  };
  return request(open);
}

/** Per-workbook persisted original XLSX bytes; never part of a Snapshot. */
export class LocalXlsxArtifactStore {
  private databasePromise: Promise<IDBDatabase | null> | null = null;

  private getDatabase(): Promise<IDBDatabase | null> {
    this.databasePromise ??= database();
    return this.databasePromise;
  }

  async save(unitId: string, artifact: XlsxSourceArtifact): Promise<XlsxArtifactRecord> {
    if (!unitId.trim()) throw new Error('Workbook unitId is required for an XLSX artifact');
    await verifyXlsxSourceArtifact(artifact);
    const record: XlsxArtifactRecord = {
      schema: 'XlsxArtifactRecord',
      version: 1,
      unitId,
      artifact: copyArtifact(artifact),
      updatedAt: new Date().toISOString(),
    };
    const db = await this.getDatabase();
    if (!db) {
      memoryArtifacts.set(unitId, copyRecord(record));
      return copyRecord(record);
    }
    const transaction = db.transaction(XLSX_ARTIFACT_STORE, 'readwrite');
    transaction.objectStore(XLSX_ARTIFACT_STORE).put(record);
    await complete(transaction);
    return copyRecord(record);
  }

  async load(unitId: string): Promise<XlsxSourceArtifact | null> {
    const db = await this.getDatabase();
    const record = db
      ? await (async () => {
        const transaction = db.transaction(XLSX_ARTIFACT_STORE, 'readonly');
        const value = await request(transaction.objectStore(XLSX_ARTIFACT_STORE).get(unitId)) as XlsxArtifactRecord | undefined;
        await complete(transaction);
        return value;
      })()
      : memoryArtifacts.get(unitId);
    if (!record || record.schema !== 'XlsxArtifactRecord' || record.version !== 1 || record.unitId !== unitId) return null;
    try {
      await verifyXlsxSourceArtifact(record.artifact);
      return copyArtifact(record.artifact);
    } catch {
      await this.remove(unitId);
      return null;
    }
  }

  async remove(unitId: string): Promise<void> {
    const db = await this.getDatabase();
    if (!db) {
      memoryArtifacts.delete(unitId);
      return;
    }
    const transaction = db.transaction(XLSX_ARTIFACT_STORE, 'readwrite');
    transaction.objectStore(XLSX_ARTIFACT_STORE).delete(unitId);
    await complete(transaction);
  }
}
