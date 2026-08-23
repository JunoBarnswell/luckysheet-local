import type { XlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import { verifyXlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import {
  openWorkspaceDatabase,
  requestResult,
  transactionComplete,
  type IndexedDbStoreOptions,
  WORKSPACE_DATABASE_NAME,
  XLSX_ARTIFACT_STORE_NAME,
} from './indexed-db';

export interface XlsxArtifactRecord {
  schema: 'XlsxArtifactRecord';
  version: 1;
  unitId: string;
  artifact: XlsxSourceArtifact;
  updatedAt: string;
}

const memoryDatabases = new Map<string, Map<string, XlsxArtifactRecord>>();

function memoryArtifacts(databaseName: string): Map<string, XlsxArtifactRecord> {
  let records = memoryDatabases.get(databaseName);
  if (!records) {
    records = new Map<string, XlsxArtifactRecord>();
    memoryDatabases.set(databaseName, records);
  }
  return records;
}

function copyArtifact(artifact: XlsxSourceArtifact): XlsxSourceArtifact {
  return { ...artifact, buffer: artifact.buffer.slice(0), detectedFeatures: [...artifact.detectedFeatures], ...(artifact.capabilityReport ? { capabilityReport: structuredClone(artifact.capabilityReport) } : {}) };
}

function copyRecord(record: XlsxArtifactRecord): XlsxArtifactRecord {
  return { ...record, artifact: copyArtifact(record.artifact) };
}

export async function buildXlsxArtifactRecord(unitId: string, artifact: XlsxSourceArtifact): Promise<XlsxArtifactRecord> {
  if (!unitId.trim()) throw new Error('Workbook unitId is required for an XLSX artifact');
  await verifyXlsxSourceArtifact(artifact);
  return {
    schema: 'XlsxArtifactRecord',
    version: 1,
    unitId,
    artifact: copyArtifact(artifact),
    updatedAt: new Date().toISOString(),
  };
}

/** Per-workbook persisted original XLSX bytes; never part of a Snapshot. */
export class LocalXlsxArtifactStore {
  private readonly options: IndexedDbStoreOptions;
  private readonly databaseName: string;
  private databasePromise: Promise<IDBDatabase | null> | null = null;

  constructor(options: IndexedDbStoreOptions = {}) {
    this.options = options;
    this.databaseName = options.databaseName ?? WORKSPACE_DATABASE_NAME;
  }

  private getDatabase(): Promise<IDBDatabase | null> {
    this.databasePromise ??= openWorkspaceDatabase(this.options);
    return this.databasePromise;
  }

  async save(unitId: string, artifact: XlsxSourceArtifact): Promise<XlsxArtifactRecord> {
    const record = await buildXlsxArtifactRecord(unitId, artifact);
    const db = await this.getDatabase();
    if (!db) {
      memoryArtifacts(this.databaseName).set(unitId, copyRecord(record));
      return copyRecord(record);
    }
    const transaction = db.transaction(XLSX_ARTIFACT_STORE_NAME, 'readwrite');
    transaction.objectStore(XLSX_ARTIFACT_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return copyRecord(record);
  }

  async load(unitId: string): Promise<XlsxSourceArtifact | null> {
    const db = await this.getDatabase();
    const record = db
      ? await (async () => {
        const transaction = db.transaction(XLSX_ARTIFACT_STORE_NAME, 'readonly');
        const value = await requestResult(transaction.objectStore(XLSX_ARTIFACT_STORE_NAME).get(unitId)) as XlsxArtifactRecord | undefined;
        await transactionComplete(transaction);
        return value;
      })()
      : memoryArtifacts(this.databaseName).get(unitId);
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
      memoryArtifacts(this.databaseName).delete(unitId);
      return;
    }
    const transaction = db.transaction(XLSX_ARTIFACT_STORE_NAME, 'readwrite');
    transaction.objectStore(XLSX_ARTIFACT_STORE_NAME).delete(unitId);
    await transactionComplete(transaction);
  }
}
