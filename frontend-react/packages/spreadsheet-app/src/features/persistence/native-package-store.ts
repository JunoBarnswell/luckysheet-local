import type { NativePackageState } from '@react-sheets/exchange-excel-ooxml';
import { loadOpcPackageGraph, verifyNativePackageState } from '@react-sheets/exchange-excel-ooxml';
import {
  openWorkspaceDatabase,
  requestResult,
  transactionComplete,
  type IndexedDbStoreOptions,
  WORKSPACE_DATABASE_NAME,
  NATIVE_PACKAGE_STORE_NAME,
} from './indexed-db';

export interface NativePackageRecord {
  schema: 'NativePackageRecord';
  version: 1;
  unitId: string;
  artifact: NativePackageState;
  updatedAt: string;
}

const memoryDatabases = new Map<string, Map<string, NativePackageRecord>>();

function memoryArtifacts(databaseName: string): Map<string, NativePackageRecord> {
  let records = memoryDatabases.get(databaseName);
  if (!records) {
    records = new Map<string, NativePackageRecord>();
    memoryDatabases.set(databaseName, records);
  }
  return records;
}

function copyArtifact(artifact: NativePackageState): NativePackageState {
  return {
    ...artifact,
    sourceBytes: artifact.sourceBytes.slice(0),
    packageGraph: structuredClone(artifact.packageGraph),
    detectedFeatures: [...artifact.detectedFeatures],
    ownership: structuredClone(artifact.ownership),
    compatibility: structuredClone(artifact.compatibility),
  };
}

function copyRecord(record: NativePackageRecord): NativePackageRecord {
  return { ...record, artifact: copyArtifact(record.artifact) };
}

function compactArtifact(artifact: NativePackageState): NativePackageState {
  return {
    ...artifact,
    sourceBytes: artifact.sourceBytes.slice(0),
    packageGraph: {
      ...structuredClone(artifact.packageGraph),
      parts: {},
      opaqueParts: {},
      contentTypesXml: artifact.packageGraph.contentTypesXml?.slice(),
    },
    detectedFeatures: [...artifact.detectedFeatures],
    ownership: structuredClone(artifact.ownership),
    compatibility: structuredClone(artifact.compatibility),
  };
}

export async function buildNativePackageRecord(unitId: string, artifact: NativePackageState): Promise<NativePackageRecord> {
  if (!unitId.trim()) throw new Error('Workbook unitId is required for a native package');
  await verifyNativePackageState(artifact);
  return {
    schema: 'NativePackageRecord',
    version: 1,
    unitId,
    artifact: compactArtifact(artifact),
    updatedAt: new Date().toISOString(),
  };
}

/** Per-workbook persisted native package bytes; never part of a Snapshot. */
export class LocalNativePackageStore {
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

  async save(unitId: string, artifact: NativePackageState): Promise<NativePackageRecord> {
    const record = await buildNativePackageRecord(unitId, artifact);
    const db = await this.getDatabase();
    if (!db) {
      memoryArtifacts(this.databaseName).set(unitId, copyRecord(record));
      return copyRecord(record);
    }
    const transaction = db.transaction(NATIVE_PACKAGE_STORE_NAME, 'readwrite');
    transaction.objectStore(NATIVE_PACKAGE_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return copyRecord(record);
  }

  async load(unitId: string): Promise<NativePackageState | null> {
    const db = await this.getDatabase();
    const record = db
      ? await (async () => {
        const transaction = db.transaction(NATIVE_PACKAGE_STORE_NAME, 'readonly');
        const value = await requestResult(transaction.objectStore(NATIVE_PACKAGE_STORE_NAME).get(unitId)) as NativePackageRecord | undefined;
        await transactionComplete(transaction);
        return value;
      })()
      : memoryArtifacts(this.databaseName).get(unitId);
    if (!record || record.schema !== 'NativePackageRecord' || record.version !== 1 || record.unitId !== unitId) return null;
    try {
      await verifyNativePackageState(record.artifact);
      let packageGraph = record.artifact.packageGraph;
      if (Object.keys(packageGraph.parts).length === 0) {
        try {
          packageGraph = loadOpcPackageGraph(record.artifact.sourceBytes, {}, record.artifact.fileName).packageGraph;
        } catch {
          // Focused persistence tests and explicit memory callers may provide
          // a verified native state whose source bytes are not a full archive.
          // Keep its canonical metadata instead of deleting a checksummed record.
        }
      }
      return copyArtifact({ ...record.artifact, packageGraph });
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
    const transaction = db.transaction(NATIVE_PACKAGE_STORE_NAME, 'readwrite');
    transaction.objectStore(NATIVE_PACKAGE_STORE_NAME).delete(unitId);
    await transactionComplete(transaction);
  }
}
