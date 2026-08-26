import type { NativePackageState } from '@react-sheets/exchange-excel-ooxml';
import { loadOpcPackageGraph, verifyNativePackageState } from '@react-sheets/exchange-excel-ooxml';
import {
  resolveWorkspaceDatabaseCoordinator,
  requestResult,
  transactionComplete,
  type IndexedDbStoreOptions,
  type WorkspaceDatabaseCoordinator,
  NATIVE_PACKAGE_STORE_NAME,
} from './indexed-db';

export interface NativePackageRecord {
  schema: 'NativePackageRecord';
  version: 1;
  unitId: string;
  artifact: NativePackageState;
  updatedAt: string;
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
  private readonly coordinator: WorkspaceDatabaseCoordinator;

  constructor(options: IndexedDbStoreOptions = {}) {
    this.coordinator = resolveWorkspaceDatabaseCoordinator(options);
  }

  async save(unitId: string, artifact: NativePackageState): Promise<NativePackageRecord> {
    const record = await buildNativePackageRecord(unitId, artifact);
    const transaction = await this.coordinator.transaction(NATIVE_PACKAGE_STORE_NAME, 'readwrite');
    transaction.objectStore(NATIVE_PACKAGE_STORE_NAME).put(record);
    await transactionComplete(transaction);
    return copyRecord(record);
  }

  async load(unitId: string): Promise<NativePackageState | null> {
    const transaction = await this.coordinator.transaction(NATIVE_PACKAGE_STORE_NAME, 'readonly');
    const record = await requestResult(transaction.objectStore(NATIVE_PACKAGE_STORE_NAME).get(unitId)) as NativePackageRecord | undefined;
    await transactionComplete(transaction);
    if (!record) return null;
    if (record.schema !== 'NativePackageRecord' || record.version !== 1 || record.unitId !== unitId) {
      throw new Error(`NATIVE_PACKAGE_SCHEMA_INVALID: ${unitId}`);
    }
    await verifyNativePackageState(record.artifact);
    const packageGraph = Object.keys(record.artifact.packageGraph.parts).length === 0
      ? loadOpcPackageGraph(record.artifact.sourceBytes, {}, record.artifact.fileName).packageGraph
      : record.artifact.packageGraph;
    return copyArtifact({ ...record.artifact, packageGraph });
  }

  async remove(unitId: string): Promise<void> {
    const transaction = await this.coordinator.transaction(NATIVE_PACKAGE_STORE_NAME, 'readwrite');
    transaction.objectStore(NATIVE_PACKAGE_STORE_NAME).delete(unitId);
    await transactionComplete(transaction);
  }
}
