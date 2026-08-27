import type { NativePackageState } from '@react-sheets/exchange-excel-ooxml';
import { loadOpcPackageGraph, verifyNativePackageState } from '@react-sheets/exchange-excel-ooxml';
import { memoryKey, type WorkspaceMemoryCoordinator } from './memory';

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

/** Per-workbook session-memory native package bytes. */
export class LocalNativePackageStore {
  constructor(private readonly coordinator: WorkspaceMemoryCoordinator) {}

  async save(unitId: string, artifact: NativePackageState): Promise<NativePackageRecord> {
    const record = await buildNativePackageRecord(unitId, artifact);
    return this.coordinator.transaction((transaction) => {
      transaction.set('nativePackages', unitId, copyRecord(record));
      return copyRecord(record);
    });
  }

  async load(unitId: string): Promise<NativePackageState | null> {
    return this.coordinator.read((transaction) => {
      const record = transaction.get<NativePackageRecord>('nativePackages', unitId);
      if (!record) return null;
      if (record.schema !== 'NativePackageRecord' || record.version !== 1 || record.unitId !== unitId) {
        throw new Error(`NATIVE_PACKAGE_SCHEMA_INVALID: ${unitId}`);
      }
      return verifyNativePackageState(record.artifact).then(() => {
        const packageGraph = Object.keys(record.artifact.packageGraph.parts).length === 0
          ? loadOpcPackageGraph(record.artifact.sourceBytes, {}, record.artifact.fileName).packageGraph
          : record.artifact.packageGraph;
        return copyArtifact({ ...record.artifact, packageGraph });
      });
    });
  }

  async remove(unitId: string): Promise<void> {
    await this.coordinator.transaction((transaction) => transaction.delete('nativePackages', memoryKey(unitId)));
  }
}
