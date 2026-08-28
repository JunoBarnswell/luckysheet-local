import type { NativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
import { loadOpcPackageGraph, verifyNativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
import { memoryKey, type WorkspaceMemoryCoordinator } from './memory';

export interface NativeDocumentRecord {
  schema: 'NativeDocumentRecord';
  version: 1;
  unitId: string;
  artifact: NativeDocumentArtifact;
  updatedAt: string;
}

function copyArtifact(artifact: NativeDocumentArtifact): NativeDocumentArtifact {
  return {
    ...artifact,
    sourceBytes: artifact.sourceBytes.slice(0),
    nativeGraph: structuredClone(artifact.nativeGraph),
    detectedFeatures: [...artifact.detectedFeatures],
    ownership: structuredClone(artifact.ownership),
    compatibility: structuredClone(artifact.compatibility),
  };
}

function copyRecord(record: NativeDocumentRecord): NativeDocumentRecord {
  return { ...record, artifact: copyArtifact(record.artifact) };
}

function compactArtifact(artifact: NativeDocumentArtifact): NativeDocumentArtifact {
  return {
    ...artifact,
    sourceBytes: artifact.sourceBytes.slice(0),
    nativeGraph: artifact.nativeGraph.kind === 'opc'
      ? { kind: 'opc' as const, package: { ...structuredClone(artifact.nativeGraph.package), parts: {}, opaqueParts: {}, contentTypesXml: artifact.nativeGraph.package.contentTypesXml?.slice() } }
      : structuredClone(artifact.nativeGraph),
    detectedFeatures: [...artifact.detectedFeatures],
    ownership: structuredClone(artifact.ownership),
    compatibility: structuredClone(artifact.compatibility),
  };
}

export async function buildNativeDocumentRecord(unitId: string, artifact: NativeDocumentArtifact): Promise<NativeDocumentRecord> {
  if (!unitId.trim()) throw new Error('Workbook unitId is required for a native document');
  await verifyNativeDocumentArtifact(artifact);
  return {
    schema: 'NativeDocumentRecord',
    version: 1,
    unitId,
    artifact: compactArtifact(artifact),
    updatedAt: new Date().toISOString(),
  };
}

/** Per-workbook session-memory native document artifact. */
export class LocalNativeDocumentStore {
  constructor(private readonly coordinator: WorkspaceMemoryCoordinator) {}

  async save(unitId: string, artifact: NativeDocumentArtifact): Promise<NativeDocumentRecord> {
    const record = await buildNativeDocumentRecord(unitId, artifact);
    return this.coordinator.transaction((transaction) => {
      transaction.set('nativeDocuments', unitId, copyRecord(record));
      return copyRecord(record);
    });
  }

  async load(unitId: string): Promise<NativeDocumentArtifact | null> {
    return this.coordinator.read((transaction) => {
      const record = transaction.get<NativeDocumentRecord>('nativeDocuments', unitId);
      if (!record) return null;
      if (record.schema !== 'NativeDocumentRecord' || record.version !== 1 || record.unitId !== unitId) {
        throw new Error(`NATIVE_DOCUMENT_SCHEMA_INVALID: ${unitId}`);
      }
      return verifyNativeDocumentArtifact(record.artifact).then(() => {
        const nativeGraph = record.artifact.nativeGraph.kind === 'opc' && Object.keys(record.artifact.nativeGraph.package.parts).length === 0
          ? { kind: 'opc' as const, package: loadOpcPackageGraph(record.artifact.sourceBytes, {}, record.artifact.fileName).packageGraph }
          : record.artifact.nativeGraph;
        return copyArtifact({ ...record.artifact, nativeGraph });
      });
    });
  }

  async remove(unitId: string): Promise<void> {
    await this.coordinator.transaction((transaction) => transaction.delete('nativeDocuments', memoryKey(unitId)));
  }
}
