import type { NativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
import { loadOpcPackageGraph, migrateNativeDocumentArtifactV1, verifyNativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
import { memoryKey, type WorkspaceMemoryCoordinator } from './memory';

export interface NativeDocumentRecord {
  schema: 'NativeDocumentRecord';
  version: 2;
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
    version: 2,
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
    return this.coordinator.transaction(async (transaction) => {
      const record = transaction.get<unknown>('nativeDocuments', unitId);
      if (record === undefined) return null;
      if (record === null || typeof record !== 'object' || Array.isArray(record)
        || !('schema' in record) || record.schema !== 'NativeDocumentRecord'
        || !('unitId' in record) || record.unitId !== unitId
        || !('artifact' in record) || !record.artifact || typeof record.artifact !== 'object' || Array.isArray(record.artifact)
        || !('updatedAt' in record) || typeof record.updatedAt !== 'string') {
        throw new Error(`NATIVE_DOCUMENT_SCHEMA_INVALID: ${unitId}`);
      }
      let canonical: NativeDocumentRecord;
      if ('version' in record && record.version === 1) {
        const artifact = await migrateNativeDocumentArtifactV1(record.artifact as NativeDocumentArtifact);
        canonical = { schema: 'NativeDocumentRecord', version: 2, unitId, artifact, updatedAt: record.updatedAt };
        transaction.set('nativeDocuments', unitId, copyRecord(canonical));
      } else if ('version' in record && record.version === 2) {
        canonical = record as NativeDocumentRecord;
        await verifyNativeDocumentArtifact(canonical.artifact);
      } else {
        throw new Error(`NATIVE_DOCUMENT_SCHEMA_INVALID: ${unitId}`);
      }
      const nativeGraph = canonical.artifact.nativeGraph.kind === 'opc' && Object.keys(canonical.artifact.nativeGraph.package.parts).length === 0
        ? { kind: 'opc' as const, package: loadOpcPackageGraph(canonical.artifact.sourceBytes, {}, canonical.artifact.fileName).packageGraph }
        : canonical.artifact.nativeGraph;
      return copyArtifact({ ...canonical.artifact, nativeGraph });
    });
  }

  async remove(unitId: string): Promise<void> {
    await this.coordinator.transaction((transaction) => transaction.delete('nativeDocuments', memoryKey(unitId)));
  }
}
