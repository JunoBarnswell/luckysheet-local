import type { NativeDocumentArtifact } from './types';

export function cloneNativeDocumentArtifact(artifact: NativeDocumentArtifact): NativeDocumentArtifact {
  assertNativeDocumentArtifact(artifact);
  return structuredClone(artifact);
}

export function assertNativeDocumentArtifact(value: unknown): asserts value is NativeDocumentArtifact {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('NATIVE_DOCUMENT_ARTIFACT_INVALID: artifact must be an object');
  const artifact = value as Record<string, unknown>;
  if (artifact.schema !== 'NativeDocumentArtifact'
    || typeof artifact.unitId !== 'string' || !artifact.unitId.trim()
    || typeof artifact.fileName !== 'string' || !artifact.fileName.trim()
    || typeof artifact.checksum !== 'string' || !/^[a-f0-9]{64}$/i.test(artifact.checksum)
    || !Number.isSafeInteger(artifact.byteLength) || Number(artifact.byteLength) < 0
    || !Number.isSafeInteger(artifact.sourceRevision) || Number(artifact.sourceRevision) < 0
    || !Number.isSafeInteger(artifact.codecRevision) || Number(artifact.codecRevision) < 1
    || !Array.isArray(artifact.detectedFeatures)
    || !artifact.compatibility || typeof artifact.compatibility !== 'object') {
    throw new Error('NATIVE_DOCUMENT_ARTIFACT_INVALID: artifact identity or metadata is incomplete');
  }
}
