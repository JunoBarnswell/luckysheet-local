import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { NATIVE_DOCUMENT_CODEC_REVISION, type CompatibilityReport, type DateSystem, type NativeDocumentFormat, type FeatureOwnershipResult, type NativeDocumentArtifact, type NativeGraph, type NativeSecurityEnvelope } from './types';
import { createCompatibilityReport } from './compatibility-report';
import { capabilityFor } from './capability-manifest';

export async function createNativeDocumentArtifact(input: {
  fileName: string;
  buffer: ArrayBuffer;
  dateSystem: DateSystem;
  nativeGraph: NativeGraph;
  snapshot?: WorkbookSnapshot;
  securityEnvelope?: NativeSecurityEnvelope;
  format?: NativeDocumentFormat;
  detectedFeatures: Iterable<string>;
  compatibility?: CompatibilityReport;
  ownership?: FeatureOwnershipResult[];
}): Promise<NativeDocumentArtifact> {
  const buffer = input.buffer.slice(0);
  const detectedFeatures = [...new Set(input.detectedFeatures)];
  const compatibility = structuredClone(input.compatibility ?? createCompatibilityReport({ fileName: input.fileName, importLevel: 'B', exportLevel: 'B', dateSystem: input.dateSystem, detectedFeatures }));
  return {
    schema: 'NativeDocumentArtifact',
    format: input.format ?? formatFromGraph(input.nativeGraph),
    fileName: input.fileName,
    sourceBytes: buffer,
    checksum: await sha256Hex(buffer),
    ...(input.securityEnvelope ? { securityEnvelope: structuredClone(input.securityEnvelope) } : {}),
    dateSystem: input.dateSystem,
    detectedFeatures,
    nativeGraph: input.nativeGraph,
    ...(input.snapshot ? { sourceSnapshotHash: nativeSnapshotHash(input.snapshot) } : {}),
    ownership: structuredClone(input.ownership ?? detectedFeatures.map((feature) => ownershipFor(feature))),
    codecRevision: NATIVE_DOCUMENT_CODEC_REVISION,
    compatibility,
  };
}

function ownershipFor(feature: string): FeatureOwnershipResult {
  const capability = capabilityFor(feature);
  const ownership = capability.edit === 'none' && capability.write === 'none' && capability.preserve !== 'none'
    ? 'preserved-owned'
    : capability.edit === 'full' && capability.write === 'full'
      ? 'editable-owned'
      : 'mixed-owned';
  return {
    feature,
    scope: `document#${feature}`,
    read: capability.read,
    edit: capability.edit,
    write: capability.write,
    preserve: capability.preserve,
    ownership,
  };
}

function formatFromGraph(graph: NativeGraph): NativeDocumentFormat {
  switch (graph.kind) {
    case 'opc': return graph.package.format;
    case 'text': return { family: 'text', variant: graph.dialect.variant };
    case 'xml': return { family: 'xmlss', variant: 'xml' };
    case 'ods': return { family: 'ods', variant: 'ods' };
    case 'sjs': return { family: 'sjs', variant: 'sjs' };
    case 'ssjson': return { family: 'ssjson', variant: 'ssjson' };
    case 'xlsb': return { family: 'xlsb', variant: 'xlsb' };
    case 'biff': return { family: 'biff', variant: 'xls' };
    case 'dbf': return { family: 'dbf', variant: 'dbf' };
  }
}

export async function verifyNativeDocumentArtifact(state: NativeDocumentArtifact): Promise<void> {
  if (state.schema !== 'NativeDocumentArtifact') throw new Error('Invalid native document artifact schema');
  if (!state.fileName || !(state.sourceBytes instanceof ArrayBuffer) || !/^[a-f0-9]{64}$/i.test(state.checksum)) throw new Error('Invalid native document artifact identity');
  if (!state.format || typeof state.format !== 'object' || !state.format.family) throw new Error('Invalid native document format');
  if (!Number.isSafeInteger(state.codecRevision) || state.codecRevision < 1) throw new Error('Invalid native codec revision');
  if (state.compatibility.schema !== 'CompatibilityReport') throw new Error('Invalid native compatibility report');
  if (!state.nativeGraph || typeof state.nativeGraph !== 'object' || !('kind' in state.nativeGraph)) throw new Error('Invalid native document graph');
  if (!['opc', 'text', 'xml', 'ods', 'sjs', 'ssjson', 'biff', 'xlsb', 'dbf'].includes(state.nativeGraph.kind)) throw new Error('Invalid native document graph kind');
  if (state.nativeGraph.kind === 'opc' && state.nativeGraph.package.schema !== 'OpcPackageGraph') throw new Error('Invalid OPC package graph');
  if (state.nativeGraph.kind === 'biff') {
    const graph = state.nativeGraph.container;
    if (graph.container !== 'cfb' || !graph.cfb || !graph.streamName || !graph.streams || !graph.sheets?.length) throw new Error('Invalid BIFF binary graph');
    if (!graph.streams[graph.streamName] || graph.cfb.entries.length === 0) throw new Error('Invalid BIFF binary stream graph');
  }
  if (state.nativeGraph.kind === 'xlsb') {
    const graph = state.nativeGraph.container;
    if (graph.container !== 'biff12' || !graph.package || !graph.package.workbookPart || !graph.package.parts[graph.package.workbookPart] || !graph.sheets?.length) throw new Error('Invalid XLSB binary graph');
  }
  const graphFamily = state.nativeGraph.kind === 'opc' ? state.nativeGraph.package.format.family : state.nativeGraph.kind === 'xml' ? 'xmlss' : state.nativeGraph.kind;
  if (state.format.family !== graphFamily) throw new Error(`Native document format/graph mismatch: ${state.format.family}/${graphFamily}`);
  if (!Array.isArray(state.ownership) || state.ownership.some((entry) => !entry || typeof entry.feature !== 'string' || !['full', 'partial', 'none'].includes(entry.read) || !['full', 'partial', 'none'].includes(entry.edit) || !['full', 'partial', 'none'].includes(entry.write) || !['full', 'partial', 'none'].includes(entry.preserve) || !['editable-owned', 'preserved-owned', 'mixed-owned'].includes(entry.ownership))) throw new Error('Invalid native document ownership manifest');
  if (state.sourceSnapshotHash !== undefined && !/^fnv1a-[a-f0-9]{8}$/i.test(state.sourceSnapshotHash)) throw new Error('Invalid native document source snapshot identity');
  const actual = await sha256Hex(state.sourceBytes);
  if (actual !== state.checksum) throw new Error('Native document checksum does not match source bytes');
}

/** Compact deterministic identity for proving an untouched Save. */
export function nativeSnapshotHash(snapshot: WorkbookSnapshot): string {
  const identity = structuredClone(snapshot);
  identity.unitId = '';
  identity.printDocuments = (identity.printDocuments ?? []).map((document) => ({ ...document, unitId: '' }));
  return nativeSnapshotHashValue(JSON.stringify(identity));
}

function nativeSnapshotHashValue(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a-${(hash >>> 0).toString(16).padStart(8, '0')}`;
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error('Web Crypto SHA-256 is required for native document artifacts');
  const digest = await subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}
