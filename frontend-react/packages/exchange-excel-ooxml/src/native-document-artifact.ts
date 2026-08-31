import { sha256Hex as canonicalSha256Hex, type WorkbookSnapshot } from '@react-sheets/core-model';
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
  if (state.nativeGraph.kind === 'opc') {
    const pkg = state.nativeGraph.package;
    if (pkg.schema !== 'OpcPackageGraph' || !pkg.assetPartById || typeof pkg.assetPartById !== 'object' || Array.isArray(pkg.assetPartById)) throw new Error('Invalid OPC package graph');
    for (const [assetId, part] of Object.entries(pkg.assetPartById)) {
      if (!assetId.trim() || !part.trim() || !pkg.parts[part]) throw new Error(`Invalid OPC asset part ownership: ${assetId}`);
    }
  }
  if (state.nativeGraph.kind === 'biff') {
    const graph = state.nativeGraph.container;
    if (graph.container !== 'cfb' || !graph.cfb || !graph.streamName || !graph.streams || !graph.sheets?.length) throw new Error('Invalid BIFF binary graph');
    if (!graph.streams[graph.streamName] || graph.cfb.entries.length === 0) throw new Error('Invalid BIFF binary stream graph');
  }
  if (state.nativeGraph.kind === 'xlsb') {
    const graph = state.nativeGraph.container;
    if (graph.container !== 'biff12' || !graph.package || !graph.package.workbookPart || !graph.package.parts[graph.package.workbookPart] || !graph.sheets?.length) throw new Error('Invalid XLSB binary graph');
  }
  if (state.nativeGraph.kind === 'opc' && state.nativeGraph.package.nativeDrawingGraph) {
    const graph = state.nativeGraph.package.nativeDrawingGraph;
    if (graph.schema !== 'NativeDrawingGraph' || !Array.isArray(graph.nodes)) throw new Error('Invalid native DrawingML ownership graph');
    const identities = new Set<string>();
    for (const node of graph.nodes) {
      if (!node || typeof node.drawingPart !== 'string' || !node.drawingPart.trim()
        || !Number.isSafeInteger(node.nativeObjectId) || node.nativeObjectId <= 0
        || !['image', 'shape', 'textbox', 'connector', 'chart', 'unknown'].includes(node.kind)
        || !['editable-owned', 'preserved-owned'].includes(node.ownership)) throw new Error('Invalid native DrawingML ownership node');
      const identity = `${node.drawingPart}:${node.nativeObjectId}`;
      if (identities.has(identity)) throw new Error(`Duplicate native DrawingML ownership node: ${identity}`);
      identities.add(identity);
      if (node.drawingId !== undefined && !node.drawingId.trim()) throw new Error(`Invalid native DrawingML canonical identity: ${identity}`);
      if (node.ownership === 'editable-owned' && !node.drawingId) throw new Error(`Editable DrawingML node has no canonical owner: ${identity}`);
    }
  }
  if (state.nativeGraph.kind === 'opc' && state.nativeGraph.package.nativeReviewGraph) {
    const graph = state.nativeGraph.package.nativeReviewGraph;
    if (graph.schema !== 'NativeReviewGraph' || !Array.isArray(graph.sheets)) throw new Error('Invalid native review ownership graph');
    const sheetParts = new Set<string>();
    for (const entry of graph.sheets) {
      if (!entry || typeof entry.sheetPart !== 'string' || !entry.sheetPart.trim()
        || (entry.commentsPart !== undefined && !entry.commentsPart.trim())
        || (entry.threadedCommentsPart !== undefined && !entry.threadedCommentsPart.trim())
        || (entry.personsPart !== undefined && !entry.personsPart.trim())) throw new Error('Invalid native review ownership entry');
      if (sheetParts.has(entry.sheetPart)) throw new Error(`Duplicate native review ownership sheet: ${entry.sheetPart}`);
      sheetParts.add(entry.sheetPart);
    }
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
  return canonicalSha256Hex(new Uint8Array(buffer));
}
