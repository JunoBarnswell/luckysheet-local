import { OOXML_CODEC_REVISION, type CompatibilityReport, type DateSystem, type ExcelDocumentFormat, type FeatureOwnershipResult, type NativePackageState, type OpcPackageGraph } from './types';
import { createCompatibilityReport } from './compatibility-report';

export async function createNativePackageState(input: {
  fileName: string;
  buffer: ArrayBuffer;
  dateSystem: DateSystem;
  packageGraph: OpcPackageGraph;
  format?: ExcelDocumentFormat;
  detectedFeatures: Iterable<string>;
  compatibility?: CompatibilityReport;
  ownership?: FeatureOwnershipResult[];
}): Promise<NativePackageState> {
  const buffer = input.buffer.slice(0);
  const compatibility = structuredClone(input.compatibility ?? createCompatibilityReport({ fileName: input.fileName, importLevel: 'B', exportLevel: 'B', dateSystem: input.dateSystem, detectedFeatures: input.detectedFeatures }));
  return {
    schema: 'NativePackageState',
    format: input.format ?? input.packageGraph.format,
    fileName: input.fileName,
    sourceBytes: buffer,
    checksum: await sha256Hex(buffer),
    dateSystem: input.dateSystem,
    detectedFeatures: [...new Set(input.detectedFeatures)],
    packageGraph: input.packageGraph,
    ownership: structuredClone(input.ownership ?? []),
    codecRevision: OOXML_CODEC_REVISION,
    compatibility,
  };
}

export async function verifyNativePackageState(state: NativePackageState): Promise<void> {
  if (state.schema !== 'NativePackageState') throw new Error('Invalid native package state schema');
  if (!Number.isSafeInteger(state.codecRevision) || state.codecRevision < 1) throw new Error('Invalid native codec revision');
  if (state.compatibility.schema !== 'CompatibilityReport') throw new Error('Invalid native compatibility report');
  if (state.packageGraph.schema !== 'OpcPackageGraph') throw new Error('Invalid OPC package graph');
  const actual = await sha256Hex(state.sourceBytes);
  if (actual !== state.checksum) throw new Error('Native package checksum does not match source bytes');
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error('Web Crypto SHA-256 is required for XLSX source artifacts');
  const digest = await subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}
