import { XLSX_CODEC_VERSION, type CompatibilityReport, type DateSystem, type XlsxSourceArtifact } from './types';
import { createCompatibilityReport } from './compatibility-report';

export async function createXlsxSourceArtifact(input: {
  fileName: string;
  buffer: ArrayBuffer;
  dateSystem: DateSystem;
  detectedFeatures: Iterable<string>;
  capabilityReport?: CompatibilityReport;
}): Promise<XlsxSourceArtifact> {
  const buffer = input.buffer.slice(0);
  return {
    schema: 'XlsxSourceArtifact',
    fileName: input.fileName,
    buffer,
    checksum: await sha256Hex(buffer),
    dateSystem: input.dateSystem,
    detectedFeatures: [...new Set(input.detectedFeatures)],
    xlsxCodecVersion: XLSX_CODEC_VERSION,
    capabilityReport: structuredClone(input.capabilityReport ?? createCompatibilityReport({ fileName: input.fileName, importLevel: 'B', exportLevel: 'B', dateSystem: input.dateSystem, detectedFeatures: input.detectedFeatures })),
  };
}

export async function verifyXlsxSourceArtifact(artifact: XlsxSourceArtifact): Promise<void> {
  if (artifact.schema !== 'XlsxSourceArtifact') throw new Error('Invalid XLSX source artifact schema');
  if (artifact.xlsxCodecVersion !== undefined && (!Number.isSafeInteger(artifact.xlsxCodecVersion) || artifact.xlsxCodecVersion < 1)) throw new Error('Invalid XLSX codec version');
  if (artifact.capabilityReport !== undefined && artifact.capabilityReport.schema !== 'CompatibilityReport') throw new Error('Invalid XLSX capability report');
  const actual = await sha256Hex(artifact.buffer);
  if (actual !== artifact.checksum) throw new Error('XLSX source artifact checksum does not match its buffer');
}

export function xlsxArtifactNeedsLayoutRepair(artifact: XlsxSourceArtifact): boolean {
  return (artifact.xlsxCodecVersion ?? 1) < XLSX_CODEC_VERSION;
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error('Web Crypto SHA-256 is required for XLSX source artifacts');
  const digest = await subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}
