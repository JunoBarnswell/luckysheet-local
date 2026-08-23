import type { DateSystem, XlsxSourceArtifact } from './types';

export async function createXlsxSourceArtifact(input: {
  fileName: string;
  buffer: ArrayBuffer;
  dateSystem: DateSystem;
  detectedFeatures: Iterable<string>;
}): Promise<XlsxSourceArtifact> {
  const buffer = input.buffer.slice(0);
  return {
    schema: 'XlsxSourceArtifact',
    fileName: input.fileName,
    buffer,
    checksum: await sha256Hex(buffer),
    dateSystem: input.dateSystem,
    detectedFeatures: [...new Set(input.detectedFeatures)],
  };
}

export async function verifyXlsxSourceArtifact(artifact: XlsxSourceArtifact): Promise<void> {
  if (artifact.schema !== 'XlsxSourceArtifact') throw new Error('Invalid XLSX source artifact schema');
  const actual = await sha256Hex(artifact.buffer);
  if (actual !== artifact.checksum) throw new Error('XLSX source artifact checksum does not match its buffer');
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error('Web Crypto SHA-256 is required for XLSX source artifacts');
  const digest = await subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}
