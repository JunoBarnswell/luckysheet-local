import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { strFromU8, strToU8, zipSync } from 'fflate';
import {
  detectPackageFeatures,
  exportSnapshotToOpcPackageGraph,
  loadOpcPackageGraph,
  parseLoadedXlsx,
} from './ooxml';
import type { DateSystem, OpcPackageGraph, XlsxZipLimits } from './types';

export { detectPackageFeatures, loadOpcPackageGraph, parseLoadedXlsx };

/** Base64 is retained only as an explicit Node/test compatibility helper. */
export function bytesToBase64(bytes: Uint8Array): string {
  if (typeof btoa === 'function') {
    let binary = '';
    for (let index = 0; index < bytes.length; index += 0x8000) {
      binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
    }
    return btoa(binary);
  }
  const bufferCtor = (globalThis as { Buffer?: { from(value: Uint8Array): { toString(encoding: string): string } } }).Buffer;
  if (!bufferCtor) throw new Error('Base64 encoding is unavailable in this host');
  return bufferCtor.from(bytes).toString('base64');
}

function base64ToBytes(value: string): Uint8Array {
  const normalized = value.replace(/^data:[^;]+;base64,/, '').replace(/\s+/g, '');
  if (typeof atob === 'function') {
    const binary = atob(normalized);
    const result = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) result[index] = binary.charCodeAt(index);
    return result;
  }
  const bufferCtor = (globalThis as { Buffer?: { from(value: string, encoding: string): Uint8Array } }).Buffer;
  if (!bufferCtor) throw new Error('Base64 decoding is unavailable in this host');
  return new Uint8Array(bufferCtor.from(normalized, 'base64'));
}

/** Readable XML view retained for callers that need to inspect a package. */
export function unzipXlsxBase64(base64: string, limits?: Partial<XlsxZipLimits>): Record<string, string> {
  const loaded = loadOpcPackageGraph(base64ToBytes(base64), limits);
  const files: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(loaded.files)) files[name] = strFromU8(bytes);
  return files;
}

/** Parse an in-memory XML file map through the independent OOXML package reader. */
export function parseXlsxXmlToSnapshot(files: Record<string, string>): WorkbookSnapshot {
  const parts: Record<string, Uint8Array> = Object.fromEntries(Object.entries(files).map(([name, value]) => [name, strToU8(value)]));
  if (!parts['xl/workbook.xml']) throw new Error('Not a valid XLSX package: xl/workbook.xml is missing');
  const loaded = loadOpcPackageGraph(zipSync(parts, { level: 0 }));
  return parseLoadedXlsx(loaded).snapshot;
}

/** Export a snapshot as a real OOXML ZIP package. */
export function exportSnapshotToXlsxBase64(
  snapshot: WorkbookSnapshot,
  preserved?: OpcPackageGraph,
  options: { dateSystem?: DateSystem; includeCachedValues?: boolean; preserveMacros?: boolean; assetBytes?: Record<string, Uint8Array> } = {},
): string {
  return bytesToBase64(new Uint8Array(exportSnapshotToOpcPackageGraph(snapshot, {
    dateSystem: options.dateSystem ?? preserved?.dateSystem ?? '1900',
    includeCachedValues: options.includeCachedValues,
    preserveMacros: options.preserveMacros ?? true,
    assetBytes: options.assetBytes,
  }, preserved)));
}

/** Export a standalone OOXML package without converting it through base64. */
export function exportSnapshotToXlsxBuffer(
  snapshot: WorkbookSnapshot,
  preserved?: OpcPackageGraph,
  options: { dateSystem?: DateSystem; includeCachedValues?: boolean; preserveMacros?: boolean; assetBytes?: Record<string, Uint8Array> } = {},
): ArrayBuffer {
  return exportSnapshotToOpcPackageGraph(snapshot, {
    dateSystem: options.dateSystem ?? preserved?.dateSystem ?? '1900',
    includeCachedValues: options.includeCachedValues,
    preserveMacros: options.preserveMacros ?? true,
    assetBytes: options.assetBytes,
  }, preserved);
}

/** Build a package for callers that need a standalone byte payload. */
export function zipXlsxParts(parts: Record<string, Uint8Array>): string {
  return bytesToBase64(zipSync(parts, { level: 6 }));
}

export function zipXlsxPartsBuffer(parts: Record<string, Uint8Array>): ArrayBuffer {
  const zipped = zipSync(parts, { level: 6 });
  return zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength) as ArrayBuffer;
}
