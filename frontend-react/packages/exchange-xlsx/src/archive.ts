import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { strFromU8, strToU8, zipSync } from 'fflate';
import {
  bytesToBase64,
  detectPackageFeatures,
  exportSnapshotToXlsxPackage,
  loadXlsxPackage,
  parseLoadedXlsx,
} from './ooxml';
import type { DateSystem, XlsxPackage, XlsxZipLimits } from './types';

export { bytesToBase64, detectPackageFeatures, loadXlsxPackage, parseLoadedXlsx };

/** Readable XML view retained for callers that need to inspect a package. */
export function unzipXlsxBase64(base64: string, limits?: Partial<XlsxZipLimits>): Record<string, string> {
  const loaded = loadXlsxPackage(base64, limits);
  const files: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(loaded.files)) files[name] = strFromU8(bytes);
  return files;
}

/** Parse an in-memory XML file map through the independent OOXML package reader. */
export function parseXlsxXmlToSnapshot(files: Record<string, string>): WorkbookSnapshot {
  const parts: Record<string, Uint8Array> = Object.fromEntries(Object.entries(files).map(([name, value]) => [name, strToU8(value)]));
  if (!parts['xl/workbook.xml']) throw new Error('Not a valid XLSX package: xl/workbook.xml is missing');
  const loaded = loadXlsxPackage(bytesToBase64(zipSync(parts, { level: 0 })));
  return parseLoadedXlsx(loaded).snapshot;
}

/** Export a snapshot as a real OOXML ZIP package. */
export function exportSnapshotToXlsxBase64(
  snapshot: WorkbookSnapshot,
  preserved?: XlsxPackage,
  options: { dateSystem?: DateSystem; includeCachedValues?: boolean; preserveMacros?: boolean } = {},
): string {
  return exportSnapshotToXlsxPackage(snapshot, {
    dateSystem: options.dateSystem ?? preserved?.dateSystem ?? '1900',
    includeCachedValues: options.includeCachedValues,
    preserveMacros: options.preserveMacros ?? true,
  }, preserved);
}

/** Build a package for callers that need a standalone byte payload. */
export function zipXlsxParts(parts: Record<string, Uint8Array>): string {
  return bytesToBase64(zipSync(parts, { level: 6 }));
}
