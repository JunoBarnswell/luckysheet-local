import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import { buildXlsxArchiveBase64, exportSnapshotToXlsxXml, parseXlsxXmlToSnapshot } from '@react-sheets/pro-features';
import { strFromU8, unzipSync } from 'fflate';

export { buildXlsxArchiveBase64, exportSnapshotToXlsxXml, parseXlsxXmlToSnapshot };

function decodeBase64(base64: string): Uint8Array {
  const normalized = base64.replace(/^data:[^;]+;base64,/, '');
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(normalized, 'base64'));
  const binary = atob(normalized);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

export function unzipXlsxBase64(base64: string): Record<string, string> {
  const entries = unzipSync(decodeBase64(base64));
  const files: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(entries)) {
    files[name] = strFromU8(bytes);
  }
  return files;
}

export function parseXlsxBase64ToSnapshot(base64: string): WorkbookSnapshotV1 {
  const files = unzipXlsxBase64(base64);
  if (!files['xl/workbook.xml']) {
    throw new Error('Not a valid XLSX package');
  }
  return parseXlsxXmlToSnapshot(files);
}

export function exportSnapshotToXlsxBase64(snapshot: WorkbookSnapshotV1): string {
  return buildXlsxArchiveBase64(snapshot);
}
