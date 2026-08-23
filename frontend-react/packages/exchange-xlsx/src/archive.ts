import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import { buildXlsxArchiveBase64, exportSnapshotToXlsxXml, parseXlsxXmlToSnapshot } from '@react-sheets/pro-features';
import { inflateRawSync } from 'node:zlib';

export { buildXlsxArchiveBase64, exportSnapshotToXlsxXml, parseXlsxXmlToSnapshot };

export function unzipXlsxBase64(base64: string): Record<string, string> {
  const buffer = Buffer.from(base64.replace(/^data:[^;]+;base64,/, ''), 'base64');
  const files: Record<string, string> = {};
  let offset = 0;

  while (offset + 4 <= buffer.length) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x04034b50) break;
    const method = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const nameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);
    const nameStart = offset + 30;
    const name = buffer.subarray(nameStart, nameStart + nameLength).toString('utf8');
    const dataStart = nameStart + nameLength + extraLength;
    const rawData = buffer.subarray(dataStart, dataStart + compressedSize);
    try {
      files[name] = (method === 0 ? rawData : inflateRawSync(rawData)).toString('utf8');
    } catch {
      // skip unreadable binary parts
    }
    offset = dataStart + compressedSize;
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
