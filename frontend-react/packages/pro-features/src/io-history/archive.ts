import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import { exportSnapshotToXlsxXml } from './xlsx';
import { bytesToBase64, createZipStore } from './zip';

/**
 * 将工作簿快照导出为真实 XLSX(zip 容器)的 Base64:
 * 复用 XML 工作表序列化,以 STORE zip 打包为合法 OOXML 包。
 */
export function buildXlsxArchiveBase64(snapshot: WorkbookSnapshotV1): string {
  const files = exportSnapshotToXlsxXml(snapshot);
  const encoder = new TextEncoder();
  const entries = Object.entries(files).map(([name, content]) => ({
    name,
    data: encoder.encode(content),
  }));
  return bytesToBase64(createZipStore(entries));
}
