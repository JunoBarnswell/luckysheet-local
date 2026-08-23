import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exportXlsx } from './export';
import { importXlsx } from './import';
import { scanSnapshotFeatures } from './feature-scan';
import { exportSnapshotToXlsxBase64 } from './archive';

describe('exchange-xlsx', () => {
  it('scans workbook features for compatibility reporting', () => {
    const workbook = new WorkbookModel('wb-xlsx', 'XLSX');
    const sheet = workbook.getSheet(workbook.activeSheetId);
    sheet.cells.set(0, 0, { value: 10, formula: '=SUM(A2:A3)' });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 },
      anchor: { row: 1, column: 0 },
    });
    const features = scanSnapshotFeatures(workbook.snapshot());
    assert.ok(features.includes('cells'));
    assert.ok(features.includes('formulas'));
    assert.ok(features.includes('merges'));
  });

  it('round-trips snapshot through xlsx archive import/export', async () => {
    const workbook = new WorkbookModel('wb-roundtrip', 'Roundtrip');
    workbook.getSheet(workbook.activeSheetId).cells.set(0, 0, { value: 'hello' });
    workbook.getSheet(workbook.activeSheetId).cells.set(1, 0, { value: 42, formula: '=A1&"!"' });
    const snapshot = workbook.snapshot();
    const base64 = exportSnapshotToXlsxBase64(snapshot);
    const imported = await importXlsx({
      fileName: 'roundtrip.xlsx',
      base64,
      options: { compatibilityTarget: 'B' },
    });
    assert.equal(imported.snapshot.sheets.length, snapshot.sheets.length);
    assert.equal(imported.report.schema, 'CompatibilityReportV1');
    const exported = await exportXlsx({
      snapshot: imported.snapshot,
      fileName: 'roundtrip.xlsx',
      options: { compatibilityTarget: 'B' },
    });
    assert.ok(exported.base64.length > 0);
    assert.equal(exported.fileName, 'roundtrip.xlsx');
  });
});
