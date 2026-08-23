import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exchangeExportXlsx, summarizeCompatibilityReport } from './features/xlsx';
import { SpreadsheetApplication } from './application';

describe('xlsx exchange', () => {
  it('exports workbook snapshots with compatibility summary', async () => {
    const workbook = new WorkbookModel('wb-export', 'Export');
    workbook.getSheet(workbook.activeSheetId).cells.set(0, 0, { value: 99 });
    const result = await exchangeExportXlsx(workbook.snapshot(), { fileName: 'export.xlsx' });
    assert.ok(result.base64 && result.base64.length > 0);
    assert.match(summarizeCompatibilityReport(result.report), /compatibility/i);
  });
});

describe('SpreadsheetApplication xlsx integration', () => {
  it('exports through xlsx.export command path', async () => {
    const app = new SpreadsheetApplication();
    app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(),
      row: 0,
      column: 0,
      value: { value: 'xlsx' },
    });
    const exported = await app.exportXlsxWorkbook('demo.xlsx');
    assert.ok(exported?.base64);
    assert.equal(app.getUiSnapshot().compatibilityReport?.fileName, 'demo.xlsx');
  });

  it('clears compatibility report from ui snapshot', () => {
    const app = new SpreadsheetApplication();
    app['compatibilityReport'] = {
      schema: 'CompatibilityReportV1',
      fileName: 'demo.xlsx',
      importLevel: 'B',
      exportLevel: 'B',
      dateSystem: '1900',
      issues: [],
      summary: { editableFeatures: 0, preservedOnly: 0, unsupported: 0 },
    };
    app.clearCompatibilityReport();
    assert.equal(app.getUiSnapshot().compatibilityReport, null);
  });
});
