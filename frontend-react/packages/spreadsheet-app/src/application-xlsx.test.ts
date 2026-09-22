import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exchangeExportDocument, exchangeImportDocument, exchangeSaveAsDocument, exchangeSaveDocument, summarizeCompatibilityReport } from './features/native-document';
import { WorkbookSession } from './workbook-session';

describe('xlsx exchange', () => {
  it('exports workbook snapshots with compatibility summary', async () => {
    const workbook = new WorkbookModel('wb-export', 'Export');
    workbook.getSheet(workbook.primarySheetId).cells.set(0, 0, { value: 99 });
    const result = await exchangeExportDocument(workbook.snapshot(), { fileName: 'export.xlsx', execution: 'inline-test' });
    assert.ok(result.buffer && result.buffer.byteLength > 0);
    assert.match(summarizeCompatibilityReport(result.report), /compatibility/i);
  });
});

describe('WorkbookSession xlsx integration', () => {
  it('exports through document.export command path', async () => {
    const app = new WorkbookSession({ nativeDocumentExecution: 'inline-test' });
    app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(),
      row: 0,
      column: 0,
      value: { value: 'xlsx' },
    });
    const exported = await app.exportDocument('demo.xlsx');
    assert.ok(exported?.buffer instanceof ArrayBuffer);
    assert.equal(app.getUiSnapshot().compatibilityReport?.fileName, 'demo.xlsx');
  });

  it('clears compatibility report from ui snapshot', () => {
    const app = new WorkbookSession();
    app['compatibilityReport'] = {
      schema: 'CompatibilityReport',
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

describe('native document Save semantics', () => {
  it('requires Save As for a different target file', async () => {
    const source = new WorkbookModel('save-source', 'Save source');
    const exported = await exchangeExportDocument(source.snapshot(), { fileName: 'source.csv', execution: 'inline-test' });
    const imported = await exchangeImportDocument({ fileName: 'source.csv', buffer: exported.buffer!, execution: 'inline-test' });
    await assert.rejects(exchangeSaveDocument(imported.snapshot!, imported.artifact, { fileName: 'other.csv', execution: 'inline-test' }), /NATIVE_DOCUMENT_SAVE_TARGET_MISMATCH/);
    const saveAs = await exchangeSaveAsDocument(imported.snapshot!, { fileName: 'other.ssjson', artifact: imported.artifact, execution: 'inline-test' });
    assert.equal(saveAs.fileName, 'other.ssjson');
    assert.equal(imported.artifact.fileName, 'source.csv');
  });
});
