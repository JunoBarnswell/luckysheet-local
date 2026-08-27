import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { it } from 'node:test';
import '../../cell-edit/performance.test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exportSnapshotToXlsxBuffer, importXlsx } from '@react-sheets/exchange-excel-ooxml';
import { createSpreadsheetRuntime, disposeSpreadsheetRuntime, hydrateRuntime } from '../../runtime';

const ROW_COUNT = 4_058;
const COLUMN_COUNT = 23;

function attachmentScaleSnapshot() {
  const workbook = new WorkbookModel('xlsx-performance', 'XLSX performance');
  const sheet = workbook.getSheet('sheet-1');
  for (let row = 0; row <= ROW_COUNT; row += 1) {
    for (let column = 0; column < COLUMN_COUNT; column += 1) {
      sheet.cells.set(row, column, { value: row === 0 ? `Column ${column + 1}` : `${column}:${row % 408}` });
    }
  }
  return workbook.snapshot();
}

it('imports and hydrates an attachment-scale value-only workbook without duplicating ordinary cells into FormulaEngine', async () => {
  const buffer = exportSnapshotToXlsxBuffer(attachmentScaleSnapshot());
  const importStartedAt = performance.now();
  const imported = await importXlsx({ fileName: 'attachment-scale.xlsx', buffer, options: { compatibilityTarget: 'B' } });
  const importedAt = performance.now();
  const runtime = createSpreadsheetRuntime({ unitId: imported.snapshot.unitId });
  hydrateRuntime(runtime, { snapshot: imported.snapshot, revision: 0 });
  await runtime.formulaCalculation;
  const hydratedAt = performance.now();
  try {
    const sheet = runtime.model.getSheet('sheet-1');
    assert.equal(sheet.cells.count(), (ROW_COUNT + 1) * COLUMN_COUNT);
    assert.equal(sheet.cells.get(1, 1)?.value, '1:1');
    assert.equal(runtime.formula.getCellResult({ sheetId: sheet.id, row: 1, column: 1 }), undefined);
    assert.ok(importedAt - importStartedAt < 1_500, `attachment-scale XLSX import exceeded 1500ms: ${Math.round(importedAt - importStartedAt)}ms`);
    assert.ok(hydratedAt - importedAt < 800, `attachment-scale runtime hydration exceeded 800ms: ${Math.round(hydratedAt - importedAt)}ms`);
  } finally {
    disposeSpreadsheetRuntime(runtime);
  }
});
