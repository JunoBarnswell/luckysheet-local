import assert from 'node:assert/strict';
import test from 'node:test';
import { registerSheetCommands, normalizeAutoFilterModel } from './index';
import { compareSortValues, resolveSortCellValue } from './data-features';
import { openCanonicalTestRuntime } from '../../core-model/src/canonical-test-runtime.test';

test('sorting consumes canonical resolved values and commits a structural permutation', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('m3-m4-sort');
  try {
    registerSheetCommands(runtime);
    const sheetId = workbook.primarySheetId;
    await runtime.execute('sheet.range.set', {
      sheetId, startRow: 0, startColumn: 0,
      values: [[{ value: 'Calculated' }, { value: 'Row' }], [{ value: 20 }, { value: 'twenty' }], [{ value: 5 }, { value: 'five' }], [{ value: 10 }, { value: 'ten' }]],
    });
    await runtime.execute('data.sort.rows', {
      sheetId, range: { sheetId, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }, criteria: [{ column: 0, ascending: true }], hasHeader: true,
    });
    assert.deepEqual([1, 2, 3].map((row) => workbook.getSheet(sheetId).cells.get(row, 1)?.value), ['five', 'ten', 'twenty']);
    assert.equal(runtime.getUndoEntries().at(-1)?.semanticCommandDescriptor.id, 'data.sort.rows');
  } finally {
    close();
  }
});

test('sort value resolution rejects formulas without a canonical result', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('m3-m4-sort-reject');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet(workbook.primarySheetId);
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 1, column: 0, value: { value: null, formula: '=A1' } });
    assert.throws(() => resolveSortCellValue(sheet, 1, 0), /formula result unavailable/);
    assert.equal(compareSortValues(2, 10) < 0, true);
  } finally {
    close();
  }
});

test('worksheet filter normalization remains a typed canonical range contract', () => {
  const filter = normalizeAutoFilterModel({ sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }, columns: {} });
  assert.deepEqual(filter.range, { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
  assert.deepEqual(filter.columns, {});
});
