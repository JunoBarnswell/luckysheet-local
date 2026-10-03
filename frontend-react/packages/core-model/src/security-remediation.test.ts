import test from 'node:test';
import assert from 'node:assert/strict';
import { assertWorkbookResourceBudget } from './resource-budget';
import { buildPivotTimelineTiles } from './pivot';
import { isTableSheetDefinition } from './table-sheet-validation';
import { compileExcelWildcard } from '../../formula-engine/src/wildcard';
import { parseSections } from '../../number-format/src/index';

test('canonical resource boundary accepts normal data and rejects nested/text budget bypasses', () => {
  assert.doesNotThrow(() => assertWorkbookResourceBudget({ text: 'hello', fontFamily: 'Calibri', width: 480 }));
  assert.throws(() => assertWorkbookResourceBudget({ text: 'x'.repeat(32768) }), /budget/);
  assert.throws(() => assertWorkbookResourceBudget({ rowHeightsPx: { '1': 1e9 } }), /budget/);
  assert.throws(() => assertWorkbookResourceBudget({ style: { numberFormat: {} } }), /string/);
  assert.throws(() => assertWorkbookResourceBudget({ fontFamily: 'Calibri\nInjected' }), /control/);
  assert.throws(() => assertWorkbookResourceBudget({ editor: { kind: 'checkbox', trueValue: 1, falseValue: 1 } }), /editor/);
});
test('wildcards handle escaped literals and do not use an unbounded backtracking regex', () => {
  assert.equal(compileExcelWildcard('a*?c')('ABBC'), true);
  assert.equal(compileExcelWildcard('~*~?')('*?'), true);
  assert.equal(compileExcelWildcard('a*c')('ab'), false);
  const attack = compileExcelWildcard('a*'.repeat(200) + 'Z');
  assert.equal(attack('a'.repeat(2000)), false);
  assert.throws(() => compileExcelWildcard('x'.repeat(32768)), /budget/);
});
test('timeline expansion applies a narrow explicit window before generating periods', () => {
  const tiles = buildPivotTimelineTiles(['1900-01-01', '9999-01-01'], 'days', { start: '2024-01-01', end: '2024-01-03' });
  assert.equal(tiles.length, 3);
  assert.equal(tiles[0]!.start, '2024-01-01');
  assert.throws(() => buildPivotTimelineTiles(['1900-01-01', '9999-01-01'], 'days'), /bound/);
});
test('TableSheet ingestion and number formats reject crash payloads', () => {
  assert.equal(isTableSheetDefinition({ viewId: 'v', columns: [{ fieldId: 'a', caption: 'A' }], grouping: [] }), true);
  assert.equal(isTableSheetDefinition({ viewId: 'v' }), false);
  assert.equal(isTableSheetDefinition({ viewId: 'v', columns: [], grouping: [{ fieldId: 'missing' }] }), false);
  assert.doesNotThrow(() => parseSections('#,##0.00'));
  assert.throws(() => parseSections('0'.repeat(256)), /255/);
});

test('chart vectors reject oversized and aggregate range bypasses before reading', async () => {
  const { assertChartProjectionBudget } = await import('../../spreadsheet-app/src/features/chart/data');
  const range = { sheetId: 's', startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 };
  assert.doesNotThrow(() => assertChartProjectionBudget([range]));
  assert.throws(() => assertChartProjectionBudget([{ ...range, endRow: 100000 }]), /budget/);
  assert.throws(() => assertChartProjectionBudget([{ ...range, endRow: 60000 }, { ...range, startColumn: 1, endColumn: 1, endRow: 60000 }]), /budget/);
});
