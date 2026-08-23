import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel, normalizePivotDefinition, type PivotModel } from '@react-sheets/core-model';
import { computePivotResult } from './pivot-engine';
import { buildPivotWriteback } from './pivot-write';

function workbook(): WorkbookModel {
  const model = new WorkbookModel('pivot-engine-test', 'Pivot'); const sheet = model.getSheet('sheet-1');
  ['Region', 'Year', 'Amount'].forEach((value, column) => sheet.cells.set(0, column, { value }));
  [['East', 2024, 1], ['East', 2025, 2], ['West', 2024, 2], ['West', 2025, 4], [null, 2025, null]].forEach((row, index) => row.forEach((value, column) => sheet.cells.set(index + 1, column, { value })));
  return model;
}
function model(values: PivotModel['valueFields']): PivotModel { return { id: 'p1', sheetId: 'sheet-1', sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }, rowFields: [], columnFields: [], valueFields: values, filterFields: [] }; }

describe('pivot domain engine', () => {
  it('normalizes the legacy serialized field arrays once at the domain boundary', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]); const normalized = normalizePivotDefinition(pivot);
    assert.deepEqual(normalized.layout.rows, []); assert.deepEqual(normalized.layout.values, pivot.valueFields); assert.equal('rowFields' in normalized, false);
  });

  it('round-trips the serializable advanced definition through workbook snapshots', () => {
    const modelData = workbook(); const pivot = model([{ field: 'Amount', summarizeBy: 'distinct-count' }]); pivot.layout = { rows: [{ field: 'Region', group: { kind: 'manual', groups: [{ name: 'Domestic', items: ['East', 'West'] }] } }], columns: [], filters: [], values: pivot.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false }; pivot.refreshPolicy = { mode: 'on-change', preserveFormatting: true, refreshOnLoad: false }; modelData.getSheet('sheet-1').pivots.push(pivot);
    const restored = WorkbookModel.fromSnapshot(modelData.snapshot()); assert.deepEqual(restored.getSheet('sheet-1').pivots[0]?.layout, pivot.layout); assert.equal(restored.getSheet('sheet-1').pivots[0]?.refreshPolicy?.mode, 'on-change');
  });

  it('calculates every supported aggregate independently', () => {
    const operations = ['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'] as const;
    const result = computePivotResult(workbook(), model(operations.map((summarizeBy) => ({ field: 'Amount', summarizeBy }))));
    assert.deepEqual(result.grandTotal?.values.slice(0, 7), [9, 4, 4, 2.25, 1, 4, 16]);
    assert.equal(result.grandTotal?.values[11], 3);
    assert.ok(Math.abs(Number(result.grandTotal?.values[7]) - Math.sqrt(19 / 12)) < 0.00001);
  });

  it('builds nested rows, column paths, source paths, grouping, filtering and sorting', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]); pivot.layout = { rows: [{ field: 'Region', sort: { direction: 'descending' } }, { field: 'Year', group: { kind: 'number', interval: 2, start: 2024 } }], columns: [], filters: [{ kind: 'condition', field: 'Amount', operator: 'greater-or-equal', value: 2 }], values: pivot.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const result = computePivotResult(workbook(), pivot);
    assert.deepEqual(result.rows.map((row) => row.label), ['West', 'East']);
    assert.equal(result.rows[0]?.children[0]?.label, '2024');
    assert.deepEqual(result.rows[0]?.children[0]?.sourceRowPaths, [{ sheetId: 'sheet-1', row: 3 }, { sheetId: 'sheet-1', row: 4 }]);
  });

  it('infers field types and executes top-items instead of treating it as a no-op', () => {
    const modelData = workbook(); const sheet = modelData.getSheet('sheet-1'); sheet.cells.set(0, 3, { value: 'Flag' }); sheet.cells.set(1, 3, { value: true }); sheet.cells.set(2, 3, { value: false }); sheet.cells.set(0, 4, { value: 'Date' }); sheet.cells.set(1, 4, { value: '2024-01-01' }); sheet.cells.set(2, 4, { value: '2024-02-01' });
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]); pivot.sourceRange.endColumn = 4; pivot.layout = { rows: [{ field: 'Region' }], columns: [], filters: [{ kind: 'top-items', field: 'Region', count: 1, valueField: 'Amount', direction: 'top' }], values: pivot.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const result = computePivotResult(modelData, pivot); assert.equal(result.fields.fields.find((field) => field.name === 'Region')?.dataType, 'text'); assert.equal(result.fields.fields.find((field) => field.name === 'Amount')?.dataType, 'number'); assert.equal(result.fields.fields.find((field) => field.name === 'Flag')?.dataType, 'boolean'); assert.equal(result.fields.fields.find((field) => field.name === 'Date')?.dataType, 'date'); assert.deepEqual(result.rows.map((row) => row.label), ['West']);
  });

  it('keeps column-specific source paths and calculates row/column/parent/running/rank/difference axes', () => {
    const base = model([{ field: 'Amount', summarizeBy: 'sum' }]); base.layout = { rows: [{ field: 'Region' }, { field: 'Year' }], columns: [], filters: [], values: base.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const parent = computePivotResult(workbook(), base); assert.equal(parent.rows.find((row) => row.label === 'East')?.kind, 'subtotal'); assert.equal(parent.rows.find((row) => row.label === 'East')?.children[0]?.kind, 'leaf');
    const column = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'column-percentage' } }]); column.layout = { rows: [{ field: 'Region' }], columns: [{ field: 'Year' }], filters: [], values: column.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const columnResult = computePivotResult(workbook(), column); const east = columnResult.rows.find((row) => row.label === 'East')!; assert.deepEqual(east.values[0]?.sourceRowPaths, [{ sheetId: 'sheet-1', row: 1 }]); assert.equal(east.values[0]?.values[0], 1 / 3);
    const running = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'running-total', axis: 'row' } }]); running.layout = { rows: [{ field: 'Region' }], columns: [{ field: 'Year' }], filters: [], values: running.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false }; const runningResult = computePivotResult(workbook(), running); const runningEast = runningResult.rows.find((row) => row.label === 'East')!; assert.deepEqual(runningEast.values.map((cell) => cell.values[0]), [1, 2]);
    const ranked = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'rank', axis: 'column', direction: 'descending' } }]); ranked.layout = { rows: [{ field: 'Region' }], columns: [{ field: 'Year' }], filters: [], values: ranked.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false }; const rankedResult = computePivotResult(workbook(), ranked); assert.deepEqual(rankedResult.rows.find((row) => row.label === 'East')?.values.map((cell) => cell.values[0]), [2, 1]);
    const parentAs = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'parent-percentage' } }]); parentAs.layout = base.layout; parentAs.layout.values = parentAs.valueFields; const parentResult = computePivotResult(workbook(), parentAs); assert.equal(parentResult.rows.find((row) => row.label === 'East')?.children[0]?.values[0]?.values[0], 1 / 3);
    const difference = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'difference', base: 'grand' } }]); difference.layout = { rows: [{ field: 'Region' }], columns: [], filters: [], values: difference.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false }; const differenceResult = computePivotResult(workbook(), difference); assert.equal(differenceResult.rows.find((row) => row.label === 'East')?.values[0]?.values[0], -6);
  });

  it('applies showAs percentages and rejects multi-worksheet execution explicitly', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'grand-percentage' } }]); pivot.layout = { rows: [{ field: 'Region' }], columns: [], filters: [], values: pivot.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const result = computePivotResult(workbook(), pivot); assert.equal(result.rows.find((row) => row.label === 'East')?.values[0]?.values[0], 3 / 9);
    pivot.dataSource = { kind: 'worksheet-ranges', ranges: [pivot.sourceRange, pivot.sourceRange], relationships: [{ id: 'r1', left: { sheetId: 'sheet-1', field: 'Region' }, right: { sheetId: 'sheet-1', field: 'Region' }, join: 'inner' }] }; assert.doesNotThrow(() => computePivotResult(workbook(), pivot));
  });

  it('writeback is a projection of the result tree and preserves non-sum aggregators', () => {
    const source = workbook(); const pivot = model([{ field: 'Amount', summarizeBy: 'product' }]); pivot.layout = { rows: [{ field: 'Region' }], columns: [], filters: [], values: pivot.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const writeback = buildPivotWriteback(pivot, source.getSheet('sheet-1')); assert.equal(writeback.values.length, 5); assert.equal(writeback.values[2]?.[1]?.value, 2); assert.equal(writeback.values[3]?.[1]?.value, 8);
  });

  it('applies slicer and timeline filters before aggregation', () => {
    const source = workbook();
    const sheet = source.getSheet('sheet-1');
    sheet.cells.set(0, 3, { value: 'Date' });
    sheet.cells.set(1, 3, { value: '2024-01-01' });
    sheet.cells.set(2, 3, { value: '2024-02-01' });
    sheet.cells.set(3, 3, { value: '2024-03-01' });
    sheet.cells.set(4, 3, { value: '2025-01-01' });
    sheet.cells.set(5, 3, { value: '2025-02-01' });
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.sourceRange.endColumn = 3;
    pivot.layout = { rows: [{ field: 'Region' }], columns: [], filters: [], values: pivot.valueFields, showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    pivot.slicers = [{ id: 's1', field: 'Region', selected: ['East'] }];
    pivot.timelines = [{ id: 't1', field: 'Date', start: '2024-01-01', end: '2024-12-31' }];
    const result = computePivotResult(source, pivot);
    assert.deepEqual(result.rows.map((row) => row.label), ['East']);
    assert.equal(result.grandTotal?.values[0], 3);
  });

  it('evaluates calculated fields and calculated items through the safe formula engine', () => {
    const source = workbook();
    const pivot = model([{ field: 'Adjusted', summarizeBy: 'sum' }]);
    pivot.layout = { rows: [{ field: 'Region' }], columns: [], filters: [], values: pivot.valueFields, calculatedFields: [{ name: 'Adjusted', formula: '=Amount*2' }], calculatedItems: [], showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    const result = computePivotResult(source, pivot);
    assert.equal(result.grandTotal?.values[0], 18);
    assert.equal(result.fields.fields.some((field) => field.name === 'Adjusted'), true);
    const itemPivot = model([{ field: 'East', summarizeBy: 'sum' }]);
    itemPivot.layout = { rows: [{ field: 'Region' }], columns: [], filters: [], values: itemPivot.valueFields, calculatedFields: [], calculatedItems: [{ field: 'Region', name: 'East', formula: '=Amount*3' }], showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false };
    assert.equal(computePivotResult(source, itemPivot).grandTotal?.values[0], 9);
  });
});
