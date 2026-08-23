import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel, type PivotModel, type PivotValueField } from '@react-sheets/core-model';
import { computePivotResult, computePivotTable, getPivotRevisionKey } from './pivot-engine';
import { buildPivotWriteback } from './pivot-write';

function workbook(): WorkbookModel {
  const model = new WorkbookModel('pivot-engine-test', 'Pivot Engine');
  const sheet = model.getSheet('sheet-1');
  const rows = [
    ['Region', 'Owner', 'Year', 'Amount'],
    ['East', 'Maya', '2024-01-10', 1],
    ['East', 'Noah', '2024-02-10', 2],
    ['West', 'Maya', '2024-03-10', 3],
    ['West', 'Noah', '2025-01-10', 4],
    ['South', 'Maya', '2025-02-10', 5],
  ];
  rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  return model;
}

function model(values: PivotValueField[]): PivotModel {
  return {
    id: 'p1',
    sheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 },
    layout: {
      rows: [],
      columns: [],
      filters: [],
      values,
      showSubtotals: true,
      showGrandTotals: true,
      compact: false,
      repeatLabels: false,
      calculatedFields: [],
      calculatedItems: [],
    },
  };
}

describe('pivot domain engine', () => {
  it('round-trips the canonical definition through workbook snapshots', () => {
    const source = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    source.layout.rows = [{ field: 'Region' }];
    source.refreshPolicy = { mode: 'on-change', preserveFormatting: true, refreshOnLoad: false };
    const workbookModel = workbook();
    workbookModel.getSheet('sheet-1').pivots.push(source);
    const restored = WorkbookModel.fromSnapshot(workbookModel.snapshot()).getSheet('sheet-1').pivots[0];
    assert.deepEqual(restored?.layout, source.layout);
    assert.equal('rowFields' in (restored ?? {}), false);
  });

  it('calculates every supported aggregate independently', () => {
    const supported = ['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'] as const;
    const source = workbook();
    for (const summarizeBy of supported) {
      const pivot = model([{ field: 'Amount', summarizeBy }]);
      pivot.layout.rows = [{ field: 'Region' }];
      const result = computePivotResult(source, pivot);
      assert.equal(result.rows.length, 3);
      assert.equal(result.grandTotal?.values.length, 1);
    }
  });

  it('builds nested rows, columns, grouping, filtering and sorting', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.layout.rows = [
      { field: 'Region', sort: { direction: 'descending' } },
      { field: 'Year', group: { kind: 'date', unit: 'year' } },
    ];
    pivot.layout.columns = [{ field: 'Owner' }];
    pivot.layout.filters = [{ kind: 'condition', field: 'Amount', operator: 'greater-or-equal', value: 2 }];
    const result = computePivotResult(workbook(), pivot);
    assert.equal(result.columnPaths.length, 2);
    assert.equal(result.rows[0]?.label, 'West');
    assert.ok(result.rows[0]?.children.length);
  });

  it('applies showAs values and preserves source paths', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum', showAs: { kind: 'grand-percentage' } }]);
    pivot.layout.rows = [{ field: 'Region' }];
    const result = computePivotResult(workbook(), pivot);
    assert.equal(result.sourceRowPaths.length, 5);
    assert.equal(result.rows[0]?.values[0]?.values[0], 3 / 15);
  });

  it('evaluates calculated fields and calculated items through the safe formula engine', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.layout.rows = [{ field: 'Region' }];
    pivot.layout.calculatedFields = [{ name: 'Adjusted', formula: '=Amount*2' }];
    const result = computePivotResult(workbook(), pivot);
    assert.equal(result.fields.fields.some((field) => field.name === 'Adjusted'), true);

    pivot.layout.calculatedItems = [{ field: 'Region', name: 'East', formula: '=Amount*3' }];
    assert.doesNotThrow(() => computePivotResult(workbook(), pivot));
  });

  it('applies slicer and timeline filters before aggregation', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.layout.rows = [{ field: 'Region' }];
    pivot.slicers = [{ id: 's1', field: 'Owner', selected: ['Maya'] }];
    pivot.timelines = [{ id: 't1', field: 'Year', start: '2024-01-01', end: '2024-12-31' }];
    const result = computePivotResult(workbook(), pivot);
    assert.equal(result.grandTotal?.values[0], 4);
  });

  it('exports a result tree without maintaining an independent aggregation path', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'product' }]);
    pivot.layout.rows = [{ field: 'Region' }];
    const result = buildPivotWriteback(pivot, workbook().getSheet('sheet-1'));
    assert.equal(result.values.at(-1)?.[0]?.value, 'Grand Total');
    assert.equal(result.values.at(-1)?.[1]?.value, 120);
  });

  it('returns a table projection from the canonical layout', () => {
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.layout.rows = [{ field: 'Region' }];
    const result = computePivotTable(workbook(), pivot);
    assert.deepEqual(result.headers, ['Region', 'SUM of Amount']);
    assert.equal(result.rows.length, 3);
  });

  it('invalidates the derived result when source, layout or filters change', () => {
    const source = workbook();
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.layout.rows = [{ field: 'Region' }];
    const first = computePivotResult(source, pivot);
    const firstKey = getPivotRevisionKey(source, pivot);
    const second = computePivotResult(source, pivot);
    assert.deepEqual(second, first);
    assert.notEqual(second, first, 'callers receive a clone, never the mutable cache entry');

    source.getSheet('sheet-1').cells.set(1, 3, { value: 10 });
    const afterSource = computePivotResult(source, pivot);
    assert.equal(afterSource.grandTotal?.values[0], 24);
    assert.notEqual(getPivotRevisionKey(source, pivot).sourceRevision, firstKey.sourceRevision);

    pivot.layout.values[0] = { field: 'Amount', summarizeBy: 'count' };
    const afterLayout = computePivotResult(source, pivot);
    assert.equal(afterLayout.grandTotal?.values[0], 5);
    assert.notEqual(getPivotRevisionKey(source, pivot).layoutRevision, firstKey.layoutRevision);

    pivot.layout.filters = [{ kind: 'manual', field: 'Region', selected: ['East'] }];
    const afterFilter = computePivotResult(source, pivot);
    assert.equal(afterFilter.grandTotal?.values[0], 2);
    assert.notEqual(getPivotRevisionKey(source, pivot).filterRevision, firstKey.filterRevision);
  });

  it('reads a source range from another sheet and applies linked slicers there', () => {
    const source = workbook();
    const sourceSheet = source.addSheet('source-2', 'Source 2');
    const sourceRows = [
      ['Region', 'Amount'],
      ['East', 7],
      ['West', 11],
    ];
    sourceRows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => sourceSheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    pivot.sheetId = 'sheet-1';
    pivot.sourceRange = { sheetId: 'source-2', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
    pivot.layout.rows = [{ field: 'Region' }];
    const linked = model([{ field: 'Amount', summarizeBy: 'sum' }]);
    linked.id = 'linked';
    linked.sheetId = 'source-2';
    linked.sourceRange = { sheetId: 'source-2', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
    linked.slicers = [{ id: 'slicer-region', field: 'Region', selected: ['West'], connectedPivotIds: [pivot.id] }];
    source.getSheet('sheet-1').pivots.push(pivot);
    sourceSheet.pivots.push(linked);
    const result = computePivotResult(source, pivot);
    assert.equal(result.grandTotal?.values[0], 11);
    assert.equal(result.rows[0]?.label, 'West');
  });
});
