import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  aggregatePivotValues,
  buildPivotGridProjection,
  computePivotResult,
  computePivotResultFromBlockSource,
  getPivotFieldCatalog,
  hitTestPivotProjection,
} from './engine';
import { buildPivotModel } from './helpers';

function workbookWithData(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-projection', 'Pivot Projection');
  const sheet = workbook.getSheet('sheet-1');
  [['Region', 'Amount'], ['East', 10], ['West', 20], ['East', 5]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  return workbook;
}

describe('native PivotGridProjection contract', () => {
  it('builds a complete canonical definition with stable field IDs', () => {
    const workbook = workbookWithData();
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    const pivot = buildPivotModel(workbook, 'sheet-1', 'canonical-pivot', sourceRange);
    assert.ok(pivot);
    assert.equal(pivot.schema, 'PivotDefinition');
    assert.deepEqual(pivot.source, { kind: 'worksheet-range', range: sourceRange });
    assert.equal(pivot.target.sheetId, 'sheet-1');
    assert.ok(pivot.fieldCatalog.fields.every((field) => field.fieldId.length > 0));
  });

  it('keeps typed members distinct and treats manual all as no filter', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-typed', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const region = getPivotFieldCatalog(workbook, pivot).fields.find((field) => field.name === 'Region')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ fieldId: getPivotFieldCatalog(workbook, pivot).fields.find((field) => field.name === 'Amount')!.fieldId, summarizeBy: 'count' }];
    pivot.layout.filters = [{ kind: 'manual', fieldId: region.fieldId, mode: 'all', memberKeys: [] }];
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 3);
    pivot.layout.filters = [{ kind: 'manual', fieldId: region.fieldId, mode: 'include', memberKeys: [{ type: 'text', value: 'East' }] }];
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 2);
    assert.notDeepEqual({ type: 'text', value: '1' }, { type: 'number', value: 1 });
    pivot.layout.filters = [{ kind: 'manual', scope: 'field', fieldId: region.fieldId, mode: 'include', memberKeys: [{ type: 'text', value: 'East' }] }];
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 8, column: 0 } };
    const fieldFiltered = buildPivotGridProjection(workbook, pivot);
    assert.equal(fieldFiltered.cells.some((cell) => cell.kind === 'filter'), false);
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 2);
  });

  it('keeps a root data row for Columns plus Values when Rows is empty', () => {
    const workbook = new WorkbookModel('pivot-columns-only', 'Pivot Columns Only');
    const sheet = workbook.getSheet('sheet-1');
    [['Month', 'Sales'], ['Jan', 10], ['Feb', 20], ['Mar', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-columns-only', sourceRange);
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const month = catalog.fields.find((field) => field.name === 'Month')!;
    const sales = catalog.fields.find((field) => field.name === 'Sales')!;
    pivot.layout.rows = [];
    pivot.layout.columns = [{ fieldId: month.fieldId }];
    pivot.layout.values = [{ fieldId: sales.fieldId, summarizeBy: 'sum' }];
    pivot.layout.showGrandTotals = true;
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 6, column: 0 } };

    const tree = computePivotResult(workbook, pivot);
    assert.equal(tree.rows.length, 1);
    assert.equal(tree.rows[0]?.nodeId, '__root__');
    assert.deepEqual(Object.fromEntries(tree.columnPaths.map((path, index) => [path[0], tree.rows[0]?.values[index]?.values[0]])), { Jan: 10, Feb: 20, Mar: 30 });
    assert.equal(tree.grandTotal?.values[0], 60);

    const projection = buildPivotGridProjection(workbook, pivot, tree);
    const values = projection.cells.filter((cell) => cell.kind === 'value');
    assert.deepEqual(Object.fromEntries(values.map((cell) => [cell.columnPath?.[0], cell.value])), { Jan: 10, Feb: 20, Mar: 30 });
    assert.equal(projection.cells.some((cell) => cell.text === 'Jan Sales'), true);
    assert.equal(projection.cells.some((cell) => cell.text === 'Grand Total'), true);
  });

  it('implements each aggregate independently', () => {
    const rows = [{ values: { value: 2 } }, { values: { value: 4 } }, { values: { value: 4 } }, { values: { value: null } }];
    assert.equal(aggregatePivotValues(rows, 'value', 'sum'), 10);
    assert.equal(aggregatePivotValues(rows, 'value', 'count'), 3);
    assert.equal(aggregatePivotValues(rows, 'value', 'count-numbers'), 3);
    assert.equal(aggregatePivotValues(rows, 'value', 'average'), 10 / 3);
    assert.equal(aggregatePivotValues(rows, 'value', 'min'), 2);
    assert.equal(aggregatePivotValues(rows, 'value', 'max'), 4);
    assert.equal(aggregatePivotValues(rows, 'value', 'product'), 32);
    assert.equal(aggregatePivotValues(rows, 'value', 'distinct-count'), 2);
    assert.equal(aggregatePivotValues(rows, 'value', 'varp'), 8 / 9);
  });

  it('returns a derived overlay, reports collisions, and supports hit testing without cell writeback', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-overlay', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 5, column: 0 } };
    const before = workbook.getSheet('sheet-1').cells.count();
    const first = buildPivotGridProjection(workbook, pivot);
    assert.equal(first.collision.status, 'clear');
    workbook.getSheet('sheet-1').cells.set(5, 0, { value: 'ordinary cell' });
    const projection = buildPivotGridProjection(workbook, pivot);
    assert.equal(workbook.getSheet('sheet-1').cells.count(), before + 1);
    assert.equal(projection.collision.status, 'collision');
    assert.deepEqual(projection.cells, first.cells);
    assert.deepEqual(projection.occupiedRange, first.occupiedRange);
    assert.equal(projection.schema, 'PivotGridProjection');
    const hit = hitTestPivotProjection(projection, 0, 0);
    assert.equal(hit.pivotId, pivot.id);
    assert.equal(hit.kind, 'cell');
  });

  it('rejects a stale layout result, localizes field captions, and preserves source-stale manual results', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-revision', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 8, column: 0 } };
    pivot.refreshPolicy.mode = 'manual';
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const firstResult = computePivotResult(workbook, pivot);
    const firstProjection = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.equal(firstProjection.cells.some((cell) => cell.text === 'Amount'), true);

    workbook.getSheet('sheet-1').cells.set(1, 1, { value: 100 });
    const stale = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.equal(stale.refresh.status, 'stale');
    assert.equal(stale.cells.some((cell) => cell.value === 35), true);

    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'count' }];
    const refreshedLayout = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.notEqual(refreshedLayout.refresh.status, 'stale');
    assert.equal(refreshedLayout.cells.some((cell) => cell.value === 3), true);
  });

  it('retains a cached block Pivot result across loading and source failure states', () => {
    const workbook = new WorkbookModel('pivot-block', 'Pivot Block');
    workbook.addDataSource({
      schema: 'DataSourceManifest',
      version: 1,
      id: 'source-block',
      name: 'Block Source',
      kind: 'chunked-table',
      sourceSheetId: 'sheet-1',
      rowCount: 2,
      fields: [
        { id: 'region', name: 'Region', ordinal: 0, type: 'text' },
        { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' },
      ],
      blockRowCount: 65_536,
      blocks: [],
      revision: 1,
    });
    const pivot = {
      schema: 'PivotDefinition' as const,
      id: 'pivot-block',
      source: { kind: 'data-source' as const, dataSourceId: 'source-block' },
      target: { sheetId: 'sheet-1', anchor: { row: 5, column: 0 } },
      fieldCatalog: {
        schema: 'PivotFieldCatalog' as const,
        fields: [
          { fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0 },
          { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 },
        ],
      },
      refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
      layout: {
        rows: [{ fieldId: 'region' }], columns: [], filters: [],
        values: [{ fieldId: 'amount', summarizeBy: 'sum' as const }],
        showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false,
        expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
      },
    };
    const loading = buildPivotGridProjection(workbook, pivot);
    assert.equal(loading.refresh.status, 'refreshing');
    assert.equal(loading.cells.some((cell) => cell.kind === 'loading'), true);
    const result = computePivotResultFromBlockSource(workbook, pivot, {
      fields: [
        { fieldId: 'region', name: 'Region', ordinal: 0, dataType: 'text' },
        { fieldId: 'amount', name: 'Amount', ordinal: 1, dataType: 'number' },
      ],
      rows: [
        { values: { region: 'East', amount: 10 }, paths: [{ sheetId: 'sheet-1', row: 1 }] },
        { values: { region: 'West', amount: 20 }, paths: [{ sheetId: 'sheet-1', row: 2 }] },
      ],
    }, 'source-block:1');
    const ready = buildPivotGridProjection(workbook, pivot, result, { sourceState: { availability: 'ready' } });
    assert.equal(ready.refresh.status, 'ready');
    assert.equal(ready.cells.some((cell) => cell.value === 30), true);
    const failed = buildPivotGridProjection(workbook, pivot, undefined, { sourceState: { availability: 'error', error: 'offline' } });
    assert.equal(failed.refresh.status, 'error');
    assert.equal(failed.cells.some((cell) => cell.value === 30), true);
  });
});
