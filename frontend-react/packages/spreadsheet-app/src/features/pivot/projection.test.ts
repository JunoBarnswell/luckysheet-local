import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  aggregatePivotValues,
  buildPivotGridProjection,
  computePivotResult,
  getPivotFieldCatalog,
  hitTestPivotProjection,
  migratePivotDefinition,
} from './engine';
import { buildPivotModel } from './helpers';

function workbookWithData(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-projection', 'Pivot Projection');
  const sheet = workbook.getSheet('sheet-1');
  [['Region', 'Amount'], ['East', 10], ['West', 20], ['East', 5]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  return workbook;
}

describe('native PivotGridProjection contract', () => {
  it('migrates one legacy definition into a complete canonical definition', () => {
    const workbook = workbookWithData();
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    const migrated = migratePivotDefinition(workbook, {
      id: 'legacy-pivot',
      sheetId: 'sheet-1',
      sourceRange,
      layout: {
        rows: [{ field: 'Region' }], columns: [], filters: [], values: [{ field: 'Amount', summarizeBy: 'sum' }],
        showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false,
      },
    });
    assert.equal(migrated.schema, 'PivotDefinition');
    assert.deepEqual(migrated.source, { kind: 'worksheet-range', range: sourceRange });
    assert.equal(migrated.target.sheetId, 'sheet-1');
    assert.equal(migrated.layout.rows[0]?.fieldId, migrated.fieldCatalog.fields.find((field) => field.name === 'Region')?.fieldId);
    assert.equal(migrated.layout.values[0]?.fieldId, migrated.fieldCatalog.fields.find((field) => field.name === 'Amount')?.fieldId);
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
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 0, column: 0 } };
    const before = workbook.getSheet('sheet-1').cells.count();
    const projection = buildPivotGridProjection(workbook, pivot);
    assert.equal(workbook.getSheet('sheet-1').cells.count(), before);
    assert.equal(projection.collision.status, 'collision');
    assert.equal(projection.schema, 'PivotGridProjection');
    const hit = hitTestPivotProjection(projection, projection.target.anchor.row, projection.target.anchor.column);
    assert.equal(hit.pivotId, pivot.id);
    assert.equal(hit.kind, 'cell');
  });
});
