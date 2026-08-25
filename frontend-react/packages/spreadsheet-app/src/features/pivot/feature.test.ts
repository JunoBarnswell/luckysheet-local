import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerPivotFeature } from './index';
import { buildPivotModel, connectedPivotIdsForSource } from './helpers';
import { computePivotResult, getPivotFieldCatalog, getPivotRevisionKey } from './engine';
import { buildPivotWriteback } from './writeback';

function seedCrossSheetWorkbook(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-feature-test', 'Pivot Feature');
  const source = workbook.addSheet('source-2', 'Source 2');
  [['Region', 'Amount'], ['East', 10], ['West', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => source.cells.set(rowIndex, columnIndex, { value })));
  return workbook;
}

function pivotDefinition(): ReturnType<typeof buildPivotModel> {
  const workbook = seedCrossSheetWorkbook();
  return buildPivotModel(workbook, 'sheet-1', 'pivot-1', { sheetId: 'source-2', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
}

function sameSheetRelationalPivot(): { workbook: WorkbookModel; pivot: NonNullable<ReturnType<typeof buildPivotModel>> } {
  const workbook = new WorkbookModel('pivot-drilldown-provenance', 'Pivot Drilldown Provenance');
  const sheet = workbook.getSheet('sheet-1');
  [['CustomerId', 'Amount'], ['c1', 100], ['c2', 200]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  // Deliberately reverse customer row order so the joined record uses row 1
  // from Orders and row 2 from Customers on the same worksheet.
  [['CustomerId', 'Region'], ['c2', 'West'], ['c1', 'East']].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, 4 + columnIndex, { value })));
  const ordersRange = { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
  const customersRange = { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 4, endColumn: 5 };
  const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-drilldown-provenance', ordersRange)!;
  pivot.target = { sheetId: 'sheet-1', anchor: { row: 8, column: 0 } };
  pivot.source = {
    kind: 'worksheet-ranges',
    ranges: [
      { sourceId: 'orders', range: ordersRange },
      { sourceId: 'customers', range: customersRange },
    ],
    relationships: [{
      id: 'orders-customers',
      left: { sourceId: 'orders', fieldId: 'source:orders:column:0' },
      right: { sourceId: 'customers', fieldId: 'source:customers:column:0' },
      join: 'left',
    }],
  };
  pivot.fieldCatalog = getPivotFieldCatalog(workbook, pivot);
  const region = pivot.fieldCatalog.fields.find((field) => field.fieldId === 'source:customers:column:1')!;
  const amount = pivot.fieldCatalog.fields.find((field) => field.fieldId === 'source:orders:column:1')!;
  pivot.layout.rows = [{ fieldId: region.fieldId }];
  pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
  return { workbook, pivot };
}

describe('pivot feature contract', () => {
  it('creates a new destination worksheet and PivotTable in one reversible transaction', () => {
    const workbook = seedCrossSheetWorkbook();
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerPivotFeature(runtime);
    const pivot = pivotDefinition();
    assert.ok(pivot);
    pivot.id = 'pivot-create';
    pivot.target = { sheetId: 'pivot-sheet', anchor: { row: 0, column: 0 } };

    const result = runtime.execute('pivot.create', {
      pivot,
      destination: { kind: 'new-sheet', sheetId: 'pivot-sheet', name: 'Pivot Output' },
    });
    assert.equal(result.mutationCount, 2);
    assert.equal(runtime.getUndoEntries().length, 1);
    assert.equal(workbook.getSheet('pivot-sheet').pivots[0]?.id, 'pivot-create');
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.sheets.has('pivot-sheet'), false);
    assert.equal(runtime.redo(), true);
    assert.equal(workbook.getSheet('pivot-sheet').pivots[0]?.id, 'pivot-create');
    assert.equal(runtime.getUndoEntries().length, 1);
  });

  it('rejects invalid create plans before the destination worksheet mutation', () => {
    const workbook = seedCrossSheetWorkbook();
    const source = workbook.getSheet('source-2');
    source.cells.set(0, 1, { value: 'Region' });
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerPivotFeature(runtime);
    const pivot = pivotDefinition();
    assert.ok(pivot);
    pivot.id = 'pivot-invalid-create';
    pivot.target = { sheetId: 'pivot-invalid-sheet', anchor: { row: 0, column: 0 } };
    const before = workbook.snapshot();
    assert.throws(() => runtime.execute('pivot.create', {
      pivot,
      destination: { kind: 'new-sheet', sheetId: 'pivot-invalid-sheet', name: 'Pivot Invalid' },
    }), /duplicated/);
    assert.deepEqual(workbook.snapshot(), before);
    assert.equal(runtime.getUndoEntries().length, 0);
  });

  it('rolls back the destination worksheet when a later mutation listener fails', () => {
    const workbook = seedCrossSheetWorkbook();
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerPivotFeature(runtime);
    const pivot = pivotDefinition();
    assert.ok(pivot);
    pivot.id = 'pivot-injected-failure';
    pivot.target = { sheetId: 'pivot-failure-sheet', anchor: { row: 0, column: 0 } };
    runtime.onMutation((mutation, source) => {
      if (source === 'command' && mutation.id === 'pivot.add') throw new Error('Injected pivot create failure');
    });
    const before = workbook.snapshot();
    assert.throws(() => runtime.execute('pivot.create', {
      pivot,
      destination: { kind: 'new-sheet', sheetId: 'pivot-failure-sheet', name: 'Pivot Failure' },
    }), /Injected pivot create failure/);
    assert.deepEqual(workbook.snapshot(), before);
    assert.equal(runtime.getUndoEntries().length, 0);
  });

  it('keeps display sheet and cross-sheet source distinct', () => {
    const workbook = seedCrossSheetWorkbook();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-1', { sheetId: 'source-2', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    assert.equal(pivot.target.sheetId, 'sheet-1');
    assert.equal(pivot.source.kind, 'worksheet-range');
    const sourceRange = pivot.source.kind === 'worksheet-range' ? pivot.source.range : undefined;
    assert.equal(sourceRange?.sheetId, 'source-2');
    workbook.getSheet('sheet-1').pivots.push(pivot);
    assert.deepEqual(connectedPivotIdsForSource(workbook, 'sheet-1', sourceRange!), ['pivot-1']);
  });

  it('drill-down creates a pure detail sheet and removes it through undo', () => {
    const workbook = seedCrossSheetWorkbook();
    const pivot = pivotDefinition();
    assert.ok(pivot);
    const runtime = new CommandRuntime(workbook);
    registerPivotFeature(runtime);
    runtime.execute('pivot.add', pivot);
    runtime.execute('pivot.drillDown', {
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      label: 'East',
      sourceRowPaths: [{ sheetId: 'source-2', row: 1 }],
      targetSheetId: 'drill-1',
      target: { row: 0, column: 0 },
    });
    assert.equal(workbook.getSheet('drill-1').cells.get(0, 0)?.value, 'Region');
    assert.equal(workbook.getSheet('drill-1').cells.get(1, 0)?.value, 'East');
    assert.equal(workbook.getSheet('drill-1').cells.get(1, 1)?.value, 10);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.sheets.has('drill-1'), false);
    assert.equal(runtime.redo(), true);
    assert.equal(workbook.getSheet('drill-1').cells.get(1, 1)?.value, 10);
  });

  it('drill-down resolves same-sheet joined rows by sourceId and recordId', () => {
    const { workbook, pivot } = sameSheetRelationalPivot();
    const runtime = new CommandRuntime(workbook);
    registerPivotFeature(runtime);
    runtime.execute('pivot.add', pivot);
    const result = computePivotResult(workbook, workbook.getSheet('sheet-1').pivots[0]!);
    const east = result.rows.find((node) => node.label === 'East')?.values[0];
    assert.ok(east);
    assert.deepEqual(east.sourceRowPaths.map((path) => [path.sourceId, path.row, path.recordId]), [
      ['orders', 1, 'orders:1'],
      ['customers', 2, 'orders:1'],
    ]);
    runtime.execute('pivot.drillDown', {
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      label: 'East',
      sourceRowPaths: east.sourceRowPaths,
      targetSheetId: 'drill-same-sheet',
      target: { row: 0, column: 0 },
    });
    const detail = workbook.getSheet('drill-same-sheet');
    assert.deepEqual([0, 1, 2, 3].map((column) => detail.cells.get(1, column)?.value), ['c1', 100, 'c1', 'East']);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.sheets.has('drill-same-sheet'), false);
    runtime.execute('pivot.drillDown', {
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      label: 'left-unmatched',
      sourceRowPaths: [{ sourceId: 'orders', recordId: 'orders:2', sheetId: 'sheet-1', row: 2 }],
      targetSheetId: 'drill-left-unmatched',
      target: { row: 0, column: 0 },
    });
    const unmatched = workbook.getSheet('drill-left-unmatched');
    assert.deepEqual([0, 1, 2, 3].map((column) => unmatched.cells.get(1, column)?.value), ['c2', 200, null, null]);
  });

  it('rejects incomplete inner-join provenance before creating a detail sheet', () => {
    const { workbook, pivot } = sameSheetRelationalPivot();
    assert.ok(pivot.source.kind === 'worksheet-ranges');
    pivot.source.relationships[0]!.join = 'inner';
    const runtime = new CommandRuntime(workbook);
    registerPivotFeature(runtime);
    runtime.execute('pivot.add', pivot);
    assert.throws(() => runtime.execute('pivot.drillDown', {
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      label: 'incomplete',
      sourceRowPaths: [{ sourceId: 'orders', recordId: 'orders:1', sheetId: 'sheet-1', row: 1 }],
      targetSheetId: 'drill-incomplete',
      target: { row: 0, column: 0 },
    }), /provenance is incomplete/);
    assert.equal(workbook.sheets.has('drill-incomplete'), false);
  });

  it('writes a cross-sheet pivot result from the complete workbook model', () => {
    const workbook = seedCrossSheetWorkbook();
    const pivot = pivotDefinition();
    assert.ok(pivot);
    const result = buildPivotWriteback(pivot, workbook);
    assert.equal(result.values.at(-1)?.[1]?.value, 30);
  });

  it('rejects unknown fields once a source header exists', () => {
    const workbook = seedCrossSheetWorkbook();
    const runtime = new CommandRuntime(workbook);
    registerPivotFeature(runtime);
    const pivot = pivotDefinition();
    assert.ok(pivot);
    pivot.layout.rows = [{ fieldId: 'Missing' }];
    assert.throws(() => runtime.execute('pivot.add', pivot), /Unknown pivot field: Missing/);
    assert.equal(workbook.getSheet('sheet-1').pivots.length, 0);
  });

  it('invalidates pure derived results by source, layout and filter revisions', () => {
    const workbook = seedCrossSheetWorkbook();
    const pivot = pivotDefinition();
    assert.ok(pivot);
    pivot.layout.rows = [{ fieldId: pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!.fieldId }];
    const first = computePivotResult(workbook, pivot);
    assert.equal(first.schema, 'PivotResultTree');
    const firstKey = getPivotRevisionKey(workbook, pivot);
    const second = computePivotResult(workbook, pivot);
    assert.deepEqual(second, first);
    assert.notEqual(second, first);
    workbook.getSheet('source-2').cells.set(1, 1, { value: 15 });
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 35);
    // Source revision is supplied by the block/data-source revision counter in
    // production; the sparse legacy CellMatrix has no mutation counter.
    pivot.layout.values[0] = { fieldId: pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!.fieldId, summarizeBy: 'count' };
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 2);
    assert.notEqual(getPivotRevisionKey(workbook, pivot).layoutRevision, firstKey.layoutRevision);
    pivot.layout.filters = [{ kind: 'manual', fieldId: pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!.fieldId, mode: 'include', memberKeys: [{ type: 'text', value: 'East' }] }];
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 1);
    assert.notEqual(getPivotRevisionKey(workbook, pivot).filterRevision, firstKey.filterRevision);
  });
});
