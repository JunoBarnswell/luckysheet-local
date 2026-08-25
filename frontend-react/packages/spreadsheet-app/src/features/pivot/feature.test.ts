import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerPivotFeature } from './index';
import { buildPivotModel, connectedPivotIdsForSource } from './helpers';
import { computePivotResult, getPivotRevisionKey } from './engine';
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
