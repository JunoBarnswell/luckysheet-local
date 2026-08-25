import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerDrawingFeature } from '../drawing';
import { registerPivotFeature } from './index';
import { buildPivotModel, connectedPivotIdsForSource } from './helpers';
import { buildPivotSlicerDrawing, buildPivotTimelineDrawing } from '../pivot-controls';
import { computePivotResult, getPivotFieldCatalog, getPivotRevisionKey } from './engine';
import { buildPivotWriteback } from './writeback';
import { clearPivotFilterFamily, clearPivotFiltersForField, setPivotColumnGrandTotals, setPivotRowGrandTotals, upsertPivotFilter } from './panel-state';

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

function pivotChartDrawing(pivotId: string) {
  return {
    drawing: {
      id: `chart-${pivotId}`,
      sheetId: 'sheet-1',
      kind: 'chart' as const,
      payloadId: `chart-payload-${pivotId}`,
      anchor: { kind: 'absolute' as const },
      transform: { x: 10, y: 10, width: 240, height: 160, rotation: 0 },
      zIndex: 1,
    },
    payload: {
      kind: 'chart' as const,
      chartId: `chart-payload-${pivotId}`,
      pivotId,
      sourceRanges: [],
      chartType: 'column' as const,
      elements: {
        hiddenData: 'show' as const,
        legend: { visible: true, position: 'bottom' as const },
        dataLabels: { visible: false },
        categoryAxis: { id: 'category', position: 'bottom' as const },
        valueAxis: { id: 'value', position: 'left' as const },
      },
    },
  };
}

function snapshotWithStableCollectionOrder(workbook: WorkbookModel) {
  const snapshot = workbook.snapshot();
  for (const sheet of snapshot.sheets) {
    const byId = <T extends { id: string }>(left: T, right: T) => left.id < right.id ? -1 : left.id > right.id ? 1 : 0;
    sheet.pivots.sort(byId);
    sheet.drawings.sort(byId);
  }
  return snapshot;
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
    const before = snapshotWithStableCollectionOrder(workbook);
    assert.throws(() => runtime.execute('pivot.create', {
      pivot,
      destination: { kind: 'new-sheet', sheetId: 'pivot-invalid-sheet', name: 'Pivot Invalid' },
    }), /duplicated/);
    assert.deepEqual(snapshotWithStableCollectionOrder(workbook), before);
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
    const before = snapshotWithStableCollectionOrder(workbook);
    assert.throws(() => runtime.execute('pivot.create', {
      pivot,
      destination: { kind: 'new-sheet', sheetId: 'pivot-failure-sheet', name: 'Pivot Failure' },
    }), /Injected pivot create failure/);
    assert.deepEqual(snapshotWithStableCollectionOrder(workbook), before);
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

  it('removes PivotChart and Pivot controls as one atomic reversible lifecycle transaction', () => {
    const workbook = seedCrossSheetWorkbook();
    const pivot = pivotDefinition();
    assert.ok(pivot);
    pivot.id = 'pivot-delete-target';
    const connectedPivot = structuredClone(pivot);
    connectedPivot.id = 'pivot-delete-connected';
    connectedPivot.target = { sheetId: 'sheet-1', anchor: { row: 20, column: 0 } };
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    registerPivotFeature(runtime);
    runtime.execute('pivot.add', pivot);
    runtime.execute('pivot.add', connectedPivot);

    const chart = pivotChartDrawing(pivot.id);
    const slicer = buildPivotSlicerDrawing({
      drawingId: 'slicer-primary',
      payloadId: 'slicer-primary-payload',
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      fieldId: 'source:source-2:column:0',
      transform: { x: 10, y: 190, width: 160, height: 80, rotation: 0 },
      zIndex: 2,
    });
    const timeline = buildPivotTimelineDrawing({
      drawingId: 'timeline-primary',
      payloadId: 'timeline-primary-payload',
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      fieldId: 'source:source-2:column:0',
      transform: { x: 10, y: 280, width: 200, height: 80, rotation: 0 },
      zIndex: 3,
    });
    const connected = buildPivotSlicerDrawing({
      drawingId: 'slicer-connected',
      payloadId: 'slicer-connected-payload',
      sheetId: 'sheet-1',
      pivotId: connectedPivot.id,
      connectedPivotIds: [pivot.id, connectedPivot.id],
      fieldId: 'source:source-2:column:0',
      transform: { x: 10, y: 370, width: 160, height: 80, rotation: 0 },
      zIndex: 4,
    });
    for (const entry of [chart, slicer, timeline, connected]) {
      runtime.execute('drawing.add', { sheetId: 'sheet-1', drawing: entry.drawing, payload: entry.payload });
    }

    const before = snapshotWithStableCollectionOrder(workbook);
    assert.throws(() => runtime.applyRemoteMutations([{
      id: 'pivot.remove',
      unitId: workbook.unitId,
      sheetId: 'sheet-1',
      params: pivot.id,
      affectedRanges: [],
    }]), /dependent drawings/);
    assert.deepEqual(snapshotWithStableCollectionOrder(workbook), before);
    assert.throws(() => runtime.applyRemoteMutations([
      {
        id: 'drawing.remove',
        unitId: workbook.unitId,
        sheetId: 'sheet-1',
        params: { sheetId: 'sheet-1', drawingId: chart.drawing.id },
        affectedRanges: [],
      },
      {
        id: 'pivot.remove',
        unitId: workbook.unitId,
        sheetId: 'sheet-1',
        params: pivot.id,
        affectedRanges: [],
      },
    ]), /dependent drawings/);
    assert.deepEqual(snapshotWithStableCollectionOrder(workbook), before);

    const result = runtime.execute('pivot.remove', { sheetId: 'sheet-1', pivotId: pivot.id });
    assert.equal(result.mutationCount, 5);
    const sheet = workbook.getSheet('sheet-1');
    assert.equal(sheet.pivots.some((entry) => entry.id === pivot.id), false);
    assert.equal(sheet.drawings.some((entry) => entry.id === chart.drawing.id), false);
    assert.equal(sheet.drawings.some((entry) => entry.id === slicer.drawing.id), false);
    assert.equal(sheet.drawings.some((entry) => entry.id === timeline.drawing.id), false);
    assert.deepEqual((sheet.drawingPayloads.get(connected.drawing.payloadId) as { connectedPivotIds?: string[] })?.connectedPivotIds, [connectedPivot.id]);

    assert.equal(runtime.undo(), true);
    assert.deepEqual(snapshotWithStableCollectionOrder(workbook), before);
    assert.equal(runtime.redo(), true);
    assert.equal(workbook.getSheet('sheet-1').pivots.some((entry) => entry.id === pivot.id), false);
    assert.equal(workbook.getSheet('sheet-1').drawingPayloads.has(chart.drawing.payloadId), false);
    assert.deepEqual((workbook.getSheet('sheet-1').drawingPayloads.get(connected.drawing.payloadId) as { connectedPivotIds?: string[] })?.connectedPivotIds, [connectedPivot.id]);
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

  it('keeps Pivot panel row and column grand-total controls independent and typed', () => {
    const pivot = pivotDefinition();
    assert.ok(pivot);
    const rowOnly = setPivotRowGrandTotals(pivot.layout, false);
    assert.equal(rowOnly.showRowGrandTotals, false);
    assert.equal(rowOnly.showColumnGrandTotals, true);
    const columnOnly = setPivotColumnGrandTotals(rowOnly, false);
    assert.equal(columnOnly.showRowGrandTotals, false);
    assert.equal(columnOnly.showColumnGrandTotals, false);
    assert.throws(() => setPivotRowGrandTotals(pivot.layout, 'false' as never), /row grand-total state is invalid/);
    assert.throws(() => setPivotColumnGrandTotals(pivot.layout, 1 as never), /column grand-total state is invalid/);
  });

  it('includes every instant on a timeline end date and excludes the next day', () => {
    const workbook = new WorkbookModel('pivot-timeline-boundary', 'Pivot Timeline Boundary');
    const sheet = workbook.getSheet('sheet-1');
    [
      ['Date', 'Amount'],
      ['2026-08-25T00:00:00', 10],
      ['2026-08-25T12:00:00', 20],
      ['2026-08-25T23:59:59', 30],
      ['2026-08-26T00:00:00', 40],
    ].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-timeline-boundary', {
      sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 1,
    });
    assert.ok(pivot);
    const dateField = pivot.fieldCatalog.fields.find((field) => field.name === 'Date');
    const amountField = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount');
    assert.ok(dateField);
    assert.ok(amountField);
    pivot.layout.rows = [];
    pivot.layout.columns = [];
    pivot.layout.values = [{ fieldId: amountField.fieldId, summarizeBy: 'sum' }];
    sheet.pivots.push(pivot);
    const timeline = buildPivotTimelineDrawing({
      drawingId: 'timeline-boundary',
      payloadId: 'timeline-boundary-payload',
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      fieldId: dateField.fieldId,
      period: { start: '2026-08-25', end: '2026-08-25' },
      transform: { x: 10, y: 20, width: 300, height: 80 },
      zIndex: 1,
    });
    sheet.drawings.push(timeline.drawing);
    sheet.drawingPayloads.set(timeline.drawing.payloadId, timeline.payload);
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 60);
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
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', fieldId: pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!.fieldId, mode: 'include', memberKeys: [{ type: 'text', value: 'East' }] }];
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 1);
    assert.notEqual(getPivotRevisionKey(workbook, pivot).filterRevision, firstKey.filterRevision);
  });

  it('keeps same-field filter families independent and rejects disabled multi-filter layouts', () => {
    const workbook = seedCrossSheetWorkbook();
    const pivot = pivotDefinition();
    assert.ok(pivot);
    const region = pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!.fieldId;
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!.fieldId;
    const manual = { kind: 'manual' as const, family: 'manual' as const, fieldId: region, mode: 'include' as const, memberKeys: [{ type: 'text' as const, value: 'East' }] };
    const label = { kind: 'condition' as const, family: 'label' as const, fieldId: region, operator: 'contains' as const, value: 'East' };
    const value = { kind: 'condition' as const, family: 'value' as const, fieldId: region, valueFieldId: amount, operator: 'greater-than' as const, value: 5 };
    const withManual = upsertPivotFilter(pivot.layout, manual);
    const withLabel = upsertPivotFilter(withManual, label);
    const withValue = upsertPivotFilter(withLabel, value);
    assert.deepEqual(withValue.filters.map((filter) => filter.family), ['manual', 'label', 'value']);
    assert.equal(computePivotResult(workbook, { ...pivot, layout: withValue }).grandTotal?.values[0], 10);
    assert.equal(upsertPivotFilter(withValue, { ...label, operator: 'equals', value: 'East' }).filters.length, 3);
    assert.equal(clearPivotFilterFamily(withValue, region, 'label').filters.length, 2);
    assert.equal(clearPivotFiltersForField(withValue, region).filters.length, 0);
    assert.throws(() => computePivotResult(workbook, { ...pivot, layout: { ...withValue, allowMultipleFiltersPerField: false } }), /Multiple Pivot filters are disabled/);
  });
});
