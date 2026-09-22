import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createPivotMemberKey, type PivotModel } from '@react-sheets/core-model';
import { WorkbookSession } from './workbook-session';
import { createInlineJsonQuery } from './features/query';
import { InlinePivotTaskPort, type PivotTaskPort } from './features/pivot/task-port';
import { clearPivotResultCache } from './features/pivot/engine';
import { resolveChartDataFromSources, type ChartPayload } from './features/chart';
import { setCellPatch, writeCellPatch } from './features/data-source';

function seed(app: WorkbookSession): { sheetId: string; pivot: PivotModel } {
  const sheetId = app.getActiveSheetId();
  app.runCommand('sheet.range.set', {
    sheetId,
    startRow: 0,
    startColumn: 0,
    values: [
      [{ value: 'Region' }, { value: 'Amount' }],
      [{ value: 'East' }, { value: 10 }],
      [{ value: 'West' }, { value: 20 }],
    ],
  });
  const range = { sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
  const fields = app.getPivotFieldCatalog(range);
  const region = fields.find((field) => field.name === 'Region')!;
  const amount = fields.find((field) => field.name === 'Amount')!;
  return {
    sheetId,
    pivot: {
      schema: 'PivotDefinition',
      id: 'pivot-test',
      source: { kind: 'worksheet-range', range },
      target: { sheetId, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: {
        rows: [{ fieldId: region.fieldId }],
        columns: [],
        filters: [],
        allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }],
        subtotalLocation: 'bottom',
        showRowGrandTotals: true,
        showColumnGrandTotals: true,
        reportLayout: 'compact',
        calculatedFields: [],
        calculatedItems: [],
        expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
      },
    },
  };
}

async function waitForPivot(app: WorkbookSession, pivotId: string): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
  for (let attempt = 0; attempt < 200; attempt += 1) {
    const state = app.getUiSnapshot().pivotTaskStates[pivotId];
    if (!state || state.status === 'idle') return;
    if (state.status === 'failed') throw new Error(`${state.error.code}: ${state.error.message}`);
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error(`Pivot task did not settle: ${pivotId}`);
}

async function waitForPivotResult(app: WorkbookSession, pivotId: string): Promise<void> {
  for (let attempt = 0; attempt < 200; attempt += 1) {
    if (app['runtime'].pivotResults[pivotId]) return;
    const error = app['runtime'].pivotErrors[pivotId];
    if (error) throw new Error(`${error.code}: ${error.message}`);
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error(`Pivot result did not become available: ${pivotId}`);
}

it('loads and refreshes a cross-sheet PivotChart dependency without projecting unrelated sheets', async () => {
  const app = new WorkbookSession();
  try {
    const { sheetId, pivot } = seed(app);
    await app.addPivot(pivot);
    app.runCommand('sheet.add', { id: 'dashboard', name: 'Dashboard' });
    app.runCommand('sheet.add', { id: 'unrelated', name: 'Unrelated' });
    app.selectSheet('dashboard');
    delete app['runtime'].pivotResults[pivot.id];
    clearPivotResultCache(app['runtime'].model, pivot.id);
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'dashboard-chart', chartType: 'combo', subtype: 'custom-combo',
      source: { kind: 'pivot', pivotId: pivot.id }, elements: {},
    };
    app.runCommand('chart.insert', {
      sheetId: 'dashboard', payload,
      drawing: {
        id: 'dashboard-drawing', sheetId: 'dashboard', kind: 'chart', payloadId: payload.chartId,
        anchor: { kind: 'absolute' }, transform: { x: 40, y: 50, width: 360, height: 240, rotation: 0 }, zIndex: 1,
      },
    });
    await waitForPivotResult(app, pivot.id);
    let snapshot = app.getUiSnapshot();
    assert.deepEqual(new Set(snapshot.projectionSheets.map((sheet) => sheet.id)), new Set([sheetId, 'dashboard']));
    assert.equal(snapshot.projectionSheets.find((sheet) => sheet.id === sheetId)?.pivotResults[pivot.id]?.grandTotal?.values[0], 30);
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 1, value: { value: 40 } });
    await waitForPivot(app, pivot.id);
    snapshot = app.getUiSnapshot();
    const results = snapshot.projectionSheets.find((sheet) => sheet.id === sheetId)!.pivotResults;
    assert.equal(results[pivot.id]?.grandTotal?.values[0], 60);
    assert.equal(resolveChartDataFromSources(payload, () => undefined, results).status.kind, 'ready');
    assert.equal(resolveChartDataFromSources({ ...payload, source: { kind: 'pivot', pivotId: 'missing' } }, () => undefined, results).status.kind, 'invalid');
  } finally { app.dispose(); }
});

class DeferredCalculatePort implements PivotTaskPort {
  private readonly inner = new InlinePivotTaskPort();
  private deferred: {
    request: Parameters<PivotTaskPort['submit']>[0];
    resolve: (result: Awaited<ReturnType<PivotTaskPort['submit']>>) => void;
  } | null = null;
  delayNextCalculate = false;

  submit(request: Parameters<PivotTaskPort['submit']>[0]): ReturnType<PivotTaskPort['submit']> {
    if (request.kind !== 'calculate' || !this.delayNextCalculate) return this.inner.submit(request);
    this.delayNextCalculate = false;
    return new Promise((resolve) => {
      this.deferred = { request, resolve };
    });
  }

  async waitForDeferred(): Promise<void> {
    for (let attempt = 0; attempt < 100; attempt += 1) {
      if (this.deferred) return;
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    throw new Error('Pivot calculate task was not deferred');
  }

  releaseDeferred(): void {
    const deferred = this.deferred;
    if (!deferred) throw new Error('No deferred Pivot calculate task exists');
    this.deferred = null;
    void this.inner.submit(deferred.request).then(deferred.resolve);
  }

  cancel(taskId: string): void {
    if (this.deferred?.request.taskId === taskId) {
      const { request, resolve } = this.deferred;
      this.deferred = null;
      resolve({ protocol: 'react-sheets/pivot-task', version: 1, taskId, generation: request.generation, status: 'cancelled' });
      return;
    }
    this.inner.cancel(taskId);
  }

  dispose(): void {
    this.deferred = null;
    this.inner.dispose();
  }
}

describe('WorkbookSession PivotTable integration', () => {
  it('enters and leaves a validated PivotTable contextual Ribbon state', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    assert.equal((await app.addPivot(pivot)).status, 'created');

    app.setActivePivotContext(pivot.id, sheetId);
    let snapshot = app.getUiSnapshot();
    assert.deepEqual(snapshot.activeContext, { kind: 'pivot', sheetId, pivotId: pivot.id });
    assert.equal(snapshot.ribbon.activeTab, 'pivotAnalyze');
    assert.equal(snapshot.panels.active, 'pivot');
    assert.equal(snapshot.panels.open, true);

    app.setActivePivotContext(null, sheetId);
    snapshot = app.getUiSnapshot();
    assert.deepEqual(snapshot.activeContext, { kind: 'none' });
    assert.equal(snapshot.ribbon.activeTab, 'home');
  });

  it('rejects a fabricated PivotTable context without changing the active context', () => {
    const app = new WorkbookSession();
    assert.throws(() => app.setActivePivotContext('missing-pivot'), /Unknown PivotTable context: missing-pivot/);
    assert.deepEqual(app.getUiSnapshot().activeContext, { kind: 'none' });
  });

  it('creates a new worksheet and PivotTable as one history entry', async () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const beforeHistory = app.getUiSnapshot().historyEntries.length;
    const created = await app.createPivotTable({ destination: { kind: 'new-sheet' } });
    assert.equal(created.status, 'created');
    if (created.status !== 'created') return;
    const pivotId = created.pivotId;
    const createdSheetId = app.getActiveSheetId();
    assert.notEqual(createdSheetId, sheetId);
    assert.equal(app.getUiSnapshot().historyEntries.length, beforeHistory + 1);
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.id, pivotId);

    app.undo();
    assert.equal(app.getUiSnapshot().sheets.some((sheet) => sheet.id === createdSheetId), false);
    app.redo();
    app.selectSheet(createdSheetId);
    const restoredSnapshot = app.getUiSnapshot();
    assert.equal(restoredSnapshot.selectedSheet.id, createdSheetId);
    assert.equal(restoredSnapshot.selectedSheet.pivots[0]?.id, pivotId);
  });

  it('keeps the worksheet scope on a named-range source through create, undo, and redo', async () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.setDefinedName({ name: 'SharedSource', formula: "='Sheet1'!A1:B3", scope: 'workbook' });
    app.setDefinedName({ name: 'SharedSource', formula: "='Sheet1'!A1:B3", scope: 'sheet', sheetId });

    const created = await app.createPivotTable({
      source: { kind: 'named-range', name: 'SharedSource', sheetId },
      destination: { kind: 'existing-sheet', sheetId, anchor: { row: 5, column: 0 } },
    });
    assert.equal(created.status, 'created', app.getUiSnapshot().notice);
    if (created.status !== 'created') return;
    const pivotId = created.pivotId;
    const source = app.getUiSnapshot().selectedSheet.pivots.find((pivot) => pivot.id === pivotId)?.source;
    assert.deepEqual(source, { kind: 'named-range', name: 'SharedSource', sheetId });

    app.undo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots.some((pivot) => pivot.id === pivotId), false);
    app.redo();
    assert.deepEqual(app.getUiSnapshot().selectedSheet.pivots.find((pivot) => pivot.id === pivotId)?.source, { kind: 'named-range', name: 'SharedSource', sheetId });
  });

  it('leaves workbook and history unchanged when create preflight rejects duplicate headers', async () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 'Region' } });
    const before = app.getUiSnapshot();
    const outcome = await app.createPivotTable({ destination: { kind: 'new-sheet' } });
    assert.equal(outcome.status, 'rejected');
    const after = app.getUiSnapshot();
    assert.equal(after.sheets.length, before.sheets.length);
    assert.equal(after.historyEntries.length, before.historyEntries.length);
    assert.equal(after.notice, 'Pivot source header is duplicated: Region');
  });

  it('rejects PivotTable creation for a viewer before any worksheet mutation', async () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);
    const before = app.getUiSnapshot();
    assert.equal((await app.createPivotTable({ destination: { kind: 'new-sheet' } })).status, 'rejected');
    const after = app.getUiSnapshot();
    assert.equal(after.sheets.length, before.sheets.length);
    assert.equal(after.historyEntries.length, before.historyEntries.length);
    assert.match(after.notice, /cannot perform|Permission denied/);
    assert.equal(sheetId, app.getActiveSheetId());
  });

  it('addPivot computes a local result tree', async () => {
    const app = new WorkbookSession();
    const { pivot } = seed(app);
    assert.equal((await app.addPivot(pivot)).status, 'created');
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.pivots.length, 1);
    assert.ok(snapshot.selectedSheet.pivotResults[pivot.id]);
    assert.ok(snapshot.selectedSheet.pivotResults[pivot.id]!.rows.length > 0);
  });

  it('insertPivotFromSelection builds a PivotTable from the current selection', async () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const created = await app.insertPivotFromSelection();
    assert.equal(created.status, 'created');
    if (created.status !== 'created') return;
    assert.ok(app.getUiSnapshot().selectedSheet.pivotResults[created.pivotId]);
  });

  it('drillDownPivot creates a detail worksheet through the canonical command', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-drill';
    await app.addPivot(pivot);
    app.runCommand('sheet.add', { id: 'other-sheet', name: 'Other' });
    app.selectSheet('other-sheet');
    const beforeCount = app.getUiSnapshot().sheets.length;
    await app.drillDownPivot(pivot.id, 'East', [{ sheetId, row: 1 }]);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.sheets.length, beforeCount + 1);
    assert.notEqual(snapshot.activeSheetId, sheetId);
  });

  it('drillDownPivot grows the detail worksheet for more than the default row extent', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-drill-large';
    if (pivot.source.kind !== 'worksheet-range') throw new Error('Expected a single worksheet Pivot source');
    const sourceSheet = app['runtime'].model.getSheet(sheetId);
    for (let row = 3; row <= 1_200; row += 1) {
      sourceSheet.cells.set(row, 0, { value: 'East' });
      sourceSheet.cells.set(row, 1, { value: row });
    }
    sourceSheet.rowCount = 2_000;
    pivot.source = { ...pivot.source, range: { ...pivot.source.range, endRow: 1_200 } };
    pivot.target = { ...pivot.target, anchor: { row: 1_300, column: 0 } };
    const added = await app.addPivot(pivot);
    if (added.status === 'rejected') throw new Error(added.error.message);
    assert.equal(added.status, 'created');

    await app.drillDownPivot(pivot.id, 'Large', Array.from({ length: 1_200 }, (_, index) => ({ sheetId, row: index + 1 })));

    const detailSheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(detailSheet.rowCount, 1_201);
    assert.equal(detailSheet.columnCount, 26);
    const detailRegion = detailSheet.dataRegions[0]!;
    const detailQuery = app['runtime'].dataContent.get(detailRegion.sourceId)!;
    assert.deepEqual((await detailQuery.getRows(1_199, 1)).value?.[0], ['East', 1_200]);
  });

  it('creates a slicer drawing and refreshes a derived result without persisted refresh state', async () => {
    const app = new WorkbookSession();
    const { pivot } = seed(app);
    pivot.id = 'pivot-slicer';
    await app.addPivot(pivot);
    app.createPivotSlicerControl(pivot.id, pivot.fieldCatalog.fields[0]!.fieldId);
    app.refreshPivot(pivot.id);
    await waitForPivot(app, pivot.id);
    assert.equal(app.listPivotControls(pivot.id).length, 1);
    assert.ok(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]);
  });

  it('recomputes linked PivotTables when a Slicer filter changes', async () => {
    const app = new WorkbookSession();
    const { pivot } = seed(app);
    pivot.id = 'pivot-slicer-filter';
    await app.addPivot(pivot);
    app.createPivotSlicerControl(pivot.id, pivot.fieldCatalog.fields[0]!.fieldId);
    const slicer = app.listPivotControls(pivot.id).find((control) => control.payload.kind === 'slicer');
    assert.ok(slicer);
    if (!slicer) return;

    app.setPivotSlicerFilter(slicer.drawing.id, 'include', [createPivotMemberKey('East')]);
    await waitForPivot(app, pivot.id);

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 10);

    app.removePivotControl(slicer.drawing.id);
    await waitForPivot(app, pivot.id);
    assert.equal(app.listPivotControls(pivot.id).length, 0);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 30);
  });

  it('refreshes only linked PivotTables across different target sheets', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-linked-primary';
    await app.addPivot(pivot);

    app.runCommand('sheet.add', { id: 'pivot-linked-secondary-sheet', name: 'Linked secondary' });
    const secondary = structuredClone(pivot);
    secondary.id = 'pivot-linked-secondary';
    secondary.target = { sheetId: 'pivot-linked-secondary-sheet', anchor: { row: 0, column: 0 } };
    await app.addPivot(secondary);

    app.createPivotSlicerControl(pivot.id, pivot.fieldCatalog.fields[0]!.fieldId);
    const slicer = app.listPivotControls(pivot.id).find((control) => control.payload.kind === 'slicer');
    assert.ok(slicer);
    if (!slicer) return;
    const connection = app.listCompatiblePivotControlConnections(pivot.id, pivot.fieldCatalog.fields[0]!.fieldId, 'slicer')
      .find((candidate) => candidate.pivotId === secondary.id);
    assert.ok(connection);
    if (!connection) return;
    app.setPivotControlConnections(slicer.drawing.id, [connection]);

    app.setPivotSlicerFilter(slicer.drawing.id, 'include', [createPivotMemberKey('East')]);
    await Promise.all([waitForPivot(app, pivot.id), waitForPivot(app, secondary.id)]);

    assert.equal(app['runtime'].pivotResults[pivot.id]?.grandTotal?.values[0], 10);
    assert.equal(app['runtime'].pivotResults[secondary.id]?.grandTotal?.values[0], 10);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 10);
  });

  it('refreshes a PivotTable when a timeline period changes through the public API', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Date' }, { value: 'Amount' }],
        [{ value: '2026-08-25T00:00:00' }, { value: 10 }],
        [{ value: '2026-08-25T12:00:00' }, { value: 20 }],
        [{ value: '2026-08-25T23:59:59' }, { value: 30 }],
        [{ value: '2026-08-26T00:00:00' }, { value: 40 }],
      ],
    });
    const range = { sheetId, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 };
    const fields = app.getPivotFieldCatalog(range);
    const date = fields.find((field) => field.name === 'Date')!;
    const amount = fields.find((field) => field.name === 'Amount')!;
    const pivot: PivotModel = {
      schema: 'PivotDefinition',
      id: 'pivot-timeline-api',
      source: { kind: 'worksheet-range', range },
      target: { sheetId, anchor: { row: 7, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: {
        rows: [],
        columns: [],
        filters: [],
        allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }],
        subtotalLocation: 'bottom',
        showRowGrandTotals: true,
        showColumnGrandTotals: true,
        reportLayout: 'compact',
        calculatedFields: [],
        calculatedItems: [],
        expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
      },
    };
    await app.addPivot(pivot);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 100);

    app.createPivotTimelineControl(pivot.id, date.fieldId);
    const timeline = app.listPivotControls(pivot.id).find((control) => control.payload.kind === 'timeline');
    assert.ok(timeline);
    if (!timeline) return;
    app.setPivotTimelinePeriod(timeline.drawing.id, '2026-08-25', '2026-08-25');
    await waitForPivot(app, pivot.id);

    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 60);
    const updatedTimeline = app.listPivotControls(pivot.id).find((control) => control.drawing.id === timeline.drawing.id);
    assert.deepEqual(updatedTimeline?.payload.kind === 'timeline' ? updatedTimeline.payload.period : undefined, { start: '2026-08-25', end: '2026-08-25' });
  });

  it('keeps an explicit block-backed worksheet source on the DataSource Pivot path', async () => {
    const app = new WorkbookSession();
    await app.loadQuery(createInlineJsonQuery('pivot-block-source', 'Pivot block source', [
      { ID: 'A', Amount: 1 },
      { ID: 'B', Amount: 2 },
    ]));
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    const region = sheet.dataRegions[0]!;

    const created = await app.createPivotTable({
      source: { kind: 'worksheet-range', range: structuredClone(region.range) },
      destination: { kind: 'new-sheet' },
    });
    assert.equal(created.status, 'created', app.getUiSnapshot().notice);
    if (created.status !== 'created') return;

    const pivot = app['runtime'].model.getSheets()
      .flatMap((entry) => entry.pivots)
      .find((entry) => entry.id === created.pivotId)!;
    assert.deepEqual(pivot.source, { kind: 'data-source', dataSourceId: region.sourceId });
    assert.deepEqual(pivot.fieldCatalog.fields.map((field) => field.fieldId), [
      `${region.sourceId}:field:0`,
      `${region.sourceId}:field:1`,
    ]);

    writeCellPatch(sheet, 1, 1, { schema: 'CellPatch', value: setCellPatch(11) });
    app.createPivotSlicerControl(pivot.id, pivot.fieldCatalog.fields[0]!.fieldId);
    app.refreshPivot(pivot.id);
    await waitForPivot(app, pivot.id);
    const result = app['runtime'].pivotResults[pivot.id];
    assert.equal(result?.grandTotal?.values[0], 13);
    const slicer = Object.values(result?.slicerItems ?? {})[0] ?? [];
    assert.deepEqual(slicer.map((item) => item.label), ['A', 'B']);
    assert.ok(result?.sourceRowPaths.length);
    await app.drillDownPivot(pivot.id, 'All rows', result!.sourceRowPaths);
    const detailSheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const detailRegion = detailSheet.dataRegions[0]!;
    assert.notEqual(detailRegion.sourceId, region.sourceId);
    const detailRows = await app['runtime'].dataContent.get(detailRegion.sourceId)!.getRows(0, 2);
    assert.deepEqual(detailRows.value, [['A', 11], ['B', 2]]);
  });

  it('loads DataSource Pivot field members only when a picker requests them', async () => {
    const app = new WorkbookSession();
    await app.loadQuery(createInlineJsonQuery('pivot-lazy-members', 'Pivot lazy members', [
      { Region: 'East', Amount: 10 },
      { Region: 'West', Amount: 20 },
    ]));
    const sheetId = app.getActiveSheetId();
    const region = app['runtime'].model.getSheet(sheetId).dataRegions[0]!;
    const created = await app.createPivotTable({
      source: { kind: 'worksheet-range', range: structuredClone(region.range) },
      destination: { kind: 'new-sheet' },
    });
    assert.equal(created.status, 'created', app.getUiSnapshot().notice);
    if (created.status !== 'created') return;

    const pivot = app['runtime'].model.getSheets().flatMap((entry) => entry.pivots).find((entry) => entry.id === created.pivotId)!;
    const regionField = pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!;
    assert.deepEqual(app.getPivotFieldCatalogForPivot(pivot.id).find((field) => field.fieldId === regionField.fieldId)?.values, []);

    await app.loadPivotFieldValues(pivot.id, regionField.fieldId);

    assert.deepEqual(app.getPivotFieldCatalogForPivot(pivot.id).find((field) => field.fieldId === regionField.fieldId)?.values, ['East', 'West']);
  });

  it('creates a DataSource Pivot timeline from Query date fields', async () => {
    const app = new WorkbookSession();
    await app.loadQuery(createInlineJsonQuery('pivot-date-block-source', 'Pivot date blocks', [
      { PostedAt: '2026-08-25T00:00:00', Amount: 10 },
      { PostedAt: '2026-08-25T12:00:00', Amount: 20 },
      { PostedAt: '2026-08-25T23:59:59', Amount: 30 },
      { PostedAt: '2026-08-26T00:00:00', Amount: 40 },
    ]));
    const sheetId = app.getActiveSheetId();
    const region = app['runtime'].model.getSheet(sheetId).dataRegions[0]!;
    const created = await app.createPivotTable({
      source: { kind: 'worksheet-range', range: structuredClone(region.range) },
      destination: { kind: 'new-sheet' },
    });
    assert.equal(created.status, 'created', app.getUiSnapshot().notice);
    if (created.status !== 'created') return;

    const pivot = app['runtime'].model.getSheets()
      .flatMap((entry) => entry.pivots)
      .find((entry) => entry.id === created.pivotId)!;
    const date = pivot.fieldCatalog.fields.find((field) => field.name === 'PostedAt');
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount');
    assert.ok(date);
    assert.ok(amount);
    if (!date || !amount) return;
    assert.equal(date.dataType, 'date');
    const layout = structuredClone(pivot.layout);
    layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const updated = await app.updatePivotLayout(pivot.id, layout);
    assert.equal(updated.status, 'updated');
    await waitForPivotResult(app, pivot.id);
    assert.equal(app['runtime'].pivotResults[pivot.id]?.grandTotal?.values[0], 100);

    app.createPivotTimelineControl(pivot.id, date.fieldId);
    const timeline = app.listPivotControls(pivot.id).find((control) => control.payload.kind === 'timeline');
    assert.ok(timeline);
    if (!timeline) return;
    app.setPivotTimelinePeriod(timeline.drawing.id, '2026-08-25', '2026-08-25');
    await waitForPivot(app, pivot.id);

    assert.equal(app['runtime'].pivotResults[pivot.id]?.grandTotal?.values[0], 60);
  });

  it('recomputes the Pivot projection when pivot.update enters through the public dispatch path', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-dispatch';
    await app.addPivot(pivot);
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
    const nextLayout = structuredClone(pivot.layout);
    nextLayout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'count' }];
    const dispatch = await app.updatePivotLayout(pivot.id, nextLayout);
    assert.equal(dispatch.status, 'updated');
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 2);
  });

  it('keeps manual and on-open results stale until the explicit refresh command', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-manual-policy';
    pivot.refreshPolicy = { mode: 'manual', preserveFormatting: true, refreshOnLoad: false };
    await app.addPivot(pivot);
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 1, value: { value: 100 } });
    const stale = app.getUiSnapshot().selectedSheet;
    assert.equal(stale.pivotResults[pivot.id]?.grandTotal?.values[0], 30);
    assert.equal(stale.pivotProjections[pivot.id]?.refresh.status, 'stale');

    app.refreshPivot(pivot.id);
    await waitForPivot(app, pivot.id);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 120);

    const onOpen = structuredClone(pivot);
    onOpen.id = 'pivot-on-open-policy';
    onOpen.target = { sheetId, anchor: { row: 12, column: 0 } };
    onOpen.refreshPolicy = { mode: 'on-open', preserveFormatting: true, refreshOnLoad: true };
    await app.addPivot(onOpen);
    app.runCommand('sheet.cell.set', { sheetId, row: 2, column: 1, value: { value: 200 } });
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[onOpen.id]?.grandTotal?.values[0], 120);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotProjections[onOpen.id]?.refresh.status, 'stale');
  });

  it('forces an explicit refresh through a truthful loading state even when the proof is current', async () => {
    const taskPort = new DeferredCalculatePort();
    const app = new WorkbookSession({ pivotTaskPort: taskPort });
    const { pivot } = seed(app);
    await app.addPivot(pivot);
    await waitForPivot(app, pivot.id);

    taskPort.delayNextCalculate = true;
    app.refreshPivot(pivot.id);
    await taskPort.waitForDeferred();

    assert.equal(app['runtime'].pivotResults[pivot.id], undefined);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]?.refresh.status, 'refreshing');

    taskPort.releaseDeferred();
    await waitForPivot(app, pivot.id);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]?.refresh.status, 'ready');
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 30);
  });

  it('discards an obsolete worker result and recalculates from the committed source revision', async () => {
    const taskPort = new DeferredCalculatePort();
    const app = new WorkbookSession({ pivotTaskPort: taskPort });
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-obsolete-worker-result';
    pivot.refreshPolicy = { mode: 'manual', preserveFormatting: true, refreshOnLoad: false };
    await app.addPivot(pivot);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 30);

    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 1, value: { value: 100 } });
    taskPort.delayNextCalculate = true;
    app.refreshPivot(pivot.id);
    await taskPort.waitForDeferred();

    app.runCommand('sheet.cell.set', { sheetId, row: 2, column: 1, value: { value: 200 } });
    taskPort.releaseDeferred();
    await waitForPivot(app, pivot.id);

    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 300);
    assert.notEqual(app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]?.refresh.status, 'stale');
  });

  it('refreshes only intersecting on-change sources and supports refresh all explicitly', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-on-change-policy';
    await app.addPivot(pivot);
    const other = structuredClone(pivot);
    other.id = 'pivot-manual-peer';
    other.target = { sheetId, anchor: { row: 12, column: 0 } };
    other.refreshPolicy = { mode: 'manual', preserveFormatting: true, refreshOnLoad: false };
    await app.addPivot(other);
    app.runCommand('sheet.cell.set', { sheetId, row: 10, column: 1, value: { value: 999 } });
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 30);
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 1, value: { value: 40 } });
    await waitForPivot(app, pivot.id);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 60);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[other.id]?.grandTotal?.values[0], 30);
    app.runCommand('sheet.cell.set', { sheetId, row: 2, column: 1, value: { value: 50 } });
    app.refreshAllPivots();
    await waitForPivot(app, other.id);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[other.id]?.grandTotal?.values[0], 90);
  });

  it('updates row and column grand-total state as one undoable layout mutation', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-grand-total-undo';
    await app.addPivot(pivot);
    const nextLayout = structuredClone(pivot.layout);
    nextLayout.showRowGrandTotals = false;
    nextLayout.showColumnGrandTotals = true;
    const dispatch = await app.updatePivotLayout(pivot.id, nextLayout);
    assert.equal(dispatch.status, 'updated');
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.showRowGrandTotals, false);
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.showColumnGrandTotals, true);
    app.undo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.showRowGrandTotals, true);
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.showColumnGrandTotals, true);
    app.redo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.showRowGrandTotals, false);
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.showColumnGrandTotals, true);
  });

  it('persists and restores the canonical report layout through undo and redo', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-report-layout-undo';
    await app.addPivot(pivot);
    const nextLayout = structuredClone(pivot.layout);
    nextLayout.reportLayout = 'tabular';
    const dispatch = await app.updatePivotLayout(pivot.id, nextLayout);
    assert.equal(dispatch.status, 'updated');
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.reportLayout, 'tabular');
    app.undo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.reportLayout, 'compact');
    app.redo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.layout.reportLayout, 'tabular');
  });

  it('persists Pivot style options through one reversible presentation update', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-style';
    await app.addPivot(pivot);
    const presentation = { styleName: 'PivotStyleMedium4', styleOptions: { showRowHeaders: false, showColumnHeaders: true, showRowStripes: true, showColumnStripes: false, showLastColumn: true } } as const;
    const dispatch = await app.updatePivotConfiguration(pivot.id, { presentation });
    assert.equal(dispatch.status, 'updated');
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.presentation?.styleName, presentation.styleName);
    assert.deepEqual(app.getUiSnapshot().selectedSheet.pivots[0]?.presentation?.styleOptions, presentation.styleOptions);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]?.presentation?.styleName, 'PivotStyleMedium4');
    app.undo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.presentation?.styleName, undefined);
    app.redo();
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.presentation?.styleOptions.showRowStripes, true);
  });

  it('toggles Pivot expansion through one reversible command while keeping the parent row', async () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-expansion';
    const region = pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!;
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }, { fieldId: amount.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    await app.addPivot(pivot);
    const tree = app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]!;
    const parent = tree.rows[0]!;
    assert.ok(parent.nodeId);
    const beforeHistory = app.getUiSnapshot().historyEntries.length;
    const collapsed = await app.togglePivotExpansion(pivot.id, parent.nodeId!);
    assert.equal(collapsed.status, 'updated');
    const collapsedProjection = app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]!;
    assert.equal(collapsedProjection.cells.some((cell) => cell.nodeId === parent.nodeId && cell.expanded === false), true);
    assert.equal(collapsedProjection.cells.filter((cell) => cell.nodeId && cell.nodeId.startsWith(`${parent.nodeId}/`)).length, 0);
    assert.equal(app.getUiSnapshot().historyEntries.length, beforeHistory + 1);
    app.undo();
    await waitForPivot(app, pivot.id);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]!.cells.some((cell) => cell.nodeId?.startsWith(`${parent.nodeId}/`)), true);
    app.redo();
    await waitForPivot(app, pivot.id);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotProjections[pivot.id]!.cells.some((cell) => cell.nodeId === parent.nodeId && cell.expanded === false), true);
  });
});
