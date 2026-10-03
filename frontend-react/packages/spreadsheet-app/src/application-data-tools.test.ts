import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

function selectRange(
  app: WorkbookSession,
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): void {
  const sheetId = app.getActiveSheetId();
  app.runCommand('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow, endRow, startColumn, endColumn }],
    primaryRangeIndex: 0,
    activeCell: { row: startRow, column: startColumn },
    anchorCell: { row: startRow, column: startColumn },
  });
}

describe('WorkbookSession data tools integration', () => {
  it('data dispatch splits delimited text into columns', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'a,b,c' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 1,
      column: 0,
      value: { value: '1,2,3' },
    });
    selectRange(app, 0, 0, 1, 0);
    const range = app.getPrimaryRange();
    const outcome = await app.dispatch({ commandId: 'data.textToColumns', params: { sheetId, range, delimiter: ',', maxColumns: 8 } });
    assert.equal(outcome.status, 'committed');

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'a');
    assert.equal(sheet.cells.get(0, 1)?.value, 'b');
    assert.equal(sheet.cells.get(0, 2)?.value, 'c');
    assert.equal(sheet.cells.get(1, 0)?.value, '1');
    assert.equal(sheet.cells.get(1, 2)?.value, '3');
  });

  it('preflights data-tool inputs before dispatch materializes intersecting regions', async () => {
    const cases: Array<{ commandId: string; params: (sheetId: string) => Record<string, unknown>; region: (sheetId: string) => { startRow: number; endRow: number; startColumn: number; endColumn: number } }> = [
      {
        commandId: 'data.textToColumns',
        params: (sheetId) => ({ sheetId, range: { sheetId, startRow: 0, endRow: 50_000, startColumn: 0, endColumn: 0 }, delimiter: ',', maxColumns: 2 }),
        region: (sheetId) => ({ startRow: 0, endRow: 50_000, startColumn: 0, endColumn: 0 }),
      },
      {
        commandId: 'data.removeDuplicates',
        params: (sheetId) => ({ sheetId, range: { sheetId, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }, columns: [2] }),
        region: () => ({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }),
      },
      {
        commandId: 'data.subtotal',
        params: (sheetId) => ({ sheetId, range: { sheetId, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }, groupColumn: 2, valueColumn: 1, functionName: 'SUM' }),
        region: () => ({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }),
      },
      {
        commandId: 'data.splitColumn',
        params: (sheetId) => ({ sheetId, row: 0, column: 0, delimiter: '', maxColumns: 4 }),
        region: () => ({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }),
      },
    ];

    for (const [index, entry] of cases.entries()) {
      const app = new WorkbookSession();
      const sheetId = app.getActiveSheetId();
      const sheet = app['runtime'].model.getSheet(sheetId);
      if (entry.commandId === 'data.textToColumns') sheet.rowCount = 50_001;
      sheet.addDataRegion({
        id: `preflight-region-${index}`,
        sourceId: `unloaded-source-${index}`,
        range: { sheetId, ...entry.region(sheetId) },
        headerRow: 0,
        revision: 0,
      });
      let materializationCalls = 0;
      app['materializeDataRegions'] = async () => { materializationCalls += 1; };

      const result = await app.dispatch({ commandId: entry.commandId, params: entry.params(sheetId) });

      assert.equal(result.status, 'rejected', entry.commandId);
      if (result.status === 'rejected') assert.equal(result.error.code, 'COMMAND_REJECTED');
      assert.equal(materializationCalls, 0, entry.commandId);
    }
  });

  it('checks permission and planner availability before loading block-backed data', async () => {
    const viewer = new WorkbookSession();
    const viewerSheetId = viewer.getActiveSheetId();
    viewer['runtime'].model.getSheet(viewerSheetId).addDataRegion({
      id: 'viewer-region',
      sourceId: 'viewer-unloaded-source',
      range: { sheetId: viewerSheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      headerRow: 0,
      revision: 0,
    });
    viewer['permission'].applyServerAccess({ unitId: 'wb-test', role: 'viewer', accessRevision: 1, regions: [] });
    viewer['permission'].setOnline(true);
    let viewerMaterializationCalls = 0;
    viewer['materializeDataRegions'] = async () => { viewerMaterializationCalls += 1; };

    const viewerResult = await viewer.dispatch({
      commandId: 'data.textToColumns',
      params: {
        sheetId: viewerSheetId,
        range: { sheetId: viewerSheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        delimiter: ',',
        maxColumns: 2,
      },
    });

    assert.equal(viewerResult.status, 'rejected');
    if (viewerResult.status === 'rejected') assert.match(viewerResult.error.message, /permission|viewer|edit/i);
    assert.equal(viewerMaterializationCalls, 0);

    const offline = new WorkbookSession();
    const offlineSheetId = offline.getActiveSheetId();
    offline['runtime'].model.getSheet(offlineSheetId).addDataRegion({
      id: 'offline-region',
      sourceId: 'offline-unloaded-source',
      range: { sheetId: offlineSheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      headerRow: 0,
      revision: 0,
    });
    let offlineMaterializationCalls = 0;
    offline['materializeDataRegions'] = async () => { offlineMaterializationCalls += 1; };

    const offlineResult = await offline.dispatch({
      commandId: 'sheet.rows.insert',
      params: { sheetId: offlineSheetId, at: 0, count: 1 },
    });

    assert.equal(offlineResult.status, 'rejected');
    if (offlineResult.status === 'rejected') assert.match(offlineResult.error.message, /STRUCTURAL_PLANNER_OFFLINE/);
    assert.equal(offlineMaterializationCalls, 0);
  });

  it('does not materialize a merge range before checking edit permission', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app['runtime'].model.getSheet(sheetId).addDataRegion({
      id: 'protected-merge-region',
      sourceId: 'protected-merge-source',
      range: { sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      headerRow: 0,
      revision: 0,
    });
    app['permission'].applyServerAccess({ unitId: 'wb-test', role: 'viewer', accessRevision: 1, regions: [] });
    app['permission'].setOnline(true);
    selectRange(app, 0, 0, 1, 1);
    let materializationCalls = 0;
    app['materializeDataRegions'] = async () => { materializationCalls += 1; };

    app.requestMergeAction('center');

    assert.equal(materializationCalls, 0);
    assert.match(app.getUiSnapshot().notice, /permission|viewer|edit/i);
  });

  it('rejects data-source creation before loading or persisting a viewer range', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app['runtime'].model.getSheet(sheetId).addDataRegion({
      id: 'viewer-source-region',
      sourceId: 'viewer-source-blocks',
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      headerRow: 0,
      revision: 0,
    });
    app['permission'].applyServerAccess({ unitId: 'wb-test', role: 'viewer', accessRevision: 1, regions: [] });
    app['permission'].setOnline(true);
    let materializationCalls = 0;
    app['materializeDataRegions'] = async () => { materializationCalls += 1; };

    await assert.rejects(() => app.createDataSourceFromSelection(), /permission|viewer|structure/i);

    assert.equal(materializationCalls, 0);
    assert.equal(app['runtime'].model.dataModel.sources.size, 0);
  });

  it('materializes the addressed cell region for splitColumn, not the current selection region', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app['runtime'].model.getSheet(sheetId).addDataRegion({
      id: 'split-target-region',
      sourceId: 'unloaded-split-source',
      range: { sheetId, startRow: 8, endRow: 8, startColumn: 3, endColumn: 3 },
      headerRow: 8,
      revision: 0,
    });
    selectRange(app, 0, 0, 0, 0);
    const materializedRegionIds: string[] = [];
    app['materializeDataRegions'] = async (regions) => {
      materializedRegionIds.push(...regions.map((region) => region.id));
    };

    const result = await app.dispatch({
      commandId: 'data.splitColumn',
      params: { sheetId, row: 8, column: 3, delimiter: ',', maxColumns: 4 },
    });

    assert.equal(result.status, 'committed');
    assert.deepEqual(materializedRegionIds, ['split-target-region']);
  });

  it('canonicalizes a command range to its sheet before matching lazy data regions', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app['runtime'].model.getSheet(sheetId).addDataRegion({
      id: 'canonical-range-region',
      sourceId: 'unloaded-canonical-range-source',
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      headerRow: 0,
      revision: 0,
    });
    const materializedRegionIds: string[] = [];
    app['materializeDataRegions'] = async (regions) => {
      materializedRegionIds.push(...regions.map((region) => region.id));
    };

    const result = await app.dispatch({
      commandId: 'data.textToColumns',
      params: {
        sheetId,
        range: { sheetId: 'stale-selection-sheet', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        delimiter: ',',
        maxColumns: 2,
      },
    });

    assert.equal(result.status, 'committed');
    assert.deepEqual(materializedRegionIds, ['canonical-range-region']);
  });

  it('data dispatch keeps unique rows', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Key' }, { value: 'Value' }],
        [{ value: 'A' }, { value: 1 }],
        [{ value: 'A' }, { value: 1 }],
        [{ value: 'B' }, { value: 2 }],
      ],
    });
    selectRange(app, 0, 0, 3, 1);
    const range = app.getPrimaryRange();
    const outcome = await app.dispatch({ commandId: 'data.removeDuplicates', params: { sheetId, range, columns: [0, 1], hasHeader: true } });
    assert.equal(outcome.status, 'committed');

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'Key');
    assert.equal(sheet.cells.get(1, 0)?.value, 'A');
    assert.equal(sheet.cells.get(2, 0)?.value, 'B');
    assert.equal(sheet.cells.get(3, 0)?.value, undefined);
  });

  it('groupRowsFromSelection adds an outline group to the sheet model', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    selectRange(app, 1, 0, 4, 3);
    app.groupRowsFromSelection();

    const groups = app['runtime'].model.getSheet(sheetId).outline?.groups ?? [];
    assert.equal(groups.length, 1);
    assert.equal(groups[0]?.axis, 'row');
    assert.equal(groups[0]?.start, 1);
    assert.equal(groups[0]?.end, 4);
  });

  it('transposeSelection swaps rows and columns through matrix.transpose', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'A' }, { value: 'B' }],
        [{ value: 1 }, { value: 2 }],
      ],
    });
    selectRange(app, 0, 0, 1, 1);
    app.transposeSelection();

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'A');
    assert.equal(sheet.cells.get(0, 1)?.value, 1);
    assert.equal(sheet.cells.get(1, 0)?.value, 'B');
    assert.equal(sheet.cells.get(1, 1)?.value, 2);
  });

  it('data dispatch writes grouped summary rows', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Group' }, { value: 'Amount' }],
        [{ value: 'East' }, { value: 10 }],
        [{ value: 'East' }, { value: 5 }],
        [{ value: 'West' }, { value: 7 }],
      ],
    });
    selectRange(app, 0, 0, 3, 1);
    const range = app.getPrimaryRange();
    const outcome = await app.dispatch({ commandId: 'data.subtotal', params: { sheetId, range, groupColumn: 0, valueColumn: 1, functionName: 'SUM' } });
    assert.equal(outcome.status, 'committed');

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(5, 0)?.value, 'Group');
    assert.equal(sheet.cells.get(6, 0)?.value, 'East');
    await app.waitForFormulaCalculation();
    assert.equal(app['runtime'].formula.getCellValue({ sheetId, row: 6, column: 1 }), 15);
    assert.equal(sheet.cells.get(7, 0)?.value, 'West');
    assert.equal(app['runtime'].formula.getCellValue({ sheetId, row: 7, column: 1 }), 7);
  });
});
