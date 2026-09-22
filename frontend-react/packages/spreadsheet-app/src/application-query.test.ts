import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';
import { createInlineJsonQuery } from './features/query';

describe('WorkbookSession query integration', () => {
  it('loads inline json queries into the active sheet', async () => {
    const app = new WorkbookSession();
    const query = createInlineJsonQuery('demo-query', 'Demo', [
      { Region: 'East', Units: 12 },
      { Region: 'West', Units: 8 },
    ]);

    await app.loadQuery(query);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.lastQueryResult?.rowCount, 2);
    assert.equal(snapshot.loadedQueries.length, 1);
    assert.equal(snapshot.queryConnectors.includes('json'), true);

    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(sheet.cells.get(0, 0)?.value, 'Region');
    assert.equal(sheet.cells.get(1, 1), undefined);
    assert.equal(sheet.dataRegions[0]?.sourceId, 'query:demo-query');
  });

  it('refreshes a loaded query through query.refresh', async () => {
    const app = new WorkbookSession();
    const query = createInlineJsonQuery('refresh-query', 'Refresh', [{ Value: 1 }]);
    await app.loadQuery(query);
    await app.refreshQuery('refresh-query');
    assert.equal(app.getUiSnapshot().lastQueryResult?.rowCount, 1);
  });

  it('rehydrates a persisted query result projection for refresh after reopen', async () => {
    const app = new WorkbookSession();
    await app.loadQuery(createInlineJsonQuery('rehydrate-query', 'Rehydrate', [
      { Region: 'East', Units: 12 },
      { Region: 'West', Units: 8 },
    ]));

    app['querySessions'].clear();
    app['lastQueryResult'] = null;
    app['restorePersistedQuerySessions']();

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.loadedQueries.length, 1);
    assert.equal(snapshot.loadedQueries[0]?.queryId, 'rehydrate-query');
    assert.equal(snapshot.loadedQueries[0]?.rowCount, 2);
    assert.deepEqual(snapshot.loadedQueries[0]?.columns, ['Region', 'Units']);
    assert.deepEqual(snapshot.loadedQueries[0]?.target.kind, 'range');
  });

  it('sorts a block-backed query through AutoFilter without materializing cells', async () => {
    const app = new WorkbookSession();
    await app.loadQuery(createInlineJsonQuery('filter-sort-query', 'Filter sort', [
      { Key: 'b', Value: 2 },
      { Key: 'a', Value: 1 },
    ]));
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    const region = sheet.dataRegions[0]!;
    const dataRegionContext = {
      ...app.getDataRegionContext(),
      range: structuredClone(region.range),
      currentRegion: structuredClone(region.range),
    };

    const filterResult = await app.dispatch({
      commandId: 'sheet.autoFilter.toggle',
      params: { sheetId, range: region.range, dataRegionContext },
    });
    assert.equal(filterResult.status, 'committed');

    const sortResult = await app.dispatch({
      commandId: 'sheet.autoFilter.sort',
      params: { sheetId, column: region.range.startColumn, ascending: true, dataRegionContext },
    });
    assert.equal(sortResult.status, 'committed');
    assert.equal(sheet.autoFilter?.sortState?.conditions[0]?.descending, false);

    const source = app['runtime'].model.getDataSource(region.sourceId);
    const loaded = await app['runtime'].dataContent.get(source.id)!.getRows(0, source.rowCount);
    assert.equal(loaded.state.availability, 'ready');
    assert.deepEqual(loaded.value?.map((row) => row[0]), ['a', 'b']);
  });

  it('anchors the AutoFilter sort context to the full region after selecting a filter column', async () => {
    const app = new WorkbookSession();
    await app.loadQuery(createInlineJsonQuery('filter-sort-column-query', 'Filter sort column', [
      { Key: 'b', Value: 2 },
      { Key: 'a', Value: 1 },
    ]));
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    const region = sheet.dataRegions[0]!;
    const dataRegionContext = {
      ...app.getDataRegionContext(),
      range: structuredClone(region.range),
      currentRegion: structuredClone(region.range),
    };
    const filterResult = await app.dispatch({
      commandId: 'sheet.autoFilter.toggle',
      params: { sheetId, range: region.range, dataRegionContext },
    });
    assert.equal(filterResult.status, 'committed');

    app.selectCell('B1');
    app.selectActiveColumn();
    assert.notDeepEqual(app.getDataRegionContext().range, region.range);
    const previousSourceRevision = app['runtime'].model.getDataSource(region.sourceId).revision;
    app.sortFilterColumn(region.range.startColumn + 1, true);
    for (let attempt = 0; attempt < 50; attempt += 1) {
      if (app['runtime'].model.getDataSource(region.sourceId).revision !== previousSourceRevision) break;
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    assert.equal(sheet.autoFilter?.sortState?.conditions[0]?.descending, false);
    const source = app['runtime'].model.getDataSource(region.sourceId);
    const loaded = await app['runtime'].dataContent.get(source.id)!.getRows(0, source.rowCount);
    assert.deepEqual(loaded.value?.map((row) => row[1]), [1, 2]);
  });

  it('tests json connector configuration', async () => {
    const app = new WorkbookSession();
    const result = await app.testQueryConnection('json', {
      data: [{ A: 1 }],
    });
    assert.equal(result.ok, true);
  });

  it('blocks query.load for viewers', async () => {
    const app = new WorkbookSession();
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);
    await assert.rejects(() => app.loadQuery(createInlineJsonQuery('blocked', 'Blocked', [{ A: 1 }])));
    assert.equal(app.getUiSnapshot().lastQueryResult, null);
    assert.match(app.getUiSnapshot().notice, /permission|viewer|query/i);
  });
});
