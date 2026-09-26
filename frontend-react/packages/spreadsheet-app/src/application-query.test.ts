import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';
import { createInlineJsonQuery } from './features/query';
import { QueryLoadError } from './features/query/query-load-error';
import { createRemoteReadySessionFixture } from './session-test-fixtures';

describe('WorkbookSession query integration', () => {
  it('rejects an overlapping load of the same query without losing the first result', async () => {
    const app = new WorkbookSession();
    try {
      const query = createInlineJsonQuery('single-query-flight', 'Single flight', [{ Value: 1 }]);
      const first = app.loadQuery(query);
      await assert.rejects(app.loadQuery(query), (error: unknown) => error instanceof QueryLoadError && error.code === 'QUERY_LOAD_IN_PROGRESS');
      await first;
      assert.equal((await app['runtime'].dataContent.get('query:single-query-flight')!.getCellValue(0, 0)).value, 1);
      await app.refreshQuery(query.id);
      assert.equal(app.getQuerySnapshot().lastResult?.sourceRevision, 1);
    } finally { app.dispose(); }
  });

  it('rejects a queued load when its canonical definition changes before execution completes', async () => {
    const app = new WorkbookSession();
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const occupied = app['runtime'].connectors.withConnector('json', async () => gate);
    try {
      const query = createInlineJsonQuery('stale-query-definition', 'Original', [{ Value: 1 }]);
      const pending = app.loadQuery(query);
      const changed = { ...query, name: 'Updated', connectorConfig: { data: [{ Value: 2 }] } };
      app.runCommand('query.definition.replace', { definition: changed });
      const rejected = assert.rejects(pending, (error: unknown) => error instanceof QueryLoadError && error.code === 'QUERY_LOAD_STALE');
      release();
      await occupied;
      await rejected;
      assert.equal(app['runtime'].model.getQueryDefinition(query.id)?.name, 'Updated');
      assert.equal(app['runtime'].model.dataModel.sources.has('query:stale-query-definition'), false);
      await app.refreshQuery(query.id);
      assert.equal((await app['runtime'].dataContent.get('query:stale-query-definition')!.getCellValue(0, 0)).value, 2);
    } finally { release(); await occupied; app.dispose(); }
  });

  it('does not commit a queued query after the session is disposed', async () => {
    const app = new WorkbookSession();
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const occupied = app['runtime'].connectors.withConnector('json', async () => gate);
    const query = createInlineJsonQuery('disposed-query', 'Disposed', [{ Value: 1 }]);
    const pending = app.loadQuery(query);
    app.dispose();
    const rejected = assert.rejects(pending, (error: unknown) => error instanceof QueryLoadError && error.code === 'QUERY_LOAD_CANCELLED');
    release();
    await occupied;
    await rejected;
    assert.equal(app['runtime'].model.dataModel.sources.has('query:disposed-query'), false);
  });

  it('refreshes one query without discarding another sheet data source reader', async () => {
    const app = createRemoteReadySessionFixture();
    try {
      const firstSheetId = app.getActiveSheetId();
      await app.loadQuery(createInlineJsonQuery('cache-first', 'First', [{ Value: 1 }]));
      const firstReader = app['runtime'].dataContent.get('query:cache-first')!;
      assert.equal((await firstReader.getCellValue(0, 0)).value, 1);
      app.runCommand('sheet.add', { id: 'cache-second-sheet', name: 'Second' });
      await app.loadQuery(createInlineJsonQuery('cache-second', 'Second', [{ Value: 2 }]), {
        kind: 'range', sheetId: 'cache-second-sheet', range: { startRow: 0, startColumn: 0 },
      });
      const secondReader = app['runtime'].dataContent.get('query:cache-second')!;
      assert.equal((await secondReader.getCellValue(0, 0)).value, 2);
      app.selectSheet(firstSheetId);
      await app.refreshQuery('cache-first');
      assert.notEqual(app['runtime'].dataContent.get('query:cache-first'), firstReader);
      assert.equal(app['runtime'].dataContent.get('query:cache-second'), secondReader);
      assert.equal(secondReader.peekCellValue(0, 0).value, 2);
    } finally { app.dispose(); }
  });

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
