import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';
import { createInlineJsonQuery } from './query-bridge';

describe('SpreadsheetApplication query integration', () => {
  it('loads inline json queries into the active sheet', async () => {
    const app = new SpreadsheetApplication();
    const query = createInlineJsonQuery('demo-query', 'Demo', [
      { Region: 'East', Units: 12 },
      { Region: 'West', Units: 8 },
    ]);

    await app.loadQuery(query);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.lastQueryResult?.rowCount, 2);
    assert.equal(snapshot.loadedQueries.length, 1);
    assert.equal(snapshot.queryConnectors.includes('json'), true);

    const sheet = app.getWorkbook().getSheet(app.getActiveSheetId());
    assert.equal(sheet.cells.get(0, 0)?.value, 'Region');
    assert.equal(sheet.cells.get(1, 1)?.value, 12);
  });

  it('refreshes a loaded query through query.refresh', async () => {
    const app = new SpreadsheetApplication();
    const query = createInlineJsonQuery('refresh-query', 'Refresh', [{ Value: 1 }]);
    await app.loadQuery(query);
    await app.refreshQuery('refresh-query');
    assert.equal(app.getUiSnapshot().lastQueryResult?.rowCount, 1);
  });

  it('tests json connector configuration', async () => {
    const app = new SpreadsheetApplication();
    const result = await app.testQueryConnection('json', {
      data: [{ A: 1 }],
    });
    assert.equal(result.ok, true);
  });

  it('blocks query.load for viewers', async () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('viewer');
    await app.loadQuery(createInlineJsonQuery('blocked', 'Blocked', [{ A: 1 }]));
    assert.equal(app.getUiSnapshot().lastQueryResult, null);
    assert.match(app.getUiSnapshot().notice, /permission|viewer|query/i);
  });
});
