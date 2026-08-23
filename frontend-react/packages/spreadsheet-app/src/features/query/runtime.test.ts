import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import {
  buildQueryResultSnapshot,
  createInlineJsonQuery,
  executeQueryDefinition,
  queryResultToRangeValues,
  resolveLoadTarget,
  summarizeQueryResult,
} from './runtime';
import { createDefaultConnectorRegistry } from './index';
import { registerQueryCommands } from './commands';

describe('query runtime', () => {
  it('executes json connector queries with pipeline filters', async () => {
    const connectors = createDefaultConnectorRegistry();
    const query = createInlineJsonQuery('q-1', 'Sales', [
      { Region: 'East', Units: 10 },
      { Region: 'West', Units: 5 },
      { Region: 'East', Units: 3 },
    ], [{
      id: 'filter-east',
      kind: 'filter',
      name: 'East only',
      config: { column: 'Region', value: 'East' },
      enabled: true,
    }]);

    const result = await executeQueryDefinition(connectors, query);
    assert.equal(result.rowCount, 2);
    assert.deepEqual(result.columns, ['Region', 'Units']);
  });

  it('maps query results to sheet range values with headers', () => {
    const values = queryResultToRangeValues({
      columns: ['Name', 'Qty'],
      rows: [['Alpha', 2]],
      rowCount: 1,
    });
    assert.equal(values.length, 2);
    assert.equal(values[0]?.[0]?.value, 'Name');
    assert.equal(values[1]?.[1]?.value, 2);
  });

  it('builds query result snapshots', () => {
    const snapshot = buildQueryResultSnapshot(
      createInlineJsonQuery('q-2', 'Demo', [{ A: 1 }]),
      { columns: ['A'], rows: [[1]], rowCount: 1 },
      { kind: 'range', sheetId: 'sheet-1', range: { startRow: 0, startColumn: 0 } },
    );
    assert.match(summarizeQueryResult(snapshot), /Loaded 1 rows/i);
    assert.equal(snapshot.target.sheetId, 'sheet-1');
  });

  it('resolves load targets from selection', () => {
    const target = resolveLoadTarget('sheet-1', {
      sheetId: 'sheet-1',
      startRow: 2,
      endRow: 4,
      startColumn: 1,
      endColumn: 3,
    });
    assert.equal(target.range?.startRow, 2);
    assert.equal(target.range?.startColumn, 1);
  });
});

describe('query commands', () => {
  it('loads query results into worksheet cells through query.load', () => {
    const model = new WorkbookModel('wb-query', 'Query');
    const runtime = new CommandRuntime(model);
    registerSheetCommands(runtime);
    registerQueryCommands(runtime.registry);
    const sheetId = model.activeSheetId;
    const query = createInlineJsonQuery('q-load', 'Load', [{ Product: 'X', Units: 9 }]);
    const result = { columns: ['Product', 'Units'], rows: [['X', 9]], rowCount: 1 };

    runtime.execute('query.load', {
      query,
      target: { kind: 'range', sheetId, range: { startRow: 0, startColumn: 0 } },
      result,
    });

    const sheet = model.getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'Product');
    assert.equal(sheet.cells.get(1, 1)?.value, 9);
  });
});
