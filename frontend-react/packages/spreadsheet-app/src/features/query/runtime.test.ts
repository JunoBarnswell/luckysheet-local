import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import {
  buildQueryResultSnapshot,
  buildQueryLoadPlan,
  createInlineJsonQuery,
  executeQueryDefinition,
  InMemoryWorkbookTableQueryStore,
  queryResultToRangeValues,
  resolveLoadTarget,
  summarizeQueryResult,
} from './runtime';
import { createDefaultConnectorRegistry, deserializeQueryDefinition, serializeQueryDefinition } from './index';
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
    assert.equal(target.range?.endRow, 4);
  });

  it('fails closed for unsupported query steps', async () => {
    const query = createInlineJsonQuery('q-unsupported', 'Unsupported', [{ A: 1 }], [{
      id: 'custom-1', kind: 'custom', name: 'custom', config: {}, enabled: true,
    }]);
    await assert.rejects(() => executeQueryDefinition(createDefaultConnectorRegistry(), query), /not implemented/i);
  });

  it('persists definitions with redacted connector secrets and source revision', () => {
    const query = { ...createInlineJsonQuery('q-persist', 'Persist', [{ A: 1 }]), connectorId: 'rest', connectorConfig: { url: 'https://example.test', apiKey: 'secret', nested: { token: 'bearer' } }, sourceRevision: 7 };
    const persisted = serializeQueryDefinition(query);
    assert.equal(persisted.connectorConfig.apiKey, '[redacted]');
    assert.equal((persisted.connectorConfig.nested as Record<string, unknown>).token, '[redacted]');
    const restored = deserializeQueryDefinition(persisted, { apiKey: 'secret' });
    assert.equal(restored.connectorConfig.apiKey, 'secret');
    assert.equal(restored.sourceRevision, 7);
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

  it('builds distinct load plans for sheet tables and pivots', () => {
    const model = new WorkbookModel('wb-query-targets', 'Query');
    const sheet = model.getSheet(model.activeSheetId);
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Sales',
      range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
      showFilterButton: true, columns: [{ id: 'a', name: 'A' }, { id: 'b', name: 'B' }],
    });
    const pivot = { id: 'pivot-1', sheetId: sheet.id, sourceRange: { sheetId: sheet.id, startRow: 5, endRow: 7, startColumn: 0, endColumn: 1 }, layout: { rows: [], columns: [], values: [], filters: [], showSubtotals: true, showGrandTotals: true } } as never;
    sheet.pivots.push(pivot);
    const query = createInlineJsonQuery('q-targets', 'Targets', [{ A: 1, B: 2 }]);
    const result = { columns: ['A', 'B'], rows: [[1, 2]], rowCount: 1 };
    const store = { get: () => undefined, set: () => undefined, delete: () => undefined };
    assert.equal(buildQueryLoadPlan(model, { query, target: { kind: 'sheet-table', sheetId: sheet.id, tableId: 'table-1' }, result }, store).mutationId, 'query.load.sheet-table');
    assert.equal(buildQueryLoadPlan(model, { query, target: { kind: 'pivot-source', pivotId: 'pivot-1' }, result }, store).mutationId, 'query.load.pivot-source');
  });

  it('applies and reverts sheet-table, workbook-table, and pivot-source loads', () => {
    const model = new WorkbookModel('wb-query-replay', 'Query');
    const runtime = new CommandRuntime(model);
    registerSheetCommands(runtime);
    const tableStore = new InMemoryWorkbookTableQueryStore();
    registerQueryCommands(runtime.registry, { tableStore });
    const sheet = model.getSheet(model.activeSheetId);
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Sales',
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
      showFilterButton: true, columns: [{ id: 'a', name: 'A' }, { id: 'b', name: 'B' }],
    });
    model.addTable({ id: 'workbook-table-1', name: 'Results', rowCount: 0, fields: [], blockSize: 128, blocks: [], revision: 0 });
    const pivot = { id: 'pivot-1', sheetId: sheet.id, sourceRange: { sheetId: sheet.id, startRow: 5, endRow: 7, startColumn: 0, endColumn: 1 }, layout: { rows: [], columns: [], values: [], filters: [], showSubtotals: true, showGrandTotals: true } } as never;
    sheet.pivots.push(pivot);
    const query = createInlineJsonQuery('q-replay', 'Replay', [{ A: 1, B: 2 }]);
    const result = { columns: ['A', 'B'], rows: [[1, 2]], rowCount: 1 };

    runtime.execute('query.load', { query, target: { kind: 'sheet-table', sheetId: sheet.id, tableId: 'table-1' }, result });
    assert.equal(sheet.cells.get(1, 1)?.value, 2);
    assert.equal(runtime.undo(), true);
    assert.equal(sheet.cells.get(1, 1), undefined);

    runtime.execute('query.load', { query, target: { kind: 'workbook-table', tableId: 'workbook-table-1' }, result });
    assert.equal(model.getTable('workbook-table-1').rowCount, 1);
    assert.equal(tableStore.get('workbook-table-1')?.result.rowCount, 1);
    assert.equal(runtime.undo(), true);
    assert.equal(model.getTable('workbook-table-1').rowCount, 0);

    runtime.execute('query.load', { query, target: { kind: 'pivot-source', pivotId: 'pivot-1' }, result });
    assert.equal(sheet.cells.get(6, 1)?.value, 2);
    assert.equal(sheet.pivots[0]?.refreshRevision, 1);
    assert.equal(runtime.undo(), true);
    assert.equal(sheet.cells.get(6, 1), undefined);
  });
});
