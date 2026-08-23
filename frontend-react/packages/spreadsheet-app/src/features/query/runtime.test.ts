import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exportSnapshotToXlsxBase64 } from '@react-sheets/exchange-xlsx';
import { CommandRegistry, CommandRuntime } from '@react-sheets/command-runtime';
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
import { createDefaultConnectorRegistry, CsvDataConnector, deserializeQueryDefinition, RestDataConnector, serializeQueryDefinition, TsvDataConnector, XlsxDataConnector } from './index';
import { QueryStepPipeline } from './query-steps';
import { registerQueryCommands } from './commands';

function queryTargetPivot(sheetId: string) {
  return {
    schema: 'PivotDefinition' as const,
    id: 'pivot-1',
    source: { kind: 'worksheet-range' as const, range: { sheetId, startRow: 5, endRow: 7, startColumn: 0, endColumn: 1 } },
    target: { sheetId, anchor: { row: 10, column: 0 } },
    fieldCatalog: {
      schema: 'PivotFieldCatalog' as const,
      fields: [
        { fieldId: 'query:field:0', name: 'A', dataType: 'number' as const, ordinal: 0 },
        { fieldId: 'query:field:1', name: 'B', dataType: 'number' as const, ordinal: 1 },
      ],
    },
    layout: { rows: [], columns: [], values: [], filters: [], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false, expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true } },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
}

describe('query runtime', () => {
  it('executes local CSV and TSV connectors with quoted fields', async () => {
    const csv = new CsvDataConnector();
    await csv.connect({ text: 'Name,Note\nAlpha,"East, sales"\nBeta,"two""quotes"""' });
    assert.deepEqual(await csv.executeQuery('ignored'), {
      columns: ['Name', 'Note'],
      rows: [['Alpha', 'East, sales'], ['Beta', 'two"quotes"']],
      rowCount: 2,
    });
    const tsv = new TsvDataConnector();
    await tsv.connect({ text: 'Name\tUnits\nAlpha\t2' });
    assert.deepEqual(await tsv.executeQuery('ignored'), { columns: ['Name', 'Units'], rows: [['Alpha', 2]], rowCount: 1 });
  });

  it('executes XLSX connector through the existing OOXML package reader', async () => {
    const workbook = new WorkbookModel('query-xlsx', 'Query XLSX');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Name' });
    sheet.cells.set(0, 1, { value: 'Units' });
    sheet.cells.set(1, 0, { value: 'Alpha' });
    sheet.cells.set(1, 1, { value: 3 });
    const connector = new XlsxDataConnector();
    await connector.connect({ base64: exportSnapshotToXlsxBase64(workbook.snapshot()) });
    assert.deepEqual(await connector.executeQuery('ignored'), { columns: ['Name', 'Units'], rows: [['Alpha', 3]], rowCount: 1 });
  });

  it('registers only local connectors and rejects server-only execution', async () => {
    const registry = createDefaultConnectorRegistry();
    assert.deepEqual(registry.list().map((connector) => connector.id), ['json', 'csv', 'tsv', 'xlsx']);
    registry.register(new RestDataConnector());
    await assert.rejects(
      () => executeQueryDefinition(registry, { id: 'remote', name: 'Remote', connectorId: 'rest', connectorConfig: {}, steps: [] }),
      /server-only/i,
    );
    await assert.rejects(
      () => executeQueryDefinition(registry, { id: 'database', name: 'Database', connectorId: 'sqlite', connectorConfig: {}, steps: [] }),
      /server-only/i,
    );
  });

  it('applies filter, select, rename, sort, group, join and pivot steps without silent passthrough', () => {
    const filtered = new QueryStepPipeline([
      { id: 'filter', kind: 'filter', name: 'East', config: { column: 'Region', value: 'East' }, enabled: true },
      { id: 'select', kind: 'select-columns', name: 'Select', config: { columns: ['Region', 'Units'] }, enabled: true },
      { id: 'rename', kind: 'rename-column', name: 'Rename', config: { from: 'Units', to: 'Count' }, enabled: true },
      { id: 'sort', kind: 'sort', name: 'Sort', config: { column: 'Count', ascending: false }, enabled: true },
    ]).applySteps({ columns: ['Region', 'Units'], rows: [['East', 2], ['East', 7], ['West', 9]] });
    assert.deepEqual(filtered, { columns: ['Region', 'Count'], rows: [['East', 7], ['East', 2]] });

    const grouped = new QueryStepPipeline([{ id: 'group', kind: 'group-by', name: 'Group', config: { by: ['Region'], aggregations: [{ column: 'Units', function: 'sum', as: 'Total' }] }, enabled: true }]).applySteps({ columns: ['Region', 'Units'], rows: [['East', 2], ['East', 7], ['West', 9]] });
    assert.deepEqual(grouped, { columns: ['Region', 'Total'], rows: [['East', 9], ['West', 9]] });

    const joined = new QueryStepPipeline([{ id: 'join', kind: 'join', name: 'Join', config: { on: ['Region'], right: { columns: ['Region', 'Manager'], rows: [['East', 'A'], ['West', 'B']] } }, enabled: true }]).applySteps({ columns: ['Region', 'Units'], rows: [['East', 2], ['West', 9]] });
    assert.deepEqual(joined, { columns: ['Region', 'Units', 'Region_right', 'Manager'], rows: [['East', 2, 'East', 'A'], ['West', 9, 'West', 'B']] });

    const pivot = new QueryStepPipeline([{ id: 'pivot', kind: 'pivot', name: 'Pivot', config: { rows: ['Region'], columns: ['Quarter'], values: ['Units'], aggregation: 'sum' }, enabled: true }]).applySteps({ columns: ['Region', 'Quarter', 'Units'], rows: [['East', 'Q1', 2], ['East', 'Q2', 7], ['West', 'Q1', 9]] });
    assert.deepEqual(pivot, { columns: ['Region', 'Q1 · Units', 'Q2 · Units'], rows: [['East', 2, 7], ['West', 9, 0]] });
    assert.throws(() => new QueryStepPipeline([{ id: 'custom', kind: 'custom', name: 'Custom', config: {}, enabled: true }]).applySteps({ columns: ['A'], rows: [[1]] }), /not implemented/i);
  });

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
  it('exposes complete mutation contracts for definition and load replay', () => {
    const registry = new CommandRegistry({ requireMutationMetadata: true });
    registerQueryCommands(registry);
    registry.assertComplete();
  });

  it('loads query results into worksheet cells through query.load', () => {
    const model = new WorkbookModel('wb-query', 'Query');
    const runtime = new CommandRuntime(model);
    registerSheetCommands(runtime);
    registerQueryCommands(runtime.registry);
    const sheetId = model.primarySheetId;
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
    assert.equal(Array.isArray(model.getQueryDefinition('q-load')?.connectorConfig.data), true);
    assert.deepEqual(model.getQueryDefinition('q-load')?.steps, []);
    const restored = WorkbookModel.fromSnapshot(model.snapshot());
    assert.deepEqual(restored.getQueryDefinition('q-load'), model.getQueryDefinition('q-load'));
  });

  it('builds distinct load plans for sheet tables and pivots', () => {
    const model = new WorkbookModel('wb-query-targets', 'Query');
    const sheet = model.getSheet(model.primarySheetId);
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Sales',
      range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
      showFilterButton: true, columns: [{ id: 'a', name: 'A' }, { id: 'b', name: 'B' }],
    });
    const pivot = queryTargetPivot(sheet.id);
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
    const sheet = model.getSheet(model.primarySheetId);
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Sales',
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
      showFilterButton: true, columns: [{ id: 'a', name: 'A' }, { id: 'b', name: 'B' }],
    });
    model.addTable({ id: 'workbook-table-1', name: 'Results', rowCount: 0, fields: [], blockSize: 128, blocks: [], revision: 0 });
    const pivot = queryTargetPivot(sheet.id);
    sheet.pivots.push(pivot);
    const query = createInlineJsonQuery('q-replay', 'Replay', [{ A: 1, B: 2 }]);
    const result = { columns: ['A', 'B'], rows: [[1, 2]], rowCount: 1 };

    runtime.execute('query.load', { query, target: { kind: 'sheet-table', sheetId: sheet.id, tableId: 'table-1' }, result });
    assert.equal(sheet.cells.get(1, 1)?.value, 2);
    assert.equal(model.getQueryDefinition('q-replay')?.id, 'q-replay');
    assert.equal(runtime.undo(), true);
    assert.equal(sheet.cells.get(1, 1), undefined);
    assert.equal(model.getQueryDefinition('q-replay'), undefined);

    runtime.execute('query.load', { query, target: { kind: 'workbook-table', tableId: 'workbook-table-1' }, result });
    assert.equal(model.getTable('workbook-table-1').rowCount, 1);
    assert.equal(tableStore.get('workbook-table-1')?.result.rowCount, 1);
    assert.equal(runtime.undo(), true);
    assert.equal(model.getTable('workbook-table-1').rowCount, 0);

    runtime.execute('query.load', { query, target: { kind: 'pivot-source', pivotId: 'pivot-1' }, result });
    assert.equal(sheet.cells.get(6, 1)?.value, 2);
    assert.equal(runtime.undo(), true);
    assert.equal(sheet.cells.get(6, 1), undefined);
  });
});
