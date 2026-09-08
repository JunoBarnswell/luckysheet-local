import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exportSnapshotToOoxmlBase64 } from '@react-sheets/exchange-excel-ooxml';
import { CommandRegistry, type CommandContext } from '@react-sheets/command-runtime';
import {
  buildQueryResultSnapshot,
  buildQueryLoadPlan,
  createInlineJsonQuery,
  executeQueryDefinition,
  prepareQueryLoadPayload,
  resolveLoadTarget,
  summarizeQueryResult,
} from './runtime';
import { createDefaultConnectorRegistry, CsvDataConnector, deserializeQueryDefinition, serializeQueryDefinition, TsvDataConnector, OoxmlDataConnector } from './index';
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
    layout: { rows: [], columns: [], values: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const, expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true } },
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
    const connector = new OoxmlDataConnector();
    await connector.connect({ base64: exportSnapshotToOoxmlBase64(workbook.snapshot()) });
    assert.deepEqual(await connector.executeQuery('ignored'), { columns: ['Name', 'Units'], rows: [['Alpha', 3]], rowCount: 1 });
  });

  it('registers typed local and server connector manifests while rejecting server execution in the local pipeline', async () => {
    const registry = createDefaultConnectorRegistry();
    assert.deepEqual(registry.list().map((connector) => connector.id), ['json', 'csv', 'tsv', 'xlsx', 'rest', 'sqlite', 'jdbc']);
    assert.equal(registry.get('rest').manifest.execution, 'server');
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

  it('prepares block-backed query load metadata without embedding result rows', async () => {
    const model = new WorkbookModel('wb-query-blocks', 'Query');
    const query = createInlineJsonQuery('q-blocks', 'Blocks', [{ Name: 'Alpha', Qty: 2 }]);
    const prepared = await prepareQueryLoadPayload(model, query, { kind: 'range', sheetId: model.primarySheetId, range: { startRow: 0, startColumn: 0 } }, {
      columns: ['Name', 'Qty'], rows: [['Alpha', 2]], rowCount: 1,
    });
    assert.equal(prepared.payload.kind, 'data-source-load');
    assert.equal('result' in prepared.payload, false);
    assert.equal(prepared.blocks.length, 1);
    assert.equal(prepared.payload.source.blocks.length, 1);
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

  it('plans query loads as canonical mutations without mutating the replica', async () => {
    const model = new WorkbookModel('wb-query', 'Query');
    const registry = new CommandRegistry({ requireMutationMetadata: true });
    registerQueryCommands(registry);
    const sheetId = model.primarySheetId;
    const query = createInlineJsonQuery('q-load', 'Load', [{ Product: 'X', Units: 9 }]);
    const result = { columns: ['Product', 'Units'], rows: [['X', 9]], rowCount: 1 };
    const prepared = await prepareQueryLoadPayload(model, query, { kind: 'range', sheetId, range: { startRow: 0, startColumn: 0 } }, result);
    const mutations: unknown[] = [];
    const context = {
      workbook: model,
      operationId: 'op-load',
      executeCommand: () => ({ operationId: 'op-load', mutationCount: 0, affectedRanges: [] }),
      applyMutation: (mutation: unknown) => mutations.push(mutation),
      recordOperation: () => ({ operationId: 'op-load' }),
    } as unknown as CommandContext;
    const planned = registry.getCommand('query.load').execute(prepared.payload, context);
    assert.equal(planned.mutationCount, 1);
    assert.equal(mutations.length, 1);
    assert.equal((mutations[0] as { id: string }).id, 'query.load.range');
    assert.equal('inverse' in (mutations[0] as object), false);
    assert.equal(model.getSheet(sheetId).dataRegions.length, 0);
    assert.equal(model.getQueryDefinition('q-load'), undefined);
  });

  it('builds distinct load plans for sheet tables and pivots', async () => {
    const model = new WorkbookModel('wb-query-targets', 'Query');
    const sheet = model.getSheet(model.primarySheetId);
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Sales',
      range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
      showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'both', columns: [{ id: 'a', name: 'A' }, { id: 'b', name: 'B' }],
    });
    const pivot = queryTargetPivot(sheet.id);
    sheet.pivots.push(pivot);
    const query = createInlineJsonQuery('q-targets', 'Targets', [{ A: 1, B: 2 }]);
    const result = { columns: ['A', 'B'], rows: [[1, 2]], rowCount: 1 };
    const sheetTable = await prepareQueryLoadPayload(model, query, { kind: 'sheet-table', sheetId: sheet.id, tableId: 'table-1' }, result);
    const pivotSource = await prepareQueryLoadPayload(model, query, { kind: 'pivot-source', pivotId: 'pivot-1' }, result);
    assert.equal(buildQueryLoadPlan(model, sheetTable.payload).mutationId, 'query.load.sheet-table');
    assert.equal(buildQueryLoadPlan(model, pivotSource.payload).mutationId, 'query.load.pivot-source');
  });

  it('rejects query loads with a mismatched source identity before planning', async () => {
    const model = new WorkbookModel('wb-query-invalid', 'Query');
    const registry = new CommandRegistry({ requireMutationMetadata: true });
    registerQueryCommands(registry);
    const prepared = await prepareQueryLoadPayload(
      model,
      createInlineJsonQuery('q-invalid', 'Invalid', [{ A: 1 }]),
      { kind: 'range', sheetId: model.primarySheetId, range: { startRow: 0, startColumn: 0 } },
      { columns: ['A'], rows: [[1]], rowCount: 1 },
    );
    const mutations: unknown[] = [];
    const context = {
      workbook: model,
      operationId: 'op-invalid',
      executeCommand: () => ({ operationId: 'op-invalid', mutationCount: 0, affectedRanges: [] }),
      applyMutation: (mutation: unknown) => mutations.push(mutation),
      recordOperation: () => ({ operationId: 'op-invalid' }),
    } as unknown as CommandContext;
    assert.throws(
      () => registry.getCommand('query.load').execute({ ...prepared.payload, sourceId: 'wrong-source' }, context),
      /source identity is invalid/i,
    );
    assert.equal(mutations.length, 0);
  });
});
