import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ServerPivotTaskPort, type ServerPivotRunRequest } from './server-task-port';

function request(taskId = 'pivot-task'): ServerPivotRunRequest {
  const definition = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-1',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
    target: { sheetId: 'sheet-1', anchor: { row: 5, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [
      { fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0 },
      { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 },
    ] },
    layout: {
      rows: [{ fieldId: 'region' }], columns: [], filters: [], allowMultipleFiltersPerField: true,
      collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const },
      values: [{ valueId: 'amount:sum', fieldId: 'amount', summarizeBy: 'sum' as const }],
      calculatedFields: [], calculatedItems: [], subtotalLocation: 'bottom' as const,
      showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const,
      expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
    },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  return {
    taskId,
    generation: 1,
    definition,
    controls: [],
    revisions: { pivotId: definition.id, sourceRevision: 'source-r1', layoutRevision: 'layout-r1', filterRevision: 'filter-r1' },
    source: definition.source.range,
    targetBounds: { rowCount: 100, columnCount: 20 },
  };
}

function result(taskId: string, revision = 7) {
  return {
    kind: 'pivot', revision, queryId: taskId, sourceRevision: revision, executionToken: `token:${taskId}`,
    rows: [{ rowId: 0, keys: ['East'], subtotal: false, grandTotal: false }],
    columns: [{ columnId: 0, keys: [] }],
    cells: [{ rowId: 0, columnId: 0, values: [10] }],
    totalGroups: 1, totalColumns: 1,
  };
}

describe('ServerPivotTaskPort', () => {
  it('prepares and executes a revision-pinned Rust analytics request', async () => {
    const calls: unknown[] = [];
    const api = {
      prepareAnalytics: async (_unitId: string, body: { queryId: string; revision: number; request: Readonly<Record<string, unknown>> }) => {
        calls.push(body);
        return { queryId: body.queryId, sourceRevision: body.revision, executionToken: `token:${body.queryId}`, expiresAt: new Date().toISOString() };
      },
      executeAnalytics: async (_unitId: string, taskId: string, body: { executionToken: string; request: Readonly<Record<string, unknown>> }) => {
        calls.push(body);
        return result(taskId);
      },
      cancelAnalytics: async () => undefined,
    };
    const port = new ServerPivotTaskPort({ unitId: 'unit-1', revision: () => 7, api });
    const tree = await port.run(request());
    assert.equal(tree.rows[0]?.label, 'East');
    assert.equal((calls[0] as { request: { kind: string; revision: number } }).request.kind, 'pivot');
    assert.equal((calls[0] as { request: { revision: number } }).request.revision, 7);
  });

  it('rejects a stale server response before publication', async () => {
    const api = {
      prepareAnalytics: async (_unitId: string, body: { queryId: string }) => ({ queryId: body.queryId, sourceRevision: 7, executionToken: `token:${body.queryId}`, expiresAt: new Date().toISOString() }),
      executeAnalytics: async (_unitId: string, taskId: string) => result(taskId, 6),
      cancelAnalytics: async () => undefined,
    };
    const port = new ServerPivotTaskPort({ unitId: 'unit-1', revision: () => 7, api });
    await assert.rejects(() => port.run(request('stale-task')), /PIVOT_TASK_REVISION_MISMATCH/);
  });

  it('cancels an in-flight request and rejects its late result', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const cancelled: string[] = [];
    const api = {
      prepareAnalytics: async (_unitId: string, body: { queryId: string }) => ({ queryId: body.queryId, sourceRevision: 7, executionToken: `token:${body.queryId}`, expiresAt: new Date().toISOString() }),
      executeAnalytics: async (_unitId: string, taskId: string) => { await gate; return result(taskId); },
      cancelAnalytics: async (_unitId: string, taskId: string) => { cancelled.push(taskId); },
    };
    const port = new ServerPivotTaskPort({ unitId: 'unit-1', revision: () => 7, api });
    const pending = port.run(request('cancel-task'));
    await Promise.resolve();
    port.cancel('cancel-task');
    release();
    await assert.rejects(() => pending, /PIVOT_TASK_CANCELLED/);
    assert.deepEqual(cancelled, ['cancel-task']);
  });
});
