import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  ServerPivotTaskPort,
  type ServerPivotTaskTransport,
  type ServerPivotSourceRegisterRequest,
  type ServerPivotPrepareRequest,
  type ServerPivotExecuteRequest,
} from './server-task-port';

function fixture() {
  const definition = {
    schema: 'PivotDefinition' as const,
    id: 'server-pivot',
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
  const revisions = { pivotId: definition.id, sourceRevision: 'source-r1', layoutRevision: 'layout-r1', filterRevision: 'filter-r1' };
  return { definition, revisions };
}

class FakeServerPivotTransport implements ServerPivotTaskTransport {
  cancelled: string[] = [];
  executeResult: Promise<Awaited<ReturnType<ServerPivotTaskTransport['execute']>>> | undefined;

  registerSource(request: ServerPivotSourceRegisterRequest): Promise<Awaited<ReturnType<ServerPivotTaskTransport['registerSource']>>> {
    return Promise.resolve({ status: 'accepted', taskId: request.taskId, generation: request.generation, revision: 7, sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision });
  }

  prepare(request: ServerPivotPrepareRequest): Promise<Awaited<ReturnType<ServerPivotTaskTransport['prepare']>>> {
    return Promise.resolve({ status: 'accepted', taskId: request.taskId, generation: request.generation, revision: request.revision, sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision, prepareToken: `prepared:${request.taskId}` });
  }

  execute(request: ServerPivotExecuteRequest): Promise<Awaited<ReturnType<ServerPivotTaskTransport['execute']>>> {
    if (this.executeResult) return this.executeResult;
    return Promise.resolve({ status: 'completed', taskId: request.taskId, generation: request.generation, revision: 7, sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision, result: { kind: 'pivot', revision: 7, rows: [{ rowId: 0, keys: ['East'], subtotal: false, grandTotal: false }], columns: [{ columnId: 0, keys: [] }], cells: [{ rowId: 0, columnId: 0, values: [10] }], totalGroups: 1, totalColumns: 1 } });
  }

  viewport = this.execute.bind(this);

  drilldown(request: Parameters<ServerPivotTaskTransport['drilldown']>[0]): ReturnType<ServerPivotTaskTransport['drilldown']> {
    return Promise.resolve({ status: 'completed', taskId: request.taskId, generation: request.generation, revision: 7, sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision, drilldown: { sourceRows: [1], values: [['East', 10]], columns: [0, 1], total: 1, offset: request.drilldown.offset } });
  }

  cancel(request: { taskId: string }): Promise<void> {
    this.cancelled.push(request.taskId);
    return Promise.resolve();
  }
}

describe('ServerPivotTaskPort', () => {
  it('runs source registration, prepare, execute and sparse result conversion at one revision', async () => {
    const { definition, revisions } = fixture();
    const transport = new FakeServerPivotTransport();
    const port = new ServerPivotTaskPort({ unitId: 'server-pivot-test', revision: () => 7, transport });
    const sourceRequest = {
      unitId: 'server-pivot-test', taskId: 'source-1', generation: 1, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision,
      source: definition.source, ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
    };
    assert.equal((await port.registerSource(sourceRequest)).status, 'accepted');
    const prepare = await port.prepare({ unitId: 'server-pivot-test', taskId: 'task-1', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, layoutRevision: revisions.layoutRevision, filterRevision: revisions.filterRevision, definition, controls: [], targetBounds: { rowCount: 100, columnCount: 20 } });
    const result = await port.execute({ unitId: 'server-pivot-test', taskId: 'task-1', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, layoutRevision: revisions.layoutRevision, filterRevision: revisions.filterRevision, prepareToken: prepare.prepareToken });
    assert.equal(result.result.rows[0]?.label, 'East');
    assert.equal(result.result.grandTotal, null);
    port.dispose();
  });

  it('rejects a server response pinned to a stale revision', async () => {
    const { definition, revisions } = fixture();
    const transport = new FakeServerPivotTransport();
    transport.execute = (request) => Promise.resolve({ status: 'completed', taskId: request.taskId, generation: request.generation, revision: 6, sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision, result: { kind: 'pivot', revision: 6, rows: [], columns: [], cells: [], totalGroups: 0, totalColumns: 0 } });
    const port = new ServerPivotTaskPort({ unitId: 'server-pivot-test', revision: () => 7, transport });
    const prepare = await port.prepare({ unitId: 'server-pivot-test', taskId: 'stale-task', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, layoutRevision: revisions.layoutRevision, filterRevision: revisions.filterRevision, definition, controls: [], targetBounds: { rowCount: 100, columnCount: 20 } });
    await assert.rejects(() => port.execute({ unitId: 'server-pivot-test', taskId: 'stale-task', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, layoutRevision: revisions.layoutRevision, filterRevision: revisions.filterRevision, prepareToken: prepare.prepareToken }), /PIVOT_TASK_REVISION_MISMATCH/);
    port.dispose();
  });

  it('cancels an in-flight task and rejects its late result', async () => {
    const { definition, revisions } = fixture();
    const transport = new FakeServerPivotTransport();
    let resolve!: (value: Awaited<ReturnType<ServerPivotTaskTransport['execute']>>) => void;
    transport.executeResult = new Promise((next) => { resolve = next; });
    const port = new ServerPivotTaskPort({ unitId: 'server-pivot-test', revision: () => 7, transport });
    const prepare = await port.prepare({ unitId: 'server-pivot-test', taskId: 'cancel-task', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, layoutRevision: revisions.layoutRevision, filterRevision: revisions.filterRevision, definition, controls: [], targetBounds: { rowCount: 100, columnCount: 20 } });
    const pending = port.execute({ unitId: 'server-pivot-test', taskId: 'cancel-task', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, layoutRevision: revisions.layoutRevision, filterRevision: revisions.filterRevision, prepareToken: prepare.prepareToken });
    port.cancel('cancel-task');
    resolve({ status: 'completed', taskId: 'cancel-task', generation: 1, revision: 7, sourceIdentity: 'source-1', sourceRevision: revisions.sourceRevision, result: { kind: 'pivot', revision: 7, rows: [], columns: [], cells: [], totalGroups: 0, totalColumns: 0 } });
    await assert.rejects(() => pending, /PIVOT_TASK_CANCELLED/);
    assert.deepEqual(transport.cancelled, ['cancel-task']);
    port.dispose();
  });
});
