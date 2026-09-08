import test, { before } from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel, type KernelReplicaManifest, type RangeRef } from '@react-sheets/core-model';
import { initializeNodeKernel } from '@react-sheets/kernel-client/node';
import {
  CommandRegistry,
  CommandRuntime,
  type KernelCommitRequest,
  type MutationInfo,
} from './index';

before(async () => initializeNodeKernel());

const range: RangeRef = {
  sheetId: 'sheet-1',
  startRow: 0,
  endRow: 0,
  startColumn: 0,
  endColumn: 0,
};

function manifest(unitId: string, revision: number): KernelReplicaManifest {
  return {
    schema: 'WorkbookManifest',
    version: 11,
    unitId,
    name: 'Runtime',
    revision,
    sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 1_048_576, columnCount: 16_384, metadata: {} }],
    pages: [],
    metadata: {},
  };
}

function runtimeFixture(unitId = 'unit-1') {
  const workbook = WorkbookModel.fromManifest(manifest(unitId, 0));
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation<{ row: number; column: number; value: string }>({
    id: 'cell.set',
    metadata: {
      schema: {
        name: 'CellSet',
        validate: (value): value is { row: number; column: number; value: string } => {
          if (!value || typeof value !== 'object') return false;
          const input = value as Record<string, unknown>;
          return Number.isInteger(input.row) && Number.isInteger(input.column) && typeof input.value === 'string';
        },
      },
      permission: { capability: 'sheet.cell.write' },
      affectedRanges: { resolve: () => [range], mode: 'exact' },
    },
  });
  runtime.registry.registerCommand<{ value: string }>({
    id: 'cell.set',
    execute: (params, context) => {
      context.applyMutation({
        id: 'cell.set',
        unitId,
        sheetId: 'sheet-1',
        params: { row: 0, column: 0, value: params.value },
        affectedRanges: [range],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [range] };
    },
  });
  const requests: KernelCommitRequest[] = [];
  runtime.setCommitPort(async (request) => {
    requests.push(structuredClone(request));
    return {
      operationId: request.operationId,
      baseRevision: request.baseRevision,
      revision: request.baseRevision + 1,
      manifest: manifest(unitId, request.baseRevision + 1),
      pages: [],
      removedPages: [],
      affectedRanges: [range],
    };
  });
  return { runtime, workbook, requests };
}

test('publishes a planned mutation only after the canonical commit acknowledgement', async () => {
  const { runtime, workbook, requests } = runtimeFixture();
  const observed: MutationInfo[] = [];
  runtime.onMutation((mutation) => observed.push(mutation));

  const result = await runtime.execute('cell.set', { value: 'A' });

  assert.equal(result.mutationCount, 1);
  assert.equal(workbook.revision, 1);
  assert.equal(requests.length, 1);
  assert.equal(requests[0]?.mutations[0]?.id, 'cell.set');
  assert.equal(requests[0]?.intent, undefined);
  assert.equal(observed.length, 1);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 1, redo: 0 });
});

test('undo and redo target immutable server history without client inverse mutations', async () => {
  const { runtime, workbook, requests } = runtimeFixture('unit-history');
  await runtime.execute('cell.set', { value: 'A' });
  const original = requests[0]!;

  assert.equal(await runtime.undo(), true);
  const undo = requests[1]!;
  assert.deepEqual(undo.mutations, []);
  assert.deepEqual(undo.intent, {
    type: 'undo',
    targetOperationId: original.operationId,
    targetBaseRevision: original.baseRevision,
  });
  assert.equal(workbook.revision, 2);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 1 });

  assert.equal(await runtime.redo(), true);
  const redo = requests[2]!;
  assert.deepEqual(redo.mutations, []);
  assert.deepEqual(redo.intent, {
    type: 'undo',
    targetOperationId: undo.operationId,
    targetBaseRevision: undo.baseRevision,
  });
  assert.equal(workbook.revision, 3);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 1, redo: 0 });
});

test('a rejected commit leaves the replica and history unchanged', async () => {
  const { runtime, workbook } = runtimeFixture('unit-rejected');
  runtime.setCommitPort(async () => { throw new Error('UNDO_CONFLICT'); });

  await assert.rejects(() => runtime.execute('cell.set', { value: 'A' }), /UNDO_CONFLICT/);
  assert.equal(workbook.revision, 0);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('registry rejects incomplete metadata and invalid affected ranges', () => {
  const registry = new CommandRegistry();
  assert.throws(() => registry.registerMutation({ id: 'missing.contract', metadata: undefined as never }), /requires canonical metadata/);
  registry.registerMutation({
    id: 'cell.set',
    metadata: {
      schema: { name: 'Params', validate: (value) => value !== null && typeof value === 'object' },
      permission: { capability: 'sheet.cell.write' },
      affectedRanges: { resolve: () => [range], mode: 'exact' },
    },
  });
  const issues = registry.validateMutation({
    id: 'cell.set',
    unitId: 'unit-1',
    sheetId: 'sheet-1',
    params: {},
    affectedRanges: [],
  });
  assert.equal(issues.some((entry) => entry.code === 'invalid-affected-ranges'), true);
});

test('history-free view commands do not create a server transaction', async () => {
  const { runtime, requests } = runtimeFixture('unit-view');
  runtime.registry.registerCommand({
    id: 'view.noop',
    history: 'none',
    execute: (_params: unknown, context) => ({ operationId: context.operationId, mutationCount: 0, affectedRanges: [] }),
  });

  const result = await runtime.execute('view.noop', {});
  assert.equal(result.mutationCount, 0);
  assert.equal(requests.length, 0);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});
