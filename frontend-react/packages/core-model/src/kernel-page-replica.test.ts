import assert from 'node:assert/strict';
import { before, test } from 'node:test';
import { initializeNodeKernel } from '../../kernel-client/src/node';
import { kernelInvoke } from '../../kernel-client/src/index';
import { KernelPageReplica, WorksheetCells, type KernelReplicaManifest, type KernelReplicaPagePayload } from './kernel-page-replica';

before(async () => { await initializeNodeKernel(); });

function authoredWorkbook(unitId: string): { manifest: KernelReplicaManifest; pages: KernelReplicaPagePayload[] } {
  kernelInvoke('create', { unitId, name: 'Page boundary', sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 2048, columnCount: 64, metadata: { hiddenRows: [0] } }] });
  return kernelInvoke('command', { unitId, baseRevision: 0, operationId: `${unitId}:write`, commandId: 'operation.apply', accessRole: 'owner', params: { mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 42 } } }] } });
}

test('resident Rust pages preserve hidden-cell values and distinguish implicit blank space', () => {
  const unitId = 'page-replica-success';
  const committed = authoredWorkbook(unitId);
  const replica = new KernelPageReplica(unitId);
  replica.open(committed.manifest, committed.pages);
  const cells = new WorksheetCells(replica, 'sheet-1');
  assert.equal(cells.get(0, 0)?.value, 42);
  assert.equal(cells.get(0, 1), undefined);
  assert.equal(cells.get(1500, 50), undefined);
  assert.equal(cells.count(), 1);
  assert.deepEqual(cells.occupiedRange('sheet-1'), { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  assert.deepEqual(cells.currentRegion(0, 0), { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  kernelInvoke('close', { unitId });
});

test('unavailable manifest pages throw until proven bytes are loaded', async () => {
  const unitId = 'page-replica-missing';
  const committed = authoredWorkbook(unitId);
  kernelInvoke('close', { unitId });
  const replica = new KernelPageReplica(unitId);
  replica.open(committed.manifest);
  assert.throws(() => replica.readCell({ sheetId: 'sheet-1', row: 0, column: 0 }), { code: 'DATA_PAGE_UNAVAILABLE' });
  const cells = new WorksheetCells(replica, 'sheet-1');
  assert.equal(cells.count(), 1, 'directory statistics never require loading cell bytes');
  assert.throws(() => cells.currentRegion(0, 0), { code: 'DATA_PAGE_UNAVAILABLE' });
  let requests = 0;
  const transport = { getPage: async () => { requests += 1; return committed.pages[0]!; } };
  const range = { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  await Promise.all([replica.loadRange(range, transport), replica.loadRange(range, transport)]);
  assert.equal(cells.get(0, 0)?.value, 42);
  await replica.loadRange(range, transport);
  assert.equal(requests, 1, 'concurrent and subsequent reads share one resident page');
  kernelInvoke('close', { unitId });
});

test('corrupt committed page bytes cannot replace the prior replica', () => {
  const unitId = 'page-replica-corrupt';
  const committed = authoredWorkbook(unitId);
  const replica = new KernelPageReplica(unitId);
  replica.open(committed.manifest, committed.pages);
  const corrupt = { ...committed.pages[0]!, checksum: '0'.repeat(64) };
  assert.throws(() => replica.open(committed.manifest, [corrupt]));
  assert.equal(replica.readCell({ sheetId: 'sheet-1', row: 0, column: 0 })?.value, 42);
  kernelInvoke('close', { unitId });
});

test('revision advance retains committed resident pages and loads bounded replacements', async () => {
  const unitId = 'page-replica-revision-retention';
  kernelInvoke('create', { unitId, name: 'Revision retention', sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 2048, columnCount: 64, metadata: {} }] });
  const initial = kernelInvoke<{ manifest: KernelReplicaManifest; pages: KernelReplicaPagePayload[] }>('command', {
    unitId,
    baseRevision: 0,
    operationId: `${unitId}:initial`,
    commandId: 'operation.apply',
    accessRole: 'owner',
    params: { mutations: [
      { id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'first' } } },
      { id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 1024, column: 0, value: { value: 'second' } } },
    ] },
  });
  const replica = new KernelPageReplica(unitId);
  replica.open(initial.manifest, initial.pages);

  const firstChange = kernelInvoke<{ manifest: KernelReplicaManifest; pages: KernelReplicaPagePayload[] }>('command', {
    unitId,
    baseRevision: 1,
    operationId: `${unitId}:first-change`,
    commandId: 'operation.apply',
    accessRole: 'owner',
    params: { mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'first-updated' } } }] },
  });
  replica.open(firstChange.manifest, firstChange.pages);
  let unchangedRequests = 0;
  await replica.loadRange(
    { sheetId: 'sheet-1', startRow: 1024, endRow: 1024, startColumn: 0, endColumn: 0 },
    { getPage: async () => { unchangedRequests += 1; throw new Error('unchanged page must remain resident'); } },
  );
  assert.equal(unchangedRequests, 0);
  assert.equal(replica.readCell({ sheetId: 'sheet-1', row: 1024, column: 0 })?.value, 'second');

  const secondChange = kernelInvoke<{ manifest: KernelReplicaManifest; pages: KernelReplicaPagePayload[] }>('command', {
    unitId,
    baseRevision: 2,
    operationId: `${unitId}:second-change`,
    commandId: 'operation.apply',
    accessRole: 'owner',
    params: { mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 1024, column: 0, value: { value: 'second-updated' } } }] },
  });
  replica.open(secondChange.manifest, secondChange.pages);
  assert.equal(replica.readCell({ sheetId: 'sheet-1', row: 1024, column: 0 })?.value, 'second-updated');
  let changedRequests = 0;
  await replica.loadRange(
    { sheetId: 'sheet-1', startRow: 1024, endRow: 1024, startColumn: 0, endColumn: 0 },
    { getPage: async () => { changedRequests += 1; throw new Error('provided replacement must remain resident'); } },
  );
  assert.equal(changedRequests, 0);
  assert.equal(replica.readCell({ sheetId: 'sheet-1', row: 1024, column: 0 })?.value, 'second-updated');
  assert.equal(replica.readCell({ sheetId: 'sheet-1', row: 0, column: 0 })?.value, 'first-updated');
  kernelInvoke('close', { unitId });
});
