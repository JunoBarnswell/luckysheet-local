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
  assert.equal(replica.isRangeResident(range), false);
  assert.equal(replica.isRangeResident({ ...range, startRow: 1500, endRow: 1500, startColumn: 50, endColumn: 50 }), true, 'logical blank pages are resident without payload bytes');
  assert.equal(replica.isPageResident('sheet-1', 0, 0), false);
  await Promise.all([replica.loadRange(range, transport), replica.loadRange(range, transport)]);
  assert.equal(replica.isRangeResident(range), true);
  assert.equal(replica.isPageResident('sheet-1', 0, 0), true);
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

test('page loading is globally bounded to four requests per replica', async () => {
  const unitId = 'page-replica-request-queue';
  kernelInvoke('create', { unitId, name: 'Request queue', sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 6144, columnCount: 64, metadata: {} }] });
  const committed = kernelInvoke<{ manifest: KernelReplicaManifest; pages: KernelReplicaPagePayload[] }>('command', {
    unitId,
    baseRevision: 0,
    operationId: `${unitId}:write`,
    commandId: 'operation.apply',
    accessRole: 'owner',
    params: { mutations: [
      ...[0, 1024, 2048, 3072, 4096, 5120].flatMap((row) => [0, 32].map((column) => ({
        id: 'cell.set',
        sheetId: 'sheet-1',
        params: { sheetId: 'sheet-1', row, column, value: { value: `${row}:${column}` } },
      }))),
    ] },
  });
  const payloads = new Map(committed.pages.map((page) => [`${page.sheetId}:${page.pageRow}:${page.pageColumn}`, page]));
  kernelInvoke('close', { unitId });

  const replica = new KernelPageReplica(unitId);
  replica.open(committed.manifest);
  let active = 0;
  let maximumActive = 0;
  let requests = 0;
  const transport = {
    getPage: async (params: { unitId: string; revision: number; sheetId: string; pageRow: number; pageColumn: number }, options?: { signal?: AbortSignal }) => {
      assert.ok(options?.signal, 'page transport receives the revision abort signal');
      assert.equal(options.signal.aborted, false);
      requests += 1;
      active += 1;
      maximumActive = Math.max(maximumActive, active);
      await new Promise((resolve) => setTimeout(resolve, 5));
      active -= 1;
      return payloads.get(`${params.sheetId}:${params.pageRow}:${params.pageColumn}`)!;
    },
  };
  const range = { sheetId: 'sheet-1', startRow: 0, endRow: 6143, startColumn: 0, endColumn: 63 };
  await replica.loadRange(range, transport);
  assert.equal(requests, committed.manifest.pages.length);
  assert.ok(maximumActive <= 4, `expected at most four active page requests, got ${maximumActive}`);
  assert.equal(replica.isRangeResident(range), true);
  assert.equal(replica.isPageResident('sheet-1', 5, 1), true);
  kernelInvoke('close', { unitId });
});

test('a response from an old revision is rejected before kernel page load', async () => {
  const unitId = 'page-replica-stale-request';
  const initial = authoredWorkbook(unitId);
  const replica = new KernelPageReplica(unitId);
  replica.open(initial.manifest);
  const range = { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  let startedResolve!: () => void;
  let payloadResolve!: (payload: KernelReplicaPagePayload) => void;
  let requestSignal: AbortSignal | undefined;
  const started = new Promise<void>((resolve) => { startedResolve = resolve; });
  const transport = {
    getPage: async (_params: { unitId: string; revision: number; sheetId: string; pageRow: number; pageColumn: number }, options?: { signal?: AbortSignal }) => {
      requestSignal = options?.signal;
      startedResolve();
      return await new Promise<KernelReplicaPagePayload>((resolve) => { payloadResolve = resolve; });
    },
  };
  const pending = replica.loadRange(range, transport);
  await started;

  const next = kernelInvoke<{ manifest: KernelReplicaManifest; pages: KernelReplicaPagePayload[] }>('command', {
    unitId,
    baseRevision: 1,
    operationId: `${unitId}:next`,
    commandId: 'operation.apply',
    accessRole: 'owner',
    params: { mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 43 } } }] },
  });
  replica.open(next.manifest);
  assert.equal(requestSignal?.aborted, true, 'revision changes abort the previous request signal');
  payloadResolve(initial.pages[0]!);
  await assert.rejects(pending, { code: 'STALE_REVISION' });
  assert.equal(replica.isRangeResident(range), false);
  kernelInvoke('close', { unitId });
});

test('failed page fetches remain unavailable and can be retried', async () => {
  const unitId = 'page-replica-fetch-failure';
  const committed = authoredWorkbook(unitId);
  kernelInvoke('close', { unitId });
  const replica = new KernelPageReplica(unitId);
  replica.open(committed.manifest);
  const range = { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  let attempts = 0;
  const transport = {
    getPage: async () => {
      attempts += 1;
      if (attempts === 1) throw new Error('temporary page service failure');
      return committed.pages[0]!;
    },
  };
  await assert.rejects(replica.loadRange(range, transport), /temporary page service failure/);
  assert.equal(replica.isPageResident('sheet-1', 0, 0), false);
  await replica.loadRange(range, transport);
  assert.equal(attempts, 2);
  assert.equal(replica.isPageResident('sheet-1', 0, 0), true);
  kernelInvoke('close', { unitId });
});
