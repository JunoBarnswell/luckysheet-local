import test from 'node:test';
import assert from 'node:assert/strict';
import {
  decodeMessage,
  decodeOperationMessage,
  encodeMessage,
  encodeOperationMessage,
  AuthenticationRequiredError,
  WorkbookApiClient,
  validateHistoryRestoreRequest,
  validateOperationEnvelope,
  validateWorkbookSnapshot,
} from './index';

test('collaboration messages round-trip through the operation wire contract', () => {
  const message = { type: 'changeset.ack' as const, operationId: 'op-1', revision: 2 };
  assert.deepEqual(decodeMessage(encodeMessage(message)), message);
});

test('OperationEnvelope excludes client actor and affected ranges', () => {
  const envelope = {
    schema: 'OperationEnvelope' as const,
    operationId: 'op-1',
    unitId: 'unit-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [{ id: 'cell.write', sheetId: 'sheet-1', params: { row: 0, column: 0, value: 'ok' } }],
    createdAt: new Date().toISOString(),
  };
  assert.deepEqual(validateOperationEnvelope(envelope), envelope);
  assert.throws(() => validateOperationEnvelope({ ...envelope, actorId: 'spoofed' }), /server-owned/);
  assert.throws(() => validateOperationEnvelope({
    ...envelope,
    mutations: [{ ...envelope.mutations[0], affectedRanges: [] }],
  }), /server-owned/);
});

test('collaboration messages reject actor-bearing presence and legacy changesets', () => {
  const message = {
    type: 'presence.updated' as const,
    unitId: 'unit-1',
    state: { row: 1, column: 1 },
  };
  assert.deepEqual(decodeOperationMessage(encodeOperationMessage(message)), message);
  assert.throws(() => decodeOperationMessage(JSON.stringify({ ...message, actorId: 'spoofed' })), /server-owned/);
  assert.throws(() => decodeOperationMessage(JSON.stringify({
    type: 'changeset.submit',
    payload: {
      schema: 'CollaborationChangeSet',
      operationId: 'old',
      unitId: 'unit-1',
      actorId: 'spoofed',
      clientSequence: 1,
      baseRevision: 0,
      mutations: [],
      createdAt: new Date().toISOString(),
    },
  })), /Unsupported operation schema/);
});

test('WorkbookApiClient injects bearer authentication and fails closed without a provider', async () => {
  let request: RequestInit | undefined;
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'token-123',
    fetchImpl: async (_input, init) => {
      request = init;
      return new Response(JSON.stringify({
        snapshot: {
          schema: 'WorkbookSnapshot',
          unitId: 'unit-1',
          name: 'Workbook',
          sheets: [{
            id: 'sheet-1',
            name: 'Sheet1',
            rowCount: 100,
            columnCount: 26,
            cells: {},
            merges: [],
            freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
            pivots: [],
            sparklines: [],
            drawings: [],
            drawingPayloads: {},
          }],
        },
        revision: 0,
      }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    },
  });
  await api.getSnapshot('unit-1');
  assert.equal(new Headers(request?.headers).get('authorization'), 'Bearer token-123');
  await assert.rejects(() => new WorkbookApiClient().getSnapshot('unit-1'), AuthenticationRequiredError);
});

test('WorkbookApiClient uses a server-issued guest share token when no bearer exists', async () => {
  let request: RequestInit | undefined;
  const api = new WorkbookApiClient({
    shareTokenProvider: () => 'guest-token',
    fetchImpl: async (_input, init) => {
      request = init;
      return new Response(JSON.stringify({
        snapshot: {
          schema: 'WorkbookSnapshot',
          unitId: 'unit-guest',
          name: 'Guest workbook',
          sheets: [{
            id: 'sheet-1', name: 'Sheet1', rowCount: 10, columnCount: 10,
            cells: {}, merges: [], freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
            pivots: [], sparklines: [], drawings: [], drawingPayloads: {},
          }],
        },
        revision: 0,
      }), { status: 200, headers: { 'content-type': 'application/json' } });
    },
  });
  await api.getSnapshot('unit-guest');
  const headers = new Headers(request?.headers);
  assert.equal(headers.get('x-workbook-share-token'), 'guest-token');
  assert.equal(headers.has('authorization'), false);
});

test('WorkbookApiClient accepts access roles only from the server projection', async () => {
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'server-token',
    fetchImpl: async () => new Response(JSON.stringify({ unitId: 'unit-access', role: 'editor' }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    }),
  });
  assert.deepEqual(await api.getAccess('unit-access'), { unitId: 'unit-access', role: 'editor' });

  const malformed = new WorkbookApiClient({
    authTokenProvider: () => 'server-token',
    fetchImpl: async () => new Response(JSON.stringify({ unitId: 'unit-access', role: 'ownerish' }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    }),
  });
  await assert.rejects(() => malformed.getAccess('unit-access'), /invalid role/);
});

test('history restore request is target-revision-only and client API posts no snapshot', async () => {
  assert.deepEqual(validateHistoryRestoreRequest({ targetRevision: 3, reason: 'rollback' }), {
    targetRevision: 3,
    reason: 'rollback',
  });
  assert.throws(() => validateHistoryRestoreRequest({ targetRevision: 3, snapshot: {} }), /unsupported fields/);
  assert.throws(() => validateHistoryRestoreRequest({ targetRevision: 3, actorId: 'spoofed' }), /unsupported fields/);
  assert.throws(() => validateHistoryRestoreRequest({ targetRevision: -1 }), /non-negative/);

  let postedBody = '';
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'token-restore',
    fetchImpl: async (_input, init) => {
      postedBody = String(init?.body ?? '');
      return new Response(JSON.stringify({}), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    },
  });
  await api.restoreToRevision('unit-1', 3, 'rollback');
  assert.deepEqual(JSON.parse(postedBody), { targetRevision: 3, reason: 'rollback' });
  assert.equal(postedBody.includes('snapshot'), false);
  assert.equal('saveSnapshot' in api, false);
});

test('snapshot trust boundary rejects versioned or legacy drawing payloads', () => {
  assert.throws(() => validateWorkbookSnapshot({ schema: 'LegacyWorkbookSnapshot', unitId: 'unit-1' }), /Unsupported workbook snapshot schema/);
  assert.throws(() => validateWorkbookSnapshot({
    schema: 'WorkbookSnapshot',
    unitId: 'unit-1',
    name: 'Workbook',
    sheets: [{
      id: 'sheet-1',
      name: 'Sheet1',
      rowCount: 10,
      columnCount: 10,
      cells: {},
      merges: [],
      pivots: [],
      sparklines: [],
      charts: [],
      drawings: [],
      drawingPayloads: {},
    }],
  }), /legacy drawing collections/);
});
