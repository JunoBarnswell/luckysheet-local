import test from 'node:test';
import assert from 'node:assert/strict';
import {
  decodeMessage,
  decodeOperationMessageV2,
  encodeMessage,
  encodeOperationMessageV2,
  AuthenticationRequiredError,
  WorkbookApiClient,
  validateOperationEnvelopeV2,
} from './index';

test('collaboration messages round-trip through the V2 wire contract', () => {
  const message = { type: 'changeset.ack' as const, operationId: 'op-1', revision: 2 };
  assert.deepEqual(decodeMessage(encodeMessage(message)), message);
});

test('OperationEnvelopeV2 excludes client actor and affected ranges', () => {
  const envelope = {
    schema: 'OperationEnvelopeV2' as const,
    operationId: 'op-v2',
    unitId: 'unit-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [{ id: 'cell.write', sheetId: 'sheet-1', params: { row: 0, column: 0, value: 'ok' } }],
    createdAt: new Date().toISOString(),
  };
  assert.deepEqual(validateOperationEnvelopeV2(envelope), envelope);
  assert.throws(() => validateOperationEnvelopeV2({ ...envelope, actorId: 'spoofed' }), /server-owned/);
  assert.throws(() => validateOperationEnvelopeV2({
    ...envelope,
    mutations: [{ ...envelope.mutations[0], affectedRanges: [] }],
  }), /server-owned/);
});

test('V2 collaboration messages reject legacy actor-bearing presence and changesets', () => {
  const message = {
    type: 'presence.updated' as const,
    unitId: 'unit-1',
    state: { row: 1, column: 1 },
  };
  assert.deepEqual(decodeOperationMessageV2(encodeOperationMessageV2(message)), message);
  assert.throws(() => decodeOperationMessageV2(JSON.stringify({ ...message, actorId: 'spoofed' })), /server-owned/);
  assert.throws(() => decodeOperationMessageV2(JSON.stringify({
    type: 'changeset.submit',
    payload: {
      schema: 'CollaborationChangeSetV1',
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
      return new Response(JSON.stringify({ snapshot: { unitId: 'unit-1' }, revision: 0 }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    },
  });
  await api.getSnapshot('unit-1');
  assert.equal(new Headers(request?.headers).get('authorization'), 'Bearer token-123');
  await assert.rejects(() => new WorkbookApiClient().getSnapshot('unit-1'), AuthenticationRequiredError);
});
