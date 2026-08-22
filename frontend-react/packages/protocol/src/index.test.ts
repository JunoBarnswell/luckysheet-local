import test from 'node:test';
import assert from 'node:assert/strict';
import { decodeMessage, encodeMessage } from './index';

test('collaboration messages round-trip through the v1 wire contract', () => {
  const message = { type: 'changeset.ack' as const, operationId: 'op-1', revision: 2 };
  assert.deepEqual(decodeMessage(encodeMessage(message)), message);
});
