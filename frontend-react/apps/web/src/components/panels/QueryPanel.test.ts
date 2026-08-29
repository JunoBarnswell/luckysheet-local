import assert from 'node:assert/strict';
import test from 'node:test';
import type { ConnectorManifest } from '@react-sheets/spreadsheet-app';
import { assertConnectorConfig, normalizeConnectorConfig } from './QueryPanel';

const jsonManifest: ConnectorManifest = {
  id: 'json', kind: 'json', execution: 'local', label: 'JSON records',
  fields: [{ key: 'data', label: 'JSON records', kind: 'multiline-text', required: true }],
};
const restManifest: ConnectorManifest = {
  id: 'rest', kind: 'rest', execution: 'server', label: 'REST source',
  fields: [
    { key: 'sourceRef', label: 'Source', kind: 'text', required: true },
    { key: 'statement', label: 'Path', kind: 'text', required: true },
    { key: 'method', label: 'Method', kind: 'select', required: true },
    { key: 'body', label: 'Body', kind: 'multiline-text', required: false },
  ],
};

test('connector config preserves the connector-specific protocol', () => {
  assert.doesNotThrow(() => assertConnectorConfig(jsonManifest, { data: '[{"A":1}]' }));
  assert.deepEqual(normalizeConnectorConfig(restManifest, {
    sourceRef: 'orders-api', statement: '/orders', method: 'POST', body: '{"limit":10}',
  }), {
    sourceRef: 'orders-api', statement: '/orders', method: 'POST', body: { limit: 10 },
  });
});

test('connector config rejects missing fields and malformed REST bodies', () => {
  assert.throws(() => assertConnectorConfig(jsonManifest, { data: '' }), /JSON records is required/);
  assert.throws(() => normalizeConnectorConfig(restManifest, {
    sourceRef: 'orders-api', statement: '/orders', method: 'POST', body: '{bad',
  }), /valid JSON/);
});
