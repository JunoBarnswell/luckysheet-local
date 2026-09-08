import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime, type CommandContext } from '@react-sheets/command-runtime';
import { WorkbookModel, type DataSourceManifest } from '@react-sheets/core-model';
import { validateOperationEnvelope } from '@react-sheets/protocol';
import { registerDataSourceFeature } from './commands';

function source(): DataSourceManifest {
  return {
    schema: 'DataSourceManifest',
    version: 1,
    id: 'source-1',
    name: 'Sales',
    kind: 'chunked-table',
    rowCount: 2,
    fields: [{ id: 'amount', name: 'Amount', ordinal: 0, type: 'number' }],
    blockRowCount: 65_536,
    blocks: [{
      id: 'block-1',
      dataSourceId: 'source-1',
      startRow: 0,
      rowCount: 2,
      storageKey: 'source-1/block-1',
      checksum: 'a'.repeat(64),
      byteLength: 4,
      encoding: 'columnar-v1',
      revision: 0,
    }],
    revision: 0,
  };
}

test('data-source commands plan typed mutations without mutating the replica', () => {
  const workbook = new WorkbookModel('unit-1', 'Book');
  const runtime = new CommandRuntime(workbook);
  const manifest = registerDataSourceFeature(runtime);
  assert.deepEqual(manifest.commandIds, ['dataSource.add', 'dataSource.update', 'dataSource.remove', 'dataRegion.add', 'dataRegion.remove']);

  const mutations: unknown[] = [];
  const context = {
    workbook,
    operationId: 'op-1',
    executeCommand: () => ({ operationId: 'op-1', mutationCount: 0, affectedRanges: [] }),
    applyMutation: (mutation: unknown) => mutations.push(mutation),
    recordOperation: () => ({ operationId: 'op-1' }),
  } as unknown as CommandContext;

  const result = runtime.registry.getCommand('dataSource.add').execute({ sheetId: 'sheet-1', source: source() }, context);
  assert.equal(result.mutationCount, 1);
  assert.equal(mutations.length, 1);
  const mutation = mutations[0] as { id: string; unitId: string; sheetId: string; params: { source: DataSourceManifest }; affectedRanges: unknown[] };
  assert.equal(mutation.id, 'dataSource.add');
  assert.equal(mutation.unitId, workbook.unitId);
  assert.equal(mutation.sheetId, 'sheet-1');
  assert.equal(mutation.params.source.id, 'source-1');
  assert.deepEqual(mutation.affectedRanges, []);
  assert.equal(workbook.dataModel.sources.size, 0);
});

test('data-source commands reject invalid manifests before planning a mutation', () => {
  const workbook = new WorkbookModel('unit-invalid', 'Book');
  const runtime = new CommandRuntime(workbook);
  registerDataSourceFeature(runtime);
  const mutations: unknown[] = [];
  const context = {
    workbook,
    operationId: 'op-invalid',
    executeCommand: () => ({ operationId: 'op-invalid', mutationCount: 0, affectedRanges: [] }),
    applyMutation: (mutation: unknown) => mutations.push(mutation),
    recordOperation: () => ({ operationId: 'op-invalid' }),
  } as unknown as CommandContext;

  assert.throws(
    () => runtime.registry.getCommand('dataSource.add').execute({ sheetId: 'sheet-1', source: { ...source(), blocks: [{ ...source().blocks[0]!, checksum: 'invalid' }] } }, context),
    /checksum/i,
  );
  assert.equal(mutations.length, 0);
});

test('operation validation rejects block bytes while accepting metadata', () => {
  const accepted = validateOperationEnvelope({
    schema: 'OperationEnvelope',
    operationId: 'op-1',
    unitId: 'unit-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [{ id: 'dataSource.add', sheetId: 'sheet-1', params: { source: source() } }],
    createdAt: '2026-08-24T00:00:00.000Z',
  });
  assert.equal(accepted.mutations.length, 1);
  const block = source().blocks[0]!;
  assert.throws(() => validateOperationEnvelope({
    ...accepted,
    mutations: [{
      id: 'dataSource.add',
      sheetId: 'sheet-1',
      params: { source: { ...source(), blocks: [{ ...block, bytes: 'not-wire-data' }] } },
    }],
  }), /not allowed|unsupported field/);
});
