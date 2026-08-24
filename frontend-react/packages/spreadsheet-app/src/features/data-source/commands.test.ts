import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type DataSourceManifest, type SheetDataRegion } from '@react-sheets/core-model';
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

function region(): SheetDataRegion {
  return {
    id: 'region-1',
    sourceId: 'source-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    headerRow: 0,
    revision: 0,
  };
}

test('data-source commands mutate canonical metadata and undo as one transaction', () => {
  const workbook = new WorkbookModel('unit-1', 'Book');
  const runtime = new CommandRuntime(workbook);
  const manifest = registerDataSourceFeature(runtime);
  assert.deepEqual(manifest.commandIds, ['dataSource.add', 'dataSource.update', 'dataSource.remove', 'dataRegion.add', 'dataRegion.remove']);

  runtime.execute('dataSource.add', { sheetId: 'sheet-1', source: source() });
  runtime.execute('dataRegion.add', { sheetId: 'sheet-1', region: region() });
  assert.equal(workbook.dataModel.sources.size, 1);
  assert.equal(workbook.getSheet('sheet-1').dataRegions.length, 1);

  assert.equal(runtime.undo(), true);
  assert.equal(workbook.getSheet('sheet-1').dataRegions.length, 0);
  assert.equal(runtime.undo(), true);
  assert.equal(workbook.dataModel.sources.size, 0);
  assert.equal(runtime.redo(), true);
  assert.equal(runtime.redo(), true);
  assert.equal(workbook.dataModel.sources.size, 1);
  assert.equal(workbook.getSheet('sheet-1').dataRegions.length, 1);
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
