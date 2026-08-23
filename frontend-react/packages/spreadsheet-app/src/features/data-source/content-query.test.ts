import assert from 'node:assert/strict';
import test from 'node:test';
import {
  type DataBlockRef,
  type DataSourceManifest,
  type TableScalar,
} from '@react-sheets/core-model';
import { LocalDataBlockStore } from '../persistence/data-block-store';
import {
  computeColumnarBlockChecksum,
  encodeColumnarBlock,
  type ColumnarBlockField,
} from './codec';
import {
  DataSourceContentQuery,
  type DataBlockReader,
} from './content-query';

const fields: ColumnarBlockField[] = [
  { id: 'code', name: 'Code', ordinal: 0, type: 'text' },
  { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' },
];

let sourceSequence = 0;

function nextSourceId(): string {
  sourceSequence += 1;
  return `content-source-${String(sourceSequence)}`;
}

async function buildBlock(
  sourceId: string,
  blockId: string,
  startRow: number,
  rows: readonly (readonly TableScalar[])[],
): Promise<{ ref: DataBlockRef; bytes: ArrayBuffer }> {
  const bytes = await encodeColumnarBlock({ fields, rows });
  const checksum = await computeColumnarBlockChecksum(bytes);
  return {
    bytes,
    ref: {
      id: blockId,
      dataSourceId: sourceId,
      startRow,
      rowCount: rows.length,
      storageKey: `${sourceId}/${blockId}`,
      checksum,
      byteLength: bytes.byteLength,
      encoding: 'columnar-v1',
      revision: 0,
    },
  };
}

function manifest(sourceId: string, rowCount: number, blocks: DataBlockRef[]): DataSourceManifest {
  return {
    schema: 'DataSourceManifest',
    version: 1,
    id: sourceId,
    name: 'Content source',
    kind: 'chunked-table',
    rowCount,
    fields: fields.map((field) => ({ ...field })),
    blockRowCount: 65_536,
    blocks,
    revision: 0,
  };
}

test('content query reads blocks, publishes loading/ready, and applies block-local overlays', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore();
  const block = await buildBlock(sourceId, 'block-1', 0, [['A', 10], ['B', 20], [null, 30]]);
  await store.put(block.ref, block.bytes);
  const query = new DataSourceContentQuery(
    manifest(sourceId, 3, [block.ref]),
    store,
    {
      overlays: new Map([[block.ref.id, {
        schema: 'SparseCellOverlay',
        revision: 1,
        cells: [{ row: 1, column: 1, value: 25 }, { row: 2, column: 0, value: 'C' }],
      }]]),
    },
  );
  const events: string[] = [];
  const unsubscribe = query.subscribe((state) => events.push(`${state.blockId}:${state.availability}`));

  const pending = query.getRowValues(0);
  assert.equal(query.getLoadState(block.ref.id)?.availability, 'loading');
  const first = await pending;
  assert.equal(first.state.availability, 'ready');
  assert.deepEqual(first.value, ['A', 10]);
  assert.deepEqual(await query.getRowValues(1).then((result) => result.value), ['B', 25]);
  assert.equal((await query.getCellValue(2, 'code')).value, 'C');
  assert.deepEqual((await query.getFieldValues('amount', 0, 3)).value, [10, 25, 30]);
  assert.deepEqual(query.getField('amount'), fields[1]);
  assert.deepEqual(events, [`${block.ref.id}:loading`, `${block.ref.id}:ready`]);
  unsubscribe();
});

test('concurrent requests share one block read and cross block reads preserve row order', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore();
  const first = await buildBlock(sourceId, 'block-1', 0, [['A', 10], ['B', 20]]);
  const second = await buildBlock(sourceId, 'block-2', 2, [['C', 30], ['D', 40]]);
  await store.put(first.ref, first.bytes);
  await store.put(second.ref, second.bytes);
  let reads = 0;
  const reader: DataBlockReader = {
    get: async (ref) => {
      reads += 1;
      return store.get(ref);
    },
  };
  const query = new DataSourceContentQuery(manifest(sourceId, 4, [first.ref, second.ref]), reader, {
    overlays: new Map([[second.ref.id, {
      schema: 'SparseCellOverlay',
      revision: 1,
      cells: [{ row: 0, column: 1, value: 35 }],
    }]]),
  });

  const concurrent = await Promise.all([query.getCellValue(0, 'amount'), query.getRowValues(1)]);
  assert.equal(concurrent[0]!.value, 10);
  assert.deepEqual(concurrent[1]!.value, ['B', 20]);
  assert.equal(reads, 1);
  const range = await query.getRows(0, 4);
  assert.equal(range.state.availability, 'ready');
  assert.deepEqual(range.value, [['A', 10], ['B', 20], ['C', 35], ['D', 40]]);
  assert.equal(reads, 2);
});

test('missing blocks return an explicit missing state and remain retryable', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore();
  const block = await buildBlock(sourceId, 'missing-block', 0, [['A', 1]]);
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), store);

  const result = await query.getRowValues(0);
  assert.equal(result.value, undefined);
  assert.equal(result.state.availability, 'missing');
  assert.match(result.state.error ?? '', /missing from local storage/i);
  assert.equal(query.getLoadState(block.ref.id)?.availability, 'missing');
});

test('invalid stored byte length and uncovered rows return explicit errors without empty data', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore();
  const block = await buildBlock(sourceId, 'bad-length-block', 0, [['A', 1]]);
  await store.put(block.ref, block.bytes);
  const badRef = { ...block.ref, byteLength: block.ref.byteLength + 1 };
  const query = new DataSourceContentQuery(manifest(sourceId, 2, [badRef]), store);
  const invalid = await query.getRowValues(0);
  assert.equal(invalid.value, undefined);
  assert.equal(invalid.state.availability, 'error');
  assert.match(invalid.state.error ?? '', /byteLength/i);

  const uncovered = new DataSourceContentQuery(manifest(sourceId, 2, [block.ref]), store);
  const missing = await uncovered.getRows(0, 2);
  assert.equal(missing.value, undefined);
  assert.equal(missing.state.availability, 'missing');
  assert.match(missing.state.error ?? '', /no data block covers/i);
});

test('invalid ranges and fields are errors, while empty ranges are ready and empty', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore();
  const block = await buildBlock(sourceId, 'query-validation-block', 0, [['A', 1]]);
  await store.put(block.ref, block.bytes);
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), store);

  const invalidRow = await query.getRowValues(-1);
  assert.equal(invalidRow.state.availability, 'error');
  const invalidField = await query.getCellValue(0, 'missing');
  assert.equal(invalidField.state.availability, 'error');
  const empty = await query.getRows(1, 0);
  assert.equal(empty.state.availability, 'ready');
  assert.deepEqual(empty.value, []);
});
