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
import {
  applyDataRegionMaterialization,
  migrateDataRegionCellPatches,
  prepareDataRegionMaterialization,
  resolveCell,
  restoreDataRegionMaterialization,
  writeCellPatch,
} from './resolved-cell';
import { WorkbookModel } from '@react-sheets/core-model';

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

test('resolved cells preserve block values when legacy sparse metadata carries a stale value', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore();
  const block = await buildBlock(sourceId, 'resolved-block', 0, [['A', 10], ['B', 20]]);
  await store.put(block.ref, block.bytes);
  const query = new DataSourceContentQuery(manifest(sourceId, 2, [block.ref]), store);
  const workbook = new WorkbookModel('resolved-cell', 'Resolved Cell');
  workbook.addDataSource(query.manifest);
  const sheet = workbook.getSheet('sheet-1');
  sheet.rowCount = 4;
  sheet.columnCount = 2;
  sheet.dataRegions.push({
    id: 'resolved-region',
    sourceId,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    headerRow: 0,
    revision: 0,
  });

  // This is the legacy shape produced by the large-data import path.  Its
  // value is stale by design; migrate it once before entering the resolver.
  sheet.cells.set(1, 1, { value: 999, style: { bold: true } });
  assert.throws(
    () => resolveCell(sheet, 1, 1, new Map([[sourceId, query]])),
    /non-canonical cell overlay/,
  );
  assert.equal(migrateDataRegionCellPatches(sheet), 1);
  assert.equal(migrateDataRegionCellPatches(sheet), 0);
  await query.getRowValues(0);
  const loaded = resolveCell(sheet, 1, 1, new Map([[sourceId, query]]));
  assert.equal(loaded?.source, 'data-block-overlay');
  assert.equal(loaded?.base?.value, 10);
  assert.equal(loaded?.cell?.value, 10);
  assert.equal(loaded?.cell?.style?.bold, true);

  writeCellPatch(sheet, 1, 1, {
    schema: 'CellPatch',
    value: { kind: 'inherit' },
    style: { kind: 'set', value: { italic: true } },
  });
  const styled = resolveCell(sheet, 1, 1, new Map([[sourceId, query]]));
  assert.equal(styled?.cell?.value, 10);
  assert.deepEqual(styled?.cell?.style, { italic: true });

  writeCellPatch(sheet, 1, 1, {
    schema: 'CellPatch',
    value: { kind: 'set', value: 42 },
  });
  const changed = resolveCell(sheet, 1, 1, new Map([[sourceId, query]]));
  assert.equal(changed?.cell?.value, 42);
  assert.equal(changed?.cell?.style?.italic, true);

  writeCellPatch(sheet, 1, 1, {
    schema: 'CellPatch',
    style: { kind: 'clear' },
  });
  const cleared = resolveCell(sheet, 1, 1, new Map([[sourceId, query]]));
  assert.equal(cleared?.cell?.value, 42);
  assert.equal(cleared?.cell?.style, undefined);

  const restored = WorkbookModel.fromSnapshot(workbook.snapshot());
  const restoredCell = resolveCell(restored.getSheet(sheet.id), 1, 1, new Map([[sourceId, query]]));
  assert.equal(restoredCell?.cell?.value, 42);
  assert.equal(restoredCell?.cell?.style, undefined);

  const prepared = await prepareDataRegionMaterialization(workbook, sheet.id, 'resolved-region', new Map([[sourceId, query]]));
  assert.equal(workbook.getSheet(sheet.id).dataRegions.length, 1);
  assert.equal(workbook.dataSources.has(sourceId), true);
  const transaction = applyDataRegionMaterialization(workbook, prepared);
  assert.equal(transaction.sourceRemoved, true);
  assert.equal(workbook.getSheet(sheet.id).dataRegions.length, 0);
  assert.equal(workbook.dataSources.has(sourceId), false);
  assert.equal(workbook.getSheet(sheet.id).cells.get(1, 1)?.value, 42);
  restoreDataRegionMaterialization(workbook, transaction);
  assert.equal(workbook.getSheet(sheet.id).dataRegions.length, 1);
  assert.equal(workbook.dataSources.has(sourceId), true);
  const restoredTransactionCell = resolveCell(workbook.getSheet(sheet.id), 1, 1, new Map([[sourceId, query]]));
  assert.equal(restoredTransactionCell?.cell?.value, 42);
});

test('resolved cells expose loading and missing states without replacing a block with an empty cell', async () => {
  const sourceId = nextSourceId();
  const block = await buildBlock(sourceId, 'unloaded-block', 0, [['A', 10]]);
  const store = new LocalDataBlockStore();
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), store);
  const workbook = new WorkbookModel('resolved-unloaded', 'Resolved Unloaded');
  const sheet = workbook.getSheet('sheet-1');
  workbook.addDataSource(query.manifest);
  sheet.dataRegions.push({
    id: 'unloaded-region',
    sourceId,
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    headerRow: 0,
    revision: 0,
  });

  const loading = resolveCell(sheet, 1, 0, new Map([[sourceId, query]]));
  assert.equal(loading?.state?.availability, 'loading');
  assert.equal(loading?.cell?.value, 'Loading…');
  const missing = await query.getCellValue(0, 0);
  assert.equal(missing.state.availability, 'missing');
  const afterFailure = resolveCell(sheet, 1, 0, new Map([[sourceId, query]]));
  assert.equal(afterFailure?.state?.availability, 'missing');
  assert.equal(afterFailure?.cell?.value, '#BLOCK!');
  await assert.rejects(
    prepareDataRegionMaterialization(workbook, sheet.id, 'unloaded-region', new Map([[sourceId, query]])),
    /could not be fully loaded|missing from local storage/i,
  );
  assert.equal(sheet.dataRegions.length, 1);
});
