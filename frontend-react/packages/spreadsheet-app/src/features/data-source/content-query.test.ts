import assert from 'node:assert/strict';
import test from 'node:test';
import {
  type DataBlockRef,
  type DataSourceManifest,
  type TableScalar,
} from '@react-sheets/core-model';
import { LocalDataBlockStore } from '../persistence/data-block-store';
import { WorkspaceMemoryCoordinator } from '../persistence/memory';
import {
  computeColumnarBlockChecksum,
  encodeColumnarBlock,
  type ColumnarBlockField,
} from './codec';
import {
  DataSourceContentQuery,
  type DataBlockReader,
} from './content-query';
import { dataSourceCellPatchIdentity, resolveCanonicalDataSourceRegion } from './canonical-region';
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

test('canonical data-source region rejects a reader with same revision but different metadata', async () => {
  const sourceId = nextSourceId();
  const stored = await buildBlock(sourceId, 'canonical-region-block', 0, [['A', 10], ['B', 20]]);
  const source: DataSourceManifest = {
    ...manifest(sourceId, 2, [stored.ref]),
    sourceSheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
  };
  const workbook = new WorkbookModel('canonical-region-workbook', 'Canonical region');
  const sheet = workbook.getSheet('sheet-1');
  workbook.addDataSource(source);
  sheet.addDataRegion({ id: 'canonical-region', sourceId, range: structuredClone(source.sourceRange!), headerRow: 0, revision: 0 });
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  await store.put(stored.ref, stored.bytes);
  const query = new DataSourceContentQuery(source, store);

  const canonical = resolveCanonicalDataSourceRegion(workbook, sourceId, query);
  assert.equal(canonical.region.id, 'canonical-region');
  writeCellPatch(sheet, 1, 1, { schema: 'CellPatch', value: { kind: 'set', value: 11 } });
  writeCellPatch(sheet, 2, 0, { schema: 'CellPatch', style: { kind: 'set', value: { bold: true } } });
  assert.deepEqual(dataSourceCellPatchIdentity(canonical).map(({ row, column }) => [row, column]), [[0, 1]]);

  const mismatched = new DataSourceContentQuery({ ...source, name: 'Different metadata' }, store);
  assert.throws(() => resolveCanonicalDataSourceRegion(workbook, sourceId, mismatched), /does not match/i);
});

test('content query reads blocks, publishes loading/ready, and applies block-local overlays', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
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

test('distinct field values stay unloaded until requested and fail closed at the member limit', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const first = await buildBlock(sourceId, 'members-1', 0, [['A', 10], ['B', 20]]);
  const second = await buildBlock(sourceId, 'members-2', 2, [['A', 30], [null, 40]]);
  await store.put(first.ref, first.bytes);
  await store.put(second.ref, second.bytes);
  let reads = 0;
  const reader: DataBlockReader = {
    get: async (ref) => {
      reads += 1;
      return store.get(ref);
    },
  };
  const source = manifest(sourceId, 4, [first.ref, second.ref]);
  source.rowOrder = [2, 0, 3, 1];
  const query = new DataSourceContentQuery(source, reader);

  assert.equal(reads, 0);
  const members = await query.getDistinctFieldValues('code');
  assert.deepEqual(members.value, ['A', null, 'B']);
  assert.equal(members.state.availability, 'ready');
  assert.equal(reads, 2);

  const resolved = await query.getDistinctFieldValues('code', undefined,
    (logicalRow, baseValue) => logicalRow === 1 ? 'Patched' : baseValue);
  assert.deepEqual(resolved.value, ['A', 'Patched', null, 'B']);

  const limited = await query.getDistinctFieldValues('code', 2);
  assert.equal(limited.value, undefined);
  assert.equal(limited.state.availability, 'error');
  assert.match(limited.state.error ?? '', /exceeds the 2 distinct-value limit/i);

  const invalidLimit = await query.getDistinctFieldValues('amount', 0);
  assert.equal(invalidLimit.value, undefined);
  assert.equal(invalidLimit.state.availability, 'error');
  assert.match(invalidLimit.state.error ?? '', /positive safe integer/i);
});

test('ensures every block is readable without returning a copied full-range matrix', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const first = await buildBlock(sourceId, 'ensure-1', 0, [['A', 10], ['B', 20]]);
  const second = await buildBlock(sourceId, 'ensure-2', 2, [['C', 30], ['D', 40]]);
  await store.put(first.ref, first.bytes);
  await store.put(second.ref, second.bytes);
  let reads = 0;
  const reader: DataBlockReader = {
    get: async (ref) => {
      reads += 1;
      return store.get(ref);
    },
  };
  const query = new DataSourceContentQuery(manifest(sourceId, 4, [first.ref, second.ref]), reader);

  const loaded = await query.ensureAllBlocksLoaded();
  assert.equal(loaded.availability, 'ready');
  assert.equal(reads, 2);
  assert.deepEqual((await query.getCellValue(3, 'code')).value, 'D');
  assert.equal(reads, 2);
});

test('returns cached block row views without copying the decoded row arrays', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const first = await buildBlock(sourceId, 'view-1', 0, [['A', 10], ['B', 20]]);
  const second = await buildBlock(sourceId, 'view-2', 2, [['C', 30], ['D', 40]]);
  await store.put(first.ref, first.bytes);
  await store.put(second.ref, second.bytes);
  const query = new DataSourceContentQuery(manifest(sourceId, 4, [first.ref, second.ref]), store);

  const firstView = await query.getAllBlockRows();
  const secondView = await query.getAllBlockRows();
  assert.equal(firstView.state.availability, 'ready');
  assert.equal(firstView.value?.length, 2);
  assert.strictEqual(firstView.value?.[0]?.rows, secondView.value?.[0]?.rows);
  assert.deepEqual(firstView.value?.[1]?.rows[1], ['D', 40]);
  const copied = await query.getRows(0, 1);
  assert.notStrictEqual(copied.value?.[0], firstView.value?.[0]?.rows[0]);
});

test('maps logical rows through a virtual sort order without changing block storage', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const first = await buildBlock(sourceId, 'order-1', 0, [['A', 10], ['B', 20]]);
  const second = await buildBlock(sourceId, 'order-2', 2, [['C', 30], ['D', 40]]);
  await store.put(first.ref, first.bytes);
  await store.put(second.ref, second.bytes);
  const query = new DataSourceContentQuery({
    ...manifest(sourceId, 4, [first.ref, second.ref]),
    rowOrder: [3, 0, 2, 1],
  }, store);

  assert.deepEqual((await query.getRows(0, 4)).value, [['D', 40], ['A', 10], ['C', 30], ['B', 20]]);
  assert.equal((await query.getCellValue(0, 'code')).value, 'D');
  assert.equal((await query.getCellValue(3, 'amount')).value, 20);
  assert.equal((await query.getLoadState(first.ref.id))?.availability, 'ready');
  assert.equal((await query.getLoadState(second.ref.id))?.availability, 'ready');
});

test('concurrent requests share one block read and cross block reads preserve row order', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
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
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const block = await buildBlock(sourceId, 'missing-block', 0, [['A', 1]]);
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), store);

  const result = await query.getRowValues(0);
  assert.equal(result.value, undefined);
  assert.equal(result.state.availability, 'missing');
  assert.match(result.state.error ?? '', /missing from local storage/i);
  assert.equal(query.getLoadState(block.ref.id)?.availability, 'missing');
});

test('render reads retain failures until an explicit read retries the block', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const block = await buildBlock(sourceId, 'retry-block', 0, [['A', 1]]);
  let reads = 0;
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), {
    get: async (ref) => { reads += 1; return store.get(ref); },
  });
  assert.equal((await query.getRowValues(0)).state.availability, 'missing');
  for (let index = 0; index < 5; index += 1) {
    assert.equal(query.peekCellValue(0, 0).state.availability, 'missing');
    query.prefetchRows(0, 1);
  }
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(reads, 1);
  await store.put(block.ref, block.bytes);
  assert.deepEqual((await query.getRowValues(0)).value, ['A', 1]);
  assert.equal(reads, 2);
  assert.equal(query.peekCellValue(0, 1).value, 1);
});

test('a loading subscriber can reenter the reader without starting another request', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const block = await buildBlock(sourceId, 'reentrant-block', 0, [['A', 1]]);
  await store.put(block.ref, block.bytes);
  let reads = 0;
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), {
    get: async (ref) => { reads += 1; return store.get(ref); },
  });
  let nested: ReturnType<DataSourceContentQuery['getRowValues']> | undefined;
  const unsubscribe = query.subscribe((state) => {
    if (state.availability === 'loading') nested = query.getRowValues(0);
  });
  try {
    assert.deepEqual((await query.getRowValues(0)).value, ['A', 1]);
    assert.ok(nested);
    assert.deepEqual((await nested).value, ['A', 1]);
    assert.equal(reads, 1);
  } finally { unsubscribe(); }
});

test('invalid stored byte length and uncovered rows return explicit errors without empty data', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const block = await buildBlock(sourceId, 'bad-length-block', 0, [['A', 1]]);
  await store.put(block.ref, block.bytes);
  const badRef = { ...block.ref, byteLength: block.ref.byteLength + 1 };
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [badRef]), store);
  const invalid = await query.getRowValues(0);
  assert.equal(invalid.value, undefined);
  assert.equal(invalid.state.availability, 'error');
  assert.match(invalid.state.error ?? '', /byteLength/i);

  assert.throws(
    () => new DataSourceContentQuery(manifest(sourceId, 2, [block.ref]), store),
    /contiguous source coverage|complete source rowCount/i,
  );
});

test('invalid ranges and fields are errors, while empty ranges are ready and empty', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
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

test('one-time overlay migration preserves block values before canonical resolution', async () => {
  const sourceId = nextSourceId();
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const block = await buildBlock(sourceId, 'resolved-block', 0, [['A', 10], ['B', 20]]);
  await store.put(block.ref, block.bytes);
  const query = new DataSourceContentQuery(manifest(sourceId, 2, [block.ref]), store);
  const workbook = new WorkbookModel('resolved-cell', 'Resolved Cell');
  workbook.addDataSource(query.manifest);
  const sheet = workbook.getSheet('sheet-1');
  sheet.rowCount = 4;
  sheet.columnCount = 2;
  sheet.addDataRegion({
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
  assert.equal(workbook.dataModel.sources.has(sourceId), true);
  const transaction = applyDataRegionMaterialization(workbook, prepared);
  assert.equal(transaction.sourceRemoved, true);
  assert.equal(workbook.getSheet(sheet.id).dataRegions.length, 0);
  assert.equal(workbook.dataModel.sources.has(sourceId), false);
  assert.equal(workbook.getSheet(sheet.id).cells.get(1, 1)?.value, 42);
  restoreDataRegionMaterialization(workbook, transaction);
  assert.equal(workbook.getSheet(sheet.id).dataRegions.length, 1);
  assert.equal(workbook.dataModel.sources.has(sourceId), true);
  const restoredTransactionCell = resolveCell(workbook.getSheet(sheet.id), 1, 1, new Map([[sourceId, query]]));
  assert.equal(restoredTransactionCell?.cell?.value, 42);
});

test('resolved cells expose loading and missing states without replacing a block with an empty cell', async () => {
  const sourceId = nextSourceId();
  const block = await buildBlock(sourceId, 'unloaded-block', 0, [['A', 10]]);
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  const query = new DataSourceContentQuery(manifest(sourceId, 1, [block.ref]), store);
  const workbook = new WorkbookModel('resolved-unloaded', 'Resolved Unloaded');
  const sheet = workbook.getSheet('sheet-1');
  workbook.addDataSource(query.manifest);
  sheet.addDataRegion({
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
