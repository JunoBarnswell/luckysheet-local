import assert from 'node:assert/strict';
import test from 'node:test';
import type {
  DataBlockRef,
  DataSourceManifest,
  PivotDefinition,
  TableScalar,
} from '@react-sheets/core-model';
import { LocalDataBlockStore } from '../persistence/data-block-store';
import { WorkspaceMemoryCoordinator } from '../persistence/memory';
import {
  computeColumnarBlockChecksum,
  encodeColumnarBlock,
  type ColumnarBlockField,
} from '../data-source/codec';
import { DataSourceContentQuery } from '../data-source/content-query';
import { readPivotBlockSource } from './block-source';
import { pivotSourceColumnValues, pivotSourceRowPaths } from './source-index';

const fields: ColumnarBlockField[] = [
  { id: 'orders:field:0', name: 'Region', ordinal: 0, type: 'text' },
  { id: 'orders:field:1', name: 'Amount', ordinal: 1, type: 'number' },
];

let sequence = 0;

function sourceId(): string {
  sequence += 1;
  return `pivot-block-source-${String(sequence)}`;
}

async function block(
  source: string,
  id: string,
  startRow: number,
  rows: readonly (readonly TableScalar[])[],
): Promise<{ ref: DataBlockRef; bytes: ArrayBuffer }> {
  const bytes = await encodeColumnarBlock({ fields, rows });
  const checksum = await computeColumnarBlockChecksum(bytes);
  return {
    bytes,
    ref: {
      id,
      dataSourceId: source,
      startRow,
      rowCount: rows.length,
      storageKey: `${source}/${id}`,
      checksum,
      byteLength: bytes.byteLength,
      encoding: 'columnar-v1',
      revision: 3,
    },
  };
}

function manifest(source: string, blocks: DataBlockRef[], rowCount: number): DataSourceManifest {
  return {
    schema: 'DataSourceManifest',
    version: 1,
    id: source,
    name: 'Orders',
    kind: 'chunked-table',
    sourceSheetId: 'source-sheet',
    rowCount,
    fields: fields.map((field) => ({ ...field })),
    blockRowCount: 65_536,
    blocks,
    revision: 3,
  };
}

function pivot(source: string): PivotDefinition {
  return {
    schema: 'PivotDefinition',
    id: 'pivot-orders',
    source: { kind: 'data-source', dataSourceId: source },
    target: { sheetId: 'pivot-sheet', anchor: { row: 0, column: 0 } },
    fieldCatalog: {
      schema: 'PivotFieldCatalog',
      fields: [
        { fieldId: fields[0]!.id, name: 'Region', dataType: 'text', ordinal: 0 },
        { fieldId: fields[1]!.id, name: 'Amount', dataType: 'number', ordinal: 1 },
      ],
    },
    layout: {
      rows: [],
      columns: [],
      filters: [],
      allowMultipleFiltersPerField: true,
      collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
      values: [],
      subtotalLocation: 'bottom',
      showRowGrandTotals: true,
      showColumnGrandTotals: true,
      reportLayout: 'outline',
    },
    refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
  };
}

test('reads a canonical data-source Pivot source with stable field ids and source row paths', async () => {
  const source = sourceId();
  const stored = await block(source, 'block-1', 0, [['East', 10], ['West', 20], ['East', 30], ['North', 40]]);
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  await store.put(stored.ref, stored.bytes);
  const query = new DataSourceContentQuery(manifest(source, [stored.ref], 4), store);
  const events: string[] = [];

  const result = await readPivotBlockSource(pivot(source), query, {
    sourceRowStart: 1,
    chunkRowCount: 2,
    onState: (state) => events.push(state.status),
  });

  assert.equal(result.status, 'ready');
  if (result.status !== 'ready') return;
  assert.equal(result.sourceRevision, 3);
  assert.deepEqual(result.source.fields, [
    { fieldId: 'orders:field:0', name: 'Region', ordinal: 0, dataType: 'text' },
    { fieldId: 'orders:field:1', name: 'Amount', ordinal: 1, dataType: 'number' },
  ]);
  assert.deepEqual(pivotSourceColumnValues(result.source, 0), ['East', 'West', 'East', 'North']);
  assert.deepEqual(pivotSourceColumnValues(result.source, 1), [10, 20, 30, 40]);
  assert.deepEqual(Array.from({ length: result.source.rowCount }, (_, row) => pivotSourceRowPaths(result.source, row)), [
    [{ sheetId: 'source-sheet', row: 1 }],
    [{ sheetId: 'source-sheet', row: 2 }],
    [{ sheetId: 'source-sheet', row: 3 }],
    [{ sheetId: 'source-sheet', row: 4 }],
  ]);
  assert.deepEqual(events, ['loading', 'ready']);
});

test('returns explicit missing state instead of an empty source when a block is unavailable', async () => {
  const source = sourceId();
  const missing = await block(source, 'missing-block', 0, [['East', 10]]);
  const query = new DataSourceContentQuery(manifest(source, [missing.ref], 1), new LocalDataBlockStore(new WorkspaceMemoryCoordinator()));
  const result = await readPivotBlockSource(pivot(source), query);

  assert.equal(result.status, 'missing');
  if (result.status === 'missing') {
    assert.match(result.error, /missing/i);
    assert.equal(result.state.blockId, 'missing-block');
  }
});

test('returns explicit error state for source identity mismatch and non-data-source Pivot sources', async () => {
  const querySource = sourceId();
  const otherSource = sourceId();
  const stored = await block(querySource, 'block-1', 0, [['East', 10]]);
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  await store.put(stored.ref, stored.bytes);
  const query = new DataSourceContentQuery(manifest(querySource, [stored.ref], 1), store);

  const mismatch = await readPivotBlockSource(pivot(otherSource), query);
  assert.equal(mismatch.status, 'error');
  if (mismatch.status === 'error') assert.match(mismatch.error, /does not match Pivot source/i);

  const worksheetPivot = { ...pivot(querySource), source: { kind: 'worksheet-range' as const, range: { sheetId: 'source-sheet', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } } };
  const unsupported = await readPivotBlockSource(worksheetPivot, query);
  assert.equal(unsupported.status, 'error');
  if (unsupported.status === 'error') assert.match(unsupported.error, /not a canonical data-source/i);
});

test('source row identity is required and never replaced with a fabricated empty result', async () => {
  const source = sourceId();
  const stored = await block(source, 'block-1', 0, [['East', 10]]);
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
  await store.put(stored.ref, stored.bytes);
  const sourceManifest = manifest(source, [stored.ref], 1);
  delete sourceManifest.sourceSheetId;
  const query = new DataSourceContentQuery(sourceManifest, store);
  const result = await readPivotBlockSource(pivot(source), query);
  assert.equal(result.status, 'error');
  if (result.status === 'error') assert.match(result.error, /worksheet identity/i);
});
