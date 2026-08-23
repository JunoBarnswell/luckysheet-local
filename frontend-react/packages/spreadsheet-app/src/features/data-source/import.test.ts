import assert from 'node:assert/strict';
import test from 'node:test';
import type { CellData, SheetSnapshot } from '@react-sheets/core-model';
import { decodeColumnarBlock } from './codec';
import {
  countNonEmptyCells,
  encodeSheetDataRegion,
  qualifiesForDataSourceImport,
} from './import';

function makeSheet(
  rowCount: number,
  columnCount: number,
  cells: Record<string, Record<string, CellData>>,
): SheetSnapshot {
  return {
    id: 'sheet-source',
    name: 'Orders',
    rowCount,
    columnCount,
    cells,
    merges: [],
    freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
    pivots: [],
    sparklines: [],
    drawings: [],
    drawingPayloads: {},
  };
}

function put(
  cells: Record<string, Record<string, CellData>>,
  row: number,
  column: number,
  value: CellData,
): void {
  (cells[String(row)] ??= {})[String(column)] = value;
}

test('large data import is gated strictly above the non-empty cell threshold', async () => {
  assert.equal(qualifiesForDataSourceImport(100_000), false);
  assert.equal(qualifiesForDataSourceImport(100_001), true);

  const cells: Record<string, Record<string, CellData>> = {
    '0': { '0': { value: 'Name' }, '1': { value: 'Amount' } },
    '1': { '0': { value: 'Ada' }, '1': { value: 10 } },
  };
  const sheet = makeSheet(2, 2, cells);
  const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } as const;
  assert.equal(countNonEmptyCells(sheet, range), 4);
  assert.equal(await encodeSheetDataRegion({ sheet, range, sourceId: 'source-small' }), undefined);
});

test('headered sheet data is split into fixed-size blocks with stable fields and metadata overlays', async () => {
  const dataRowCount = 65_537;
  const cells: Record<string, Record<string, CellData>> = {};
  put(cells, 0, 0, {
    value: 'Name',
    formula: '="Name"',
    style: { bold: true, background: '#e2e8f0' },
    comment: { id: 'comment-header', author: 'user', text: 'Header', createdAt: '2026-08-24T00:00:00.000Z' },
  });
  put(cells, 0, 1, { value: 'Amount' });
  for (let row = 0; row < dataRowCount; row += 1) {
    const sheetRow = row + 1;
    put(cells, sheetRow, 0, {
      value: `Order ${row}`,
      ...(row === 0 ? {
        formula: '="Order "&ROW()',
        style: { italic: true },
        comment: { id: 'comment-data', author: 'user', text: 'Data', createdAt: '2026-08-24T00:00:00.000Z' },
      } : {}),
    });
    put(cells, sheetRow, 1, { value: row * 10 });
  }
  const sheet = makeSheet(dataRowCount + 1, 2, cells);
  const range = {
    sheetId: sheet.id,
    startRow: 0,
    endRow: dataRowCount,
    startColumn: 0,
    endColumn: 1,
  } as const;
  const input = { sheet, range, sourceId: 'source-orders', sourceName: 'Orders', regionId: 'region-orders' };

  const result = await encodeSheetDataRegion(input);
  assert.ok(result);
  assert.equal(result.nonEmptyCellCount, (dataRowCount + 1) * 2);
  assert.deepEqual(result.manifest.fields, [
    { id: 'source-orders:field:0', name: 'Name', ordinal: 0, type: 'text' },
    { id: 'source-orders:field:1', name: 'Amount', ordinal: 1, type: 'number' },
  ]);
  assert.deepEqual(result.manifest.sourceRange, range);
  assert.deepEqual(result.region, {
    id: 'region-orders',
    sourceId: 'source-orders',
    range,
    headerRow: 0,
    revision: 0,
  });
  assert.deepEqual(result.header, ['Name', 'Amount']);
  assert.equal(result.headerMetadata.cells.length, 1);
  assert.deepEqual(result.headerMetadata.cells[0], {
    row: 0,
    column: 0,
    formula: '="Name"',
    style: { bold: true, background: '#e2e8f0' },
    comment: { id: 'comment-header', author: 'user', text: 'Header', createdAt: '2026-08-24T00:00:00.000Z' },
  });

  assert.equal(result.blocks.length, 2);
  assert.deepEqual(result.blocks.map(({ ref }) => ({
    startRow: ref.startRow,
    rowCount: ref.rowCount,
    encoding: ref.encoding,
    byteLength: ref.byteLength,
  })), [
    { startRow: 0, rowCount: 65_536, encoding: 'columnar-v1', byteLength: result.blocks[0]!.payload.byteLength },
    { startRow: 65_536, rowCount: 1, encoding: 'columnar-v1', byteLength: result.blocks[1]!.payload.byteLength },
  ]);
  assert.equal(result.blocks[0]!.ref.checksum.length, 64);
  assert.equal(result.blocks[0]!.ref.byteLength, result.blocks[0]!.payload.byteLength);
  assert.deepEqual(result.blocks[0]!.metadata.cells[0], {
    row: 0,
    column: 0,
    formula: '="Order "&ROW()',
    style: { italic: true },
    comment: { id: 'comment-data', author: 'user', text: 'Data', createdAt: '2026-08-24T00:00:00.000Z' },
  });

  const decoded = await decodeColumnarBlock(result.blocks[0]!.payload, {
    expectedChecksum: result.blocks[0]!.ref.checksum,
    expectedRowCount: 65_536,
    expectedFields: result.manifest.fields,
  });
  assert.deepEqual(decoded.rows[0], ['Order 0', 0]);
  assert.deepEqual(decoded.rows[65_535], ['Order 65_535', 655_350]);
  const tail = await decodeColumnarBlock(result.blocks[1]!.payload, {
    expectedChecksum: result.blocks[1]!.ref.checksum,
    expectedRowCount: 1,
    expectedFields: result.manifest.fields,
  });
  assert.deepEqual(tail.rows[0], ['Order 65536', 655_360]);
});

test('field identities and block references are deterministic for the same snapshot', async () => {
  const cells: Record<string, Record<string, CellData>> = {
    '0': { '0': { value: 'Date' }, '1': { value: 'Value' } },
  };
  for (let row = 1; row <= 50_001; row += 1) {
    put(cells, row, 0, { value: 45_000 + row, numberFormat: 'yyyy-mm-dd' });
    put(cells, row, 1, { value: row });
  }
  const sheet = makeSheet(50_002, 2, cells);
  const range = { sheetId: sheet.id, startRow: 0, endRow: 50_001, startColumn: 0, endColumn: 1 } as const;
  const first = await encodeSheetDataRegion({ sheet, range, sourceId: 'source-deterministic' });
  const second = await encodeSheetDataRegion({ sheet, range, sourceId: 'source-deterministic' });
  assert.ok(first);
  assert.ok(second);
  assert.deepEqual(second.manifest.fields, first.manifest.fields);
  assert.deepEqual(second.blocks.map(({ ref }) => ref), first.blocks.map(({ ref }) => ref));
  assert.deepEqual(second.blocks.map(({ payload }) => Array.from(new Uint8Array(payload))), first.blocks.map(({ payload }) => Array.from(new Uint8Array(payload))));
  assert.equal(first.manifest.fields[0]!.type, 'date');
});
