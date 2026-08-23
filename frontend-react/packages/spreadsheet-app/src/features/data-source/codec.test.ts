import assert from 'node:assert/strict';
import test from 'node:test';
import {
  COLUMNAR_BLOCK_ENCODING,
  COLUMNAR_BLOCK_FORMAT_VERSION,
  COLUMNAR_BLOCK_MAGIC,
  computeColumnarBlockChecksum,
  decodeColumnarBlock,
  encodeColumnarBlock,
  validateSparseCellOverlay,
  type ColumnarBlockField,
} from './codec';

const fields: ColumnarBlockField[] = [
  { id: 'code', name: 'Code', ordinal: 0, type: 'text' },
  { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' },
  { id: 'postedAt', name: 'Posted At', ordinal: 2, type: 'date' },
  { id: 'active', name: 'Active', ordinal: 3, type: 'boolean' },
];

test('columnar block round-trips typed columns, repeated text, and null validity', async () => {
  const rows = [
    ['A-1', 10.5, 45_200, true],
    [null, null, null, null],
    ['A-1', 0, 45_201, false],
    ['B-2', -3.25, 45_202, true],
  ] as const;
  const bytes = await encodeColumnarBlock({ fields, rows });
  assert.equal(new TextDecoder().decode(new Uint8Array(bytes, 0, 4)), COLUMNAR_BLOCK_MAGIC);
  assert.notEqual(new Uint8Array(bytes).byteLength, 0);
  const checksum = await computeColumnarBlockChecksum(bytes);
  const decoded = await decodeColumnarBlock(bytes, {
    expectedRowCount: rows.length,
    expectedFields: fields,
    expectedChecksum: checksum,
  });
  assert.equal(decoded.encoding, COLUMNAR_BLOCK_ENCODING);
  assert.equal(decoded.formatVersion, COLUMNAR_BLOCK_FORMAT_VERSION);
  assert.deepEqual(decoded.fields, fields);
  assert.deepEqual(decoded.rows, rows);
  assert.equal(decoded.checksum, checksum);
});

test('mixed fields retain explicit number, boolean, text, and null tags', async () => {
  const mixedFields: ColumnarBlockField[] = [{ id: 'value', name: 'Value', ordinal: 0, type: 'mixed' }];
  const rows = [[1], [true], ['text'], [null], [-4.5], [false], ['text']] as const;
  const bytes = await encodeColumnarBlock({ fields: mixedFields, rows });
  const decoded = await decodeColumnarBlock(bytes, { expectedFields: mixedFields });
  assert.deepEqual(decoded.rows, rows);
});

test('one block accepts the exact row limit and rejects a larger block', async () => {
  const blockRows = Array.from({ length: 65_536 }, (_, index) => [index] as const);
  const blockFields: ColumnarBlockField[] = [{ id: 'n', name: 'N', ordinal: 0, type: 'number' }];
  await assert.doesNotReject(() => encodeColumnarBlock({ fields: blockFields, rows: blockRows }));
  const tooManyRows = [...blockRows, [65_536] as const];
  await assert.rejects(
    () => encodeColumnarBlock({ fields: blockFields, rows: tooManyRows }),
    /row count cannot exceed/i,
  );
});

test('encoder rejects ragged rows, invalid values, and inconsistent field schema', async () => {
  const numberField: ColumnarBlockField[] = [{ id: 'n', name: 'N', ordinal: 0, type: 'number' }];
  await assert.rejects(() => encodeColumnarBlock({ fields: numberField, rows: [] }), /at least one row/i);
  await assert.rejects(() => encodeColumnarBlock({ fields: numberField, rows: [[1, 2]] }), /field count/i);
  await assert.rejects(() => encodeColumnarBlock({ fields: numberField, rows: [[Number.NaN]] }), /finite number/i);
  await assert.rejects(() => encodeColumnarBlock({ fields: numberField, rows: [['not a number']] }), /does not match number/i);
  await assert.rejects(() => encodeColumnarBlock({
    fields: [{ id: 'n', name: 'N', ordinal: 1, type: 'number' }],
    rows: [[1]],
  }), /ordinal 0/i);
});

test('decoder rejects corruption, truncation, row mismatch, and schema mismatch', async () => {
  const sourceRows = [['x', 1]] as const;
  const sourceFields: ColumnarBlockField[] = [
    { id: 'text', name: 'Text', ordinal: 0, type: 'text' },
    { id: 'number', name: 'Number', ordinal: 1, type: 'number' },
  ];
  const bytes = await encodeColumnarBlock({ fields: sourceFields, rows: sourceRows });
  const corrupted = bytes.slice(0);
  const corruptedBytes = new Uint8Array(corrupted);
  const checksumByteIndex = corrupted.byteLength - 33;
  corruptedBytes[checksumByteIndex] = corruptedBytes[checksumByteIndex]! ^ 0x01;
  await assert.rejects(() => decodeColumnarBlock(corrupted), /checksum/i);
  await assert.rejects(() => decodeColumnarBlock(bytes.slice(0, bytes.byteLength - 1)), /truncated|trailing|missing/i);
  await assert.rejects(() => decodeColumnarBlock(bytes, { expectedRowCount: 2 }), /row count/i);
  await assert.rejects(() => decodeColumnarBlock(bytes, {
    expectedFields: [{ id: 'other', name: 'Other', ordinal: 0, type: 'text' }],
  }), /field schema/i);
});

test('sparse overlay validates coordinates and remains a separate metadata shape', () => {
  const overlay = {
    schema: 'SparseCellOverlay',
    revision: 3,
    cells: [{ row: 1, column: 2, value: 'edited' }, { row: 4, column: 0, value: null }],
  } as const;
  assert.doesNotThrow(() => validateSparseCellOverlay(overlay, { rowCount: 5, columnCount: 3 }));
  assert.throws(() => validateSparseCellOverlay({
    ...overlay,
    cells: [...overlay.cells, { row: 1, column: 2, value: 'duplicate' }],
  }, { rowCount: 5, columnCount: 3 }), /duplicate sparse cell/i);
  assert.throws(() => validateSparseCellOverlay({
    ...overlay,
    cells: [{ row: 5, column: 0, value: 'outside' }],
  }, { rowCount: 5, columnCount: 3 }), /outside the block bounds/i);
});
