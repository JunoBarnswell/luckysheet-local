import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  createPivotSourceIndex,
  estimatePivotSourceIndexBytes,
  pivotSourceColumnValues,
  pivotSourceIndexTransferables,
  pivotSourceRowPaths,
} from './source-index';

describe('PivotSourceIndex', () => {
  it('preserves OOXML text semantics while encoding repeated strings once', () => {
    const index = createPivotSourceIndex({
      columns: [
        { field: { fieldId: 'text-number', name: 'OCR Number', ordinal: 0 }, values: ['4.60', '4.60', '260'] },
        { field: { fieldId: 'number', name: 'Typed Number', ordinal: 1 }, values: [4.6, 4.6, 260] },
      ],
      rowPaths: [[{ sheetId: 'sheet-1', row: 1 }], [{ sheetId: 'sheet-1', row: 2 }], [{ sheetId: 'sheet-1', row: 3 }]],
    });

    assert.equal(index.fields[0]!.dataType, 'text');
    assert.equal(index.columns[0]!.kind, 'dictionary');
    if (index.columns[0]!.kind === 'dictionary') assert.deepEqual(index.columns[0]!.dictionary, ['4.60', '260']);
    assert.deepEqual(pivotSourceColumnValues(index, 0), ['4.60', '4.60', '260']);
    assert.equal(index.fields[1]!.dataType, 'number');
    assert.equal(index.columns[1]!.kind, 'number');
    assert.deepEqual(pivotSourceColumnValues(index, 1), [4.6, 4.6, 260]);
    assert.deepEqual(pivotSourceRowPaths(index, 2), [{ sheetId: 'sheet-1', row: 3 }]);
  });

  it('rejects declared data-source types that do not match their values', () => {
    assert.throws(() => createPivotSourceIndex({
      columns: [{ field: { fieldId: 'amount', name: 'Amount', ordinal: 0, dataType: 'number' }, values: [10, '20'] }],
      rowPaths: [[{ sheetId: 'sheet-1', row: 1 }], [{ sheetId: 'sheet-1', row: 2 }]],
    }), /incompatible with number/);
  });

  it('keeps attachment-scale buffers below thirty-five percent of the former row-object estimate', () => {
    const rowCount = 4_058;
    const columnCount = 23;
    const columns = Array.from({ length: columnCount }, (_, column) => ({
      field: { fieldId: `field:${column}`, name: `Field ${column}`, ordinal: column },
      values: Array.from({ length: rowCount }, (_, row) => `V${column}:${row % 32}`),
    }));
    const index = createPivotSourceIndex({
      columns,
      rowPaths: Array.from({ length: rowCount }, (_, row) => [{ sheetId: 'sheet-1', row }]),
    });
    const legacyRowObjectEstimate = rowCount * columnCount * 56 + rowCount * 48;
    assert.ok(estimatePivotSourceIndexBytes(index) < legacyRowObjectEstimate * 0.35);
    assert.equal(new Set(pivotSourceIndexTransferables(index)).size, columnCount + 1);
  });
});
