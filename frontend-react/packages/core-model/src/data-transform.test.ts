import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { applyRowPermutation, createRowPermutationPlan, type RangeRef, WorkbookModel } from './index';

function range(sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number): RangeRef {
  return { sheetId, startRow, endRow, startColumn, endColumn };
}

describe('canonical row permutation metadata plan', () => {
  it('moves only metadata whose exact cell is inside the sort rectangle', () => {
    const workbook = new WorkbookModel('permutation-metadata', 'Permutation metadata');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 30;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    sheet.review.setNote(0, 1, { id: 'inside-note', author: 'u', text: 'inside', createdAt: 'now', visible: true });
    sheet.review.setNote(0, 25, { id: 'outside-note', author: 'u', text: 'outside', createdAt: 'now', visible: true });
    sheet.hyperlinks.set('0:1', { id: 'inside-link', target: { kind: 'url', url: 'https://inside.invalid' } });
    sheet.hyperlinks.set('0:25', { id: 'outside-link', target: { kind: 'url', url: 'https://outside.invalid' } });
    sheet.review.addThread({ id: 'outside-comment', sheetId: sheet.id, row: 0, column: 25, author: 'u', text: 'outside', createdAt: 'now', replies: [] });
    sheet.drawings.push(
      { id: 'inside-drawing', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 1 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'inside' },
      { id: 'outside-drawing', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 25 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'outside' },
    );

    applyRowPermutation(sheet, createRowPermutationPlan(range(sheet.id, 0, 1, 0, 1), [1, 0]));

    assert.equal(sheet.cells.get(0, 0)?.value, 'second');
    assert.equal(sheet.review.getNoteAt(1, 1)?.id, 'inside-note');
    assert.equal(sheet.review.getNoteAt(0, 25)?.id, 'outside-note');
    assert.equal(sheet.hyperlinks.get('1:1')?.id, 'inside-link');
    assert.equal(sheet.hyperlinks.get('0:25')?.id, 'outside-link');
    assert.equal(sheet.review.getThreadsAt(0, 25)[0]?.row, 0);
    assert.equal(sheet.review.getThreadsAt(0, 25)[0]?.column, 25);
    assert.equal(sheet.drawings.find((item) => item.id === 'inside-drawing')?.anchor.row, 1);
    assert.equal(sheet.drawings.find((item) => item.id === 'outside-drawing')?.anchor.row, 0);
  });

  it('splits a non-contiguous conditional-format target instead of widening it', () => {
    const workbook = new WorkbookModel('permutation-segments', 'Permutation segments');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 8;
    sheet.conditionalFormats.push({ id: 'cf-1', sheetId: sheet.id, ranges: [range(sheet.id, 0, 1, 0, 1)], type: 'highlight', style: { bold: true } });

    applyRowPermutation(sheet, createRowPermutationPlan(range(sheet.id, 0, 3, 0, 1), [2, 0, 3, 1]));

    assert.deepEqual(sheet.conditionalFormats[0]?.ranges, [range(sheet.id, 1, 1, 0, 1), range(sheet.id, 3, 3, 0, 1)]);
  });

  it('rejects a single-range owner that cannot represent exact target segments atomically', () => {
    const workbook = new WorkbookModel('permutation-reject', 'Permutation reject');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 8;
    sheet.cells.set(0, 0, { value: 'a' });
    sheet.cells.set(1, 0, { value: 'b' });
    sheet.protectionRules.push({ id: 'protected', scope: 'range', sheetId: sheet.id, range: range(sheet.id, 0, 1, 0, 1), locked: true, allow: {} });

    assert.throws(() => applyRowPermutation(sheet, createRowPermutationPlan(range(sheet.id, 0, 3, 0, 1), [2, 0, 3, 1])), /cannot exactly remap protection/);
    assert.equal(sheet.cells.get(0, 0)?.value, 'a');
    assert.equal(sheet.cells.get(1, 0)?.value, 'b');
    assert.deepEqual(sheet.protectionRules[0]?.range, range(sheet.id, 0, 1, 0, 1));
  });
});
