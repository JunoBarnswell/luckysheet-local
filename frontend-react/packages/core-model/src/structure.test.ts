import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CellMatrix, StructuralTransform, WorkbookModel } from './index';

function seedWorkbook(): { workbook: WorkbookModel; sheetId: string } {
  const workbook = new WorkbookModel('unit-test', 'Structural');
  const sheet = workbook.addSheet('s1', 'Sheet1');
  sheet.cells.set(0, 0, { value: 'A0' });
  sheet.cells.set(1, 0, { value: 'A1' });
  sheet.cells.set(2, 0, { value: 'A2' });
  return { workbook, sheetId: sheet.id };
}

describe('structural operations', () => {
  it('shiftRows moves cells below the insertion point', () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { value: 'top' });
    matrix.set(5, 1, { value: 'bottom' });
    matrix.shiftRows(3, 2, 1);
    assert.equal(matrix.get(0, 0)?.value, 'top');
    assert.equal(matrix.get(7, 1)?.value, 'bottom');
    matrix.shiftRows(3, 2, -1);
    assert.equal(matrix.get(5, 1)?.value, 'bottom');
  });

  it('StructuralTransform insertRows keeps merges and freeze consistent', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.merges.push({ range: { sheetId: 's1', startRow: 2, endRow: 3, startColumn: 1, endColumn: 1 }, anchor: { row: 2, column: 1 } });
    sheet.freeze = { xSplit: 0, ySplit: 2, startRow: 2, startColumn: 0 };
    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 2, count: 3 });
    assert.equal(sheet.cells.get(2, 0)?.value ?? null, null);
    assert.equal(sheet.cells.get(5, 0)?.value, 'A2');
    assert.equal(sheet.merges[0]!.range.startRow, 5);
    assert.equal(sheet.freeze.ySplit, 5);
    assert.equal(sheet.rowCount, 1003);
  });

  it('StructuralTransform deleteRows removes region and returns extracted cells for undo', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.cells.set(1, 0, { value: 'gone' });
    const removed = StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId: sheet.id, at: 1, count: 1 }).removedCells;
    assert.ok(removed.some((entry) => entry.cell.value === 'gone'));
    assert.equal(sheet.cells.get(1, 0)?.value, 'A2');
    // 恢复
    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 1, count: 1 });
    for (const entry of removed) sheet.cells.set(entry.row, entry.column, entry.cell);
    assert.equal(sheet.cells.get(1, 0)?.value, 'gone');
  });

  it('insertColumns shifts widths and hidden columns', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.columnWidths[2] = 200;
    sheet.hiddenColumns.add(3);
    StructuralTransform.apply(workbook, { kind: 'insert-columns', sheetId: sheet.id, at: 1, count: 2 });
    assert.equal(sheet.columnWidths[4], 200);
    assert.ok(sheet.hiddenColumns.has(5));
    assert.ok(!sheet.hiddenColumns.has(3));
  });

  it('structural formula rewrite uses AST references, including absolute and quoted refs', () => {
    const workbook = new WorkbookModel('unit-formula-structure', 'Structural Formula');
    const sheet = workbook.getSheet('sheet-1');
    sheet.name = 'Input Sheet';
    sheet.cells.set(5, 0, { value: null, formula: "=SUM($A$1,'Input Sheet'!$B$1,A1)+\"A1\"" });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 });

    assert.equal(sheet.cells.get(6, 0)?.formula, "=SUM($A$2,'Input Sheet'!$B$2,A2)+\"A1\"");
  });

  it('shift-cells moves formulas and metadata as one bounded transform', () => {
    const workbook = new WorkbookModel('unit-shift-cells', 'Shift Cells');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(1, 1, { value: null, formula: '=A1+B1' });
    sheet.notes.set('1:1', { id: 'n1', author: 'u', text: 'note', createdAt: 'now', visible: true });
    sheet.commentThreads.push({ id: 'c1', sheetId: sheet.id, row: 1, column: 1, author: 'u', text: 'comment', createdAt: 'now', replies: [] });

    StructuralTransform.apply(workbook, {
      kind: 'shift-cells-down',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
    });

    assert.equal(sheet.cells.get(2, 1)?.formula, '=A2+B2');
    assert.equal(sheet.cells.get(1, 1), undefined);
    assert.ok(sheet.notes.has('2:1'));
    assert.equal(sheet.commentThreads[0]?.row, 2);
  });

  it('structural row shifts update chart source ranges in the canonical drawing payload', () => {
    const workbook = new WorkbookModel('unit-drawing-structure', 'Drawing Structure');
    const sheet = workbook.getSheet('sheet-1');
    sheet.drawings.push({
      id: 'drawing-chart-1',
      sheetId: sheet.id,
      kind: 'chart',
      payloadId: 'chart-1',
      anchor: { kind: 'absolute' },
      transform: { x: 0, y: 0, width: 100, height: 80 },
      zIndex: 1,
    });
    sheet.drawingPayloads.set('chart-1', {
      kind: 'chart',
      chartId: 'chart-1',
      chartType: 'combo',
      stacked: 'percent',
      sourceRanges: [{ sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 1 }],
      categoryRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 },
      series: [{ name: 'Sales', range: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 } }],
    });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 1, count: 2 });

    const payload = sheet.drawingPayloads.get('chart-1');
    assert.equal(payload?.kind, 'chart');
    if (payload?.kind !== 'chart') throw new Error('Expected chart payload');
    assert.equal(payload.chartType, 'combo');
    assert.equal(payload.stacked, 'percent');
    assert.deepEqual(payload.sourceRanges[0], { sheetId: sheet.id, startRow: 3, endRow: 5, startColumn: 0, endColumn: 1 });
    assert.deepEqual(payload.categoryRange, { sheetId: sheet.id, startRow: 3, endRow: 5, startColumn: 0, endColumn: 0 });
    assert.deepEqual(payload.series?.[0]?.range, { sheetId: sheet.id, startRow: 3, endRow: 5, startColumn: 1, endColumn: 1 });
    assert.equal(sheet.drawings.filter((drawing) => drawing.kind === 'chart').length, 1);
  });

  it('structural transforms keep sheet-backed workbook table sources aligned', () => {
    const workbook = new WorkbookModel('unit-workbook-table-structure', 'Workbook Table Structure');
    const sheet = workbook.getSheet('sheet-1');
    workbook.addTable({
      id: 'table-1',
      name: 'Sales',
      sourceSheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 4, startColumn: 2, endColumn: 4 },
      rowCount: 3,
      fields: [],
      blockSize: 128,
      blocks: [],
      revision: 0,
    });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 1, count: 2 });
    assert.deepEqual(workbook.getTable('table-1').sourceRange, {
      sheetId: sheet.id,
      startRow: 3,
      endRow: 6,
      startColumn: 2,
      endColumn: 4,
    });

    assert.throws(
      () => StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId: sheet.id, at: 3, count: 1 }),
      /workbook table table-1 requires an explicit table operation/,
    );
  });

  it('move-range clears stale destinations, offsets formulas, and rewrites external references', () => {
    const workbook = new WorkbookModel('unit-move-range', 'Move Range');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 7 });
    sheet.cells.set(0, 1, { value: null, formula: '=A1' });
    sheet.cells.set(2, 3, { value: 'stale' });
    sheet.cells.set(0, 3, { value: null, formula: '=A1' });

    StructuralTransform.apply(workbook, {
      kind: 'move-range',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      targetOrigin: { row: 2, column: 2 },
    });

    assert.equal(sheet.cells.get(2, 2)?.value, 7);
    assert.equal(sheet.cells.get(2, 3)?.formula, '=C3');
    assert.equal(sheet.cells.get(0, 0), undefined);
    assert.equal(sheet.cells.get(0, 3)?.formula, '=C3');
  });

  it('rejects a bounded shift that would silently drop an anchored object', () => {
    const workbook = new WorkbookModel('unit-shift-reject', 'Shift Reject');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(1, 0, { value: 'keep' });
    sheet.drawings.push({
      id: 'drawing-edge',
      sheetId: sheet.id,
      kind: 'shape',
      payloadId: 'shape-edge',
      anchor: { kind: 'one-cell', row: 1, column: 0 },
      transform: { x: 0, y: 0, width: 20, height: 20 },
      zIndex: 0,
    });
    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'shift-cells-down',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    }), /drawing drawing-edge would leave/);
    assert.equal(sheet.cells.get(1, 0)?.value, 'keep');
    assert.equal(sheet.drawings[0]?.anchor.row, 1);
  });
});
