import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';

function fixture() {
  const workbook = new WorkbookModel('dimension-uat', 'Dimension UAT');
  const runtime = new CommandRuntime(workbook); registerSheetCommands(runtime);
  const sheetId = workbook.sheetOrder[0]!; const sheet = workbook.getSheet(sheetId);
  return { runtime, sheetId, sheet };
}
test('dimension size and unhide commit and undo as one transaction', () => {
  const { runtime, sheetId, sheet } = fixture();
  sheet.hiddenColumns.add(0); sheet.hiddenRows.add(0);
  const oldWidth = sheet.defaultColumnWidthPx, oldHeight = sheet.defaultRowHeightPx;
  runtime.execute('sheet.dimensions.apply', { sheetId, columns: [{ column: 0, widthPx: 180, hidden: false }], rows: [{ row: 0, heightPx: 42, hidden: false }] });
  assert.equal(sheet.hiddenColumns.has(0), false); assert.equal(sheet.hiddenRows.has(0), false);
  assert.equal(sheet.columnWidthsPx[0], 180); assert.equal(sheet.rowHeightsPx[0], 42);
  runtime.undo();
  assert.equal(sheet.hiddenColumns.has(0), true); assert.equal(sheet.hiddenRows.has(0), true);
  assert.equal(sheet.columnWidthsPx[0], oldWidth); assert.equal(sheet.rowHeightsPx[0], oldHeight);
  runtime.redo(); assert.equal(sheet.columnWidthsPx[0], 180); assert.equal(sheet.hiddenColumns.has(0), false);
});
test('an invalid dimension in a multi-axis plan causes zero size or visibility changes', () => {
  const { runtime, sheetId, sheet } = fixture(); sheet.hiddenRows.add(0);
  assert.throws(() => runtime.execute('sheet.dimensions.apply', { sheetId, rows: [{ row: 0, heightPx: 42, hidden: false }], columns: [{ column: sheet.columnCount, widthPx: 180 }] }), /Invalid column dimension plan/);
  assert.equal(sheet.hiddenRows.has(0), true); assert.equal(sheet.rowHeightsPx[0], undefined); assert.deepEqual(sheet.columnWidthsPx, {});
});
