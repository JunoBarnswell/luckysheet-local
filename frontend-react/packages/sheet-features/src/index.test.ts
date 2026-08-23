import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import {
  registerSheetCommands,
  formatTsv,
  parseTsv,
  shiftFormula,
  copyRangeToClipboardData,
} from './index';

test('sheet commands: cell.set, range.set, and undo/redo', () => {
  const workbook = new WorkbookModel('unit-sheet-cmd', 'Commands');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);

  // 1. Single cell set
  runtime.execute('sheet.cell.set', {
    sheetId: 'sheet-1',
    row: 0,
    column: 0,
    value: { value: 'Title', style: { bold: true } },
  });
  const sheet = workbook.getSheet('sheet-1');
  assert.equal(sheet.cells.get(0, 0)?.value, 'Title');
  assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);

  // 2. Undo
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0), undefined);

  // 3. Redo
  runtime.redo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'Title');

  // 4. Batch range set
  runtime.execute('sheet.range.set', {
    sheetId: 'sheet-1',
    startRow: 1,
    startColumn: 0,
    values: [
      [{ value: 10 }, { value: 20 }],
      [{ value: 30 }, { value: 40 }],
    ],
  });
  assert.equal(sheet.cells.get(1, 0)?.value, 10);
  assert.equal(sheet.cells.get(1, 1)?.value, 20);
  assert.equal(sheet.cells.get(2, 0)?.value, 30);
  assert.equal(sheet.cells.get(2, 1)?.value, 40);

  runtime.undo();
  assert.equal(sheet.cells.get(1, 0), undefined);
  assert.equal(sheet.cells.get(2, 1), undefined);
  assert.equal(sheet.cells.get(0, 0)?.value, 'Title'); // untouched
});

test('sheet commands: range.clear, style.set, and merges', () => {
  const workbook = new WorkbookModel('unit-sheet-cmd2', 'Commands2');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  // Setup cells
  sheet.cells.set(0, 0, { value: 'Hello', style: { bold: true, background: '#ff0000' } });
  sheet.cells.set(0, 1, { value: 'World', style: { bold: true } });

  // Style batch
  runtime.execute('sheet.style.set', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    style: { background: '#00ff00' },
  });
  assert.equal(sheet.cells.get(0, 0)?.style?.background, '#00ff00');
  assert.equal(sheet.cells.get(0, 1)?.style?.background, '#00ff00');

  // Merges
  runtime.execute('sheet.merge.set', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
  });
  assert.equal(sheet.merges.length, 1);
  assert.equal(sheet.isMerged(0, 0)?.anchor.column, 0);

  runtime.execute('sheet.merge.remove', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
  });
  assert.equal(sheet.merges.length, 0);

  // Range clear
  runtime.execute('sheet.range.clear', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    mode: 'all',
  });
  assert.equal(sheet.cells.get(0, 0), undefined);
  assert.equal(sheet.cells.get(0, 1), undefined);
});

test('sheet commands: sort and autofill', () => {
  const workbook = new WorkbookModel('unit-sort', 'SortAutofill');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  // Populate data for sorting
  sheet.cells.set(0, 0, { value: 'Name' });
  sheet.cells.set(0, 1, { value: 'Score' });
  sheet.cells.set(1, 0, { value: 'Charlie' });
  sheet.cells.set(1, 1, { value: 70 });
  sheet.cells.set(2, 0, { value: 'Alice' });
  sheet.cells.set(2, 1, { value: 95 });
  sheet.cells.set(3, 0, { value: 'Bob' });
  sheet.cells.set(3, 1, { value: 85 });

  // Sort ascending by Score (col 1), with header
  runtime.execute('sheet.sort', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    sortColumn: 1,
    ascending: true,
    hasHeader: true,
  });

  assert.equal(sheet.cells.get(0, 0)?.value, 'Name'); // header intact
  assert.equal(sheet.cells.get(1, 0)?.value, 'Charlie'); // 70
  assert.equal(sheet.cells.get(2, 0)?.value, 'Bob');     // 85
  assert.equal(sheet.cells.get(3, 0)?.value, 'Alice');   // 95

  // Autofill formula
  sheet.cells.set(5, 0, { value: null, formula: '=A1+B1' });
  runtime.execute('sheet.autofill', {
    sheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0 },
    targetRange: { sheetId: 'sheet-1', startRow: 5, endRow: 7, startColumn: 0, endColumn: 0 },
  });
  assert.equal(sheet.cells.get(6, 0)?.formula, '=A2+B2');
  assert.equal(sheet.cells.get(7, 0)?.formula, '=A3+B3');
});

test('clipboard: TSV format, parse, and formula shifting', () => {
  const tsv = formatTsv([
    [{ value: 'Product' }, { value: 'Price' }],
    [{ value: 'Keyboard' }, { value: 120 }],
  ]);
  assert.equal(tsv, 'Product\tPrice\nKeyboard\t120');

  const parsed = parseTsv('Product\tPrice\nKeyboard\t120\nTotal\t=SUM(B2:B2)');
  assert.equal(parsed.length, 3);
  assert.equal(parsed[0]?.[0]?.value, 'Product');
  assert.equal(parsed[1]?.[1]?.value, 120);
  assert.equal(parsed[2]?.[1]?.formula, '=SUM(B2:B2)');

  assert.equal(shiftFormula('=A1+$B$1+C$1+$D1', 2, 3), '=D3+$B$1+F$1+$D3');
});

test('sheet commands: hide and unhide rows and columns are undoable commands', () => {
  const workbook = new WorkbookModel('unit-hidden', 'Hidden');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  runtime.execute('sheet.row.hide', { sheetId: sheet.id, index: 2 });
  runtime.execute('sheet.column.hide', { sheetId: sheet.id, index: 3 });
  assert.equal(sheet.hiddenRows.has(2), true);
  assert.equal(sheet.hiddenColumns.has(3), true);

  runtime.undo();
  assert.equal(sheet.hiddenColumns.has(3), false);
  runtime.undo();
  assert.equal(sheet.hiddenRows.has(2), false);

  runtime.redo();
  runtime.redo();
  assert.equal(sheet.hiddenRows.has(2), true);
  assert.equal(sheet.hiddenColumns.has(3), true);

  runtime.execute('sheet.rows.unhide.all', { sheetId: sheet.id });
  runtime.execute('sheet.columns.unhide.all', { sheetId: sheet.id });
  assert.equal(sheet.hiddenRows.size, 0);
  assert.equal(sheet.hiddenColumns.size, 0);
  runtime.undo();
  assert.equal(sheet.hiddenColumns.has(3), true);
  runtime.undo();
  assert.equal(sheet.hiddenRows.has(2), true);
});

test('sheet commands: row insert/delete use StructuralTransform and preserve undo', () => {
  const workbook = new WorkbookModel('unit-structural', 'Structural');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  sheet.cells.set(2, 0, { value: 42, formula: '=A1+1' });
  sheet.merges.push({
    range: { sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 1 },
    anchor: { row: 2, column: 0 },
  });
  sheet.freeze = { xSplit: 0, ySplit: 1, startRow: 0, startColumn: 0 };

  runtime.execute('sheet.rows.insert', { sheetId: sheet.id, at: 1, count: 2 });
  assert.equal(sheet.rowCount, 1002);
  assert.equal(sheet.cells.get(4, 0)?.value, 42);
  assert.equal(sheet.cells.get(4, 0)?.formula, '=A1+1');
  assert.equal(sheet.merges[0]?.range.startRow, 4);
  assert.equal(sheet.freeze.ySplit, 3);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, 42);
  assert.equal(sheet.merges[0]?.range.startRow, 2);

  runtime.execute('sheet.rows.delete', { sheetId: sheet.id, at: 2, count: 1 });
  assert.equal(sheet.cells.get(2, 0), undefined);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, 42);
});
