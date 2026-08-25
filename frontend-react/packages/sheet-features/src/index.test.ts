import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import {
  registerSheetCommands,
  formatTsv,
  parseTsv,
  parseClipboardPayload,
  shiftFormula,
  FormulaRelocationError,
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
  assert.equal(shiftFormula('=SUM(A2#)', 1, 1), '=SUM(B3#)');
  assert.throws(() => shiftFormula('=A1+', 1, 1), FormulaRelocationError);
});

test('clipboard payload carries provenance and paste modes preserve their contracts', () => {
  const workbook = new WorkbookModel('unit-paste-modes', 'Paste Modes');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));
  sheet.cells.set(0, 0, { value: 12, formula: '=A2', style: { bold: true }, numberFormat: '0.00' });
  sheet.cells.set(0, 1, { value: 'keep', style: { italic: true }, numberFormat: '@' });

  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0,
  });
  assert.equal(payload.source, 'internal');
  assert.equal(payload.mime, 'application/x-react-sheets-cells');

  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 0, column: 1 },
    clipboard: payload,
    transfer: 'copy',
    mode: 'values',
  });
  assert.deepEqual(sheet.cells.get(0, 1), { value: 12 });

  sheet.cells.set(0, 1, { value: 'keep', formula: '=A1', style: { italic: true }, numberFormat: '@' });
  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 0, column: 1 },
    clipboard: payload,
    transfer: 'copy',
    mode: 'formats',
  });
  assert.equal(sheet.cells.get(0, 1)?.value, 'keep');
  assert.equal(sheet.cells.get(0, 1)?.formula, '=A1');
  assert.equal(sheet.cells.get(0, 1)?.style?.bold, true);
  assert.equal(sheet.cells.get(0, 1)?.numberFormat, '0.00');
});

test('clipboard uses quoted TSV and host-neutral HTML representations', () => {
  const text = formatTsv([[{ value: 'a\tb' }, { value: 'line\nnext' }, { value: '"quoted"' }]]);
  assert.deepEqual(parseTsv(text).map((row) => row.map((cell) => cell.value)), [['a\tb', 'line\nnext', '"quoted"']]);
  const payload = {
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    values: [],
    transfer: 'copy' as const,
    representations: [{ mime: 'text/html', data: '<table><tr><td>42</td><td data-formula="=A1+1">43</td></tr></table>' }],
  };
  assert.deepEqual(parseClipboardPayload(payload).map((row) => row.map((cell) => cell.value)), [[42, null]]);
  assert.equal(parseClipboardPayload(payload)[0]?.[1]?.formula, '=A1+1');
  assert.deepEqual(parseTsv(' 42 \t true ').map((row) => row.map((cell) => cell.value)), [[' 42 ', ' true ']]);
});

test('sheet.cell.commitText is the single raw-text input path', () => {
  const workbook = new WorkbookModel('unit-commit-text', 'Commit Text');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));
  sheet.cells.set(0, 0, {
    value: 'old',
    style: { bold: true, background: '#fff' },
    styleId: 'style-1',
    numberFormat: '@',
    displayValue: 'old',
  });

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: ' 42 ' });
  assert.equal(sheet.cells.get(0, 0)?.value, ' 42 ');
  assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);
  assert.equal(sheet.cells.get(0, 0)?.styleId, 'style-1');
  assert.equal(sheet.cells.get(0, 0)?.displayValue, undefined);

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: '=A1 + 1' });
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1 + 1');
  assert.equal(sheet.cells.get(0, 0)?.value, null);
  assert.equal(sheet.cells.get(0, 0)?.style?.background, '#fff');

  assert.throws(
    () => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: '=A1 +' }),
    /formula|expression|unexpected/i,
  );
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1 + 1');

  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, ' 42 ');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'old');
  assert.equal(sheet.cells.get(0, 0)?.formula, undefined);
});

test('sheet.cell.commitText parses scalars, validates input and protects spill children', () => {
  const workbook = new WorkbookModel('unit-commit-text-guards', 'Commit Guards');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: '42' });
  assert.equal(sheet.cells.get(0, 0)?.value, 42);
  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 1, text: 'TRUE' });
  assert.equal(sheet.cells.get(0, 1)?.value, true);
  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 2, text: ' true ' });
  assert.equal(sheet.cells.get(0, 2)?.value, ' true ');

  sheet.dataValidations.push({
    id: 'list-commit',
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }],
    type: 'list',
    listSource: { kind: 'values', values: ['Allowed'] },
    alertStyle: 'stop',
  });
  assert.throws(
    () => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 1, column: 0, text: 'Rejected' }),
    /允许的列表|validation/i,
  );
  assert.equal(sheet.cells.get(1, 0), undefined);

  sheet.spillRanges.push({
    sheetId: sheet.id,
    anchor: { row: 2, column: 0 },
    range: { sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 1 },
    values: [[1, 2]],
    state: 'ok',
  });
  assert.throws(
    () => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 2, column: 1, text: 'blocked' }),
    /spill child/i,
  );
  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 2, column: 0, text: 'anchor' });
  assert.equal(sheet.cells.get(2, 0)?.value, 'anchor');
});

test('cut paste is one cross-sheet transaction and preserves formula references', () => {
  const workbook = new WorkbookModel('unit-cut', 'Cut');
  const target = workbook.addSheet('sheet-2', 'Target');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const source = workbook.getSheet('sheet-1');
  source.cells.set(1, 1, { value: null, formula: '=A1+$B$1' });
  target.cells.set(4, 4, { value: 'old' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: source.id,
    startRow: 1,
    endRow: 1,
    startColumn: 1,
    endColumn: 1,
  });
  payload.transfer = 'move';
  const result = runtime.execute('sheet.range.paste', {
    sheetId: target.id,
    targetOrigin: { row: 4, column: 4 },
    clipboard: payload,
    transfer: 'move',
  });
  assert.equal(result.mutationCount, 1);
  assert.equal(source.cells.get(1, 1), undefined);
  assert.equal(target.cells.get(4, 4)?.formula, '=A1+$B$1');
  assert.equal(target.cells.get(4, 4)?.value, null);
  runtime.undo();
  assert.equal(source.cells.get(1, 1)?.formula, '=A1+$B$1');
  assert.equal(target.cells.get(4, 4)?.value, 'old');
  runtime.redo();
  assert.equal(source.cells.get(1, 1), undefined);
  assert.equal(target.cells.get(4, 4)?.formula, '=A1+$B$1');
});

test('copy paste shifts relative references while preserving mixed and absolute references', () => {
  const workbook = new WorkbookModel('unit-copy', 'Copy');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(1, 1, { value: null, formula: '=A1+$B$1+C$1+$D1' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 1,
    endRow: 1,
    startColumn: 1,
    endColumn: 1,
  });
  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 4, column: 4 },
    clipboard: payload,
    transfer: 'copy',
  });
  assert.equal(sheet.cells.get(4, 4)?.formula, '=D4+$B$1+F$1+$D4');
  assert.equal(sheet.cells.get(1, 1)?.formula, '=A1+$B$1+C$1+$D1');
});

test('paste replay rejects a transfer mismatch before touching the workbook', () => {
  const workbook = new WorkbookModel('unit-paste-contract', 'Paste Contract');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: null, formula: '=A1' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0,
  });
  payload.transfer = 'move';
  assert.throws(() => runtime.applyRemoteMutations([{
    id: 'range.paste',
    unitId: workbook.unitId,
    sheetId: sheet.id,
    params: {
      sheetId: sheet.id,
      targetOrigin: { row: 1, column: 1 },
      clipboard: payload,
      transfer: 'copy',
      mode: 'all',
      values: [[{ value: null, formula: '=A1' }]],
      startRow: 1,
      startColumn: 1,
      sourceRange: structuredClone(payload.range),
      clearSource: true,
    },
    affectedRanges: [],
  }]), /Invalid mutation history/);
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1');
  assert.equal(sheet.cells.get(1, 1), undefined);
});

test('copy paste rejects a malformed formula before creating a mutation', () => {
  const workbook = new WorkbookModel('unit-paste-fail-close', 'Paste Fail Close');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: null, formula: '=A1+' });
  sheet.cells.set(1, 1, { value: 'unchanged' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0,
  });
  assert.throws(() => runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 1, column: 1 },
    clipboard: payload,
    transfer: 'copy',
  }), FormulaRelocationError);
  assert.equal(sheet.cells.get(1, 1)?.value, 'unchanged');
  assert.equal(runtime.getHistoryDepth().undo, 0);
});

test('range clear modes are independent and restore auxiliary metadata', () => {
  const workbook = new WorkbookModel('unit-clear-modes', 'Clear Modes');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, {
    value: 10,
    formula: '=A1',
    style: { bold: true },
    hyperlink: 'https://example.com',
    note: { id: 'cell-note', author: 'u', text: 'note', createdAt: '2026-01-01', visible: true },
  });
  sheet.notes.set('0:0', { id: 'note', author: 'u', text: 'standalone', createdAt: '2026-01-01', visible: true });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, mode: 'formats' });
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1');
  assert.equal(sheet.cells.get(0, 0)?.style, undefined);
  assert.equal(sheet.cells.get(0, 0)?.hyperlink, 'https://example.com');
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, mode: 'notes' });
  assert.equal(sheet.cells.get(0, 0)?.note, undefined);
  assert.equal(sheet.notes.has('0:0'), false);
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, mode: 'hyperlinks' });
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
  assert.equal(sheet.cells.get(0, 0)?.hyperlink, undefined);
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, mode: 'all' });
  assert.equal(sheet.cells.get(0, 0), undefined);
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
});

function sheetId(workbook: WorkbookModel): string {
  return workbook.primarySheetId;
}

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
  sheet.pane = { kind: 'frozen', xSplit: 0, ySplit: 1, startRow: 0, startColumn: 0, state: 'frozen' };

  runtime.execute('sheet.rows.insert', { sheetId: sheet.id, at: 1, count: 2 });
  assert.equal(sheet.rowCount, 1002);
  assert.equal(sheet.cells.get(4, 0)?.value, 42);
  assert.equal(sheet.cells.get(4, 0)?.formula, '=A1+1');
  assert.equal(sheet.merges[0]?.range.startRow, 4);
  assert.equal(sheet.pane.kind === 'frozen' ? sheet.pane.ySplit : 0, 3);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, 42);
  assert.equal(sheet.merges[0]?.range.startRow, 2);

  runtime.execute('sheet.rows.delete', { sheetId: sheet.id, at: 2, count: 1 });
  assert.equal(sheet.cells.get(2, 0), undefined);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, 42);
});

test('sheet.cells.insert undo restores the complete affected band', () => {
  const workbook = new WorkbookModel('unit-shift-undo', 'Shift Undo');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));
  sheet.cells.set(0, 0, { value: 'top' });
  sheet.cells.set(1, 0, { value: 'bottom' });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };

  runtime.execute('sheet.cells.insert', { sheetId: sheet.id, range, operation: 'insert', axis: 'row' });
  assert.equal(sheet.cells.get(0, 0), undefined);
  assert.equal(sheet.cells.get(1, 0)?.value, 'top');
  assert.equal(sheet.cells.get(2, 0)?.value, 'bottom');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'top');
  assert.equal(sheet.cells.get(1, 0)?.value, 'bottom');
});
