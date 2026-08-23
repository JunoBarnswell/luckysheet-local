import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  computeConditionalOverlays,
  registerSheetCommands,
  validateDataInput,
} from './index';

function runtime(): { workbook: WorkbookModel; commands: CommandRuntime } {
  const workbook = new WorkbookModel('m3-m4', 'M3/M4');
  const commands = new CommandRuntime(workbook);
  registerSheetCommands(commands);
  return { workbook, commands };
}

test('scoped defined names survive command undo and snapshot round-trip', () => {
  const { workbook, commands } = runtime();
  const local = workbook.addSheet('sheet-2', 'Local');
  commands.execute('workbook.name.set', { name: 'Rate', value: '0.1' });
  commands.execute('workbook.name.set', { name: 'Rate', formula: '0.2', scope: 'sheet', sheetId: local.id });
  assert.equal(workbook.getDefinedName('Rate', local.id)?.formula, '0.2');
  commands.undo();
  assert.equal(workbook.getDefinedNameExact('Rate', 'sheet', local.id), undefined);
  const restored = WorkbookModel.fromSnapshot(workbook.snapshot());
  assert.equal(restored.getDefinedName('Rate')?.formula, '0.1');
});

test('total row inserts a real row and does not overwrite data below table', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.activeSheetId);
  sheet.cells.set(0, 0, { value: 'Name' });
  sheet.cells.set(1, 0, { value: 'A' });
  sheet.cells.set(2, 0, { value: 'B' });
  sheet.cells.set(3, 0, { value: 'outside' });
  commands.execute('sheetTable.add', {
    id: 'table-1', sheetId: sheet.id, name: 'Sales',
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: false,
    showBandedColumns: false, showFilterButton: false,
    columns: [{ id: 'name', name: 'Name', totalsFunction: 'count' }],
  });
  commands.execute('sheetTable.toggleTotalRow', { sheetId: sheet.id, tableId: 'table-1', enabled: true });
  assert.equal(sheet.cells.get(4, 0)?.value, 'outside');
  assert.match(sheet.cells.get(3, 0)?.formula ?? '', /^=SUBTOTAL\(/);
  commands.undo();
  assert.equal(sheet.cells.get(3, 0)?.value, 'outside');
});

test('sort and remove duplicates preserve formulas and use structural row deletion', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.activeSheetId);
  commands.execute('sheet.range.set', {
    sheetId: sheet.id, startRow: 0, startColumn: 0,
    values: [
      [{ value: 'Key' }, { value: 'Value' }, { value: 'Formula' }],
      [{ value: 'B' }, { value: 2 }, { formula: '=B2*2', value: null }],
      [{ value: 'A' }, { value: 1 }, { formula: '=B3*2', value: null }],
      [{ value: 'A' }, { value: 1 }, { formula: '=B4*2', value: null }],
    ],
  });
  commands.execute('sheet.sort', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }, sortColumn: 0, ascending: true, hasHeader: true });
  assert.equal(sheet.cells.get(1, 0)?.value, 'A');
  assert.equal(sheet.cells.get(1, 2)?.formula, '=B3*2');
  commands.execute('data.removeDuplicates', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }, columns: [0], hasHeader: true });
  assert.equal(sheet.cells.get(3, 0), undefined);
  assert.equal(sheet.cells.get(2, 2)?.formula, '=B2*2');
});

test('conditional format priority/stop and validation alert style are represented', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.activeSheetId);
  commands.execute('sheet.range.set', { sheetId: sheet.id, startRow: 0, startColumn: 0, values: [[{ value: 10 }], [{ value: 1 }]] });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: { id: 'top', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }], type: 'topBottom', topBottom: { direction: 'top', rank: 1 }, priority: 1, stopIfTrue: true, style: { bold: true } },
  });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: { id: 'lower', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }], type: 'highlight', operator: 'lessThan', value1: 20, priority: 2, style: { italic: true } },
  });
  assert.equal(computeConditionalOverlays(sheet).get('0:0')?.style?.bold, true);
  commands.execute('sheet.dv.add', {
    sheetId: sheet.id,
    rule: { id: 'list', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }], type: 'list', listSource: { kind: 'values', values: ['A', 'B'] }, multiSelect: true, alertStyle: 'warning' },
  });
  const result = validateDataInput(sheet, 0, 1, 'A,B');
  assert.equal(result.valid, true);
  assert.equal(result.blocking, false);
});

test('transpose fails closed when a selected range contains a drawing', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.activeSheetId);
  sheet.drawings.push({ id: 'd1', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 0 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'p1' });
  assert.throws(() => commands.execute('matrix.transpose', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } }), /drawing anchors/);
});
