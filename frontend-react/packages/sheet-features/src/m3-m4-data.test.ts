import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  computeConditionalOverlays,
  computeFilterHiddenRows,
  getAutoFilterValueDomain,
  registerSheetCommands,
  normalizeAutoFilterModel,
  normalizeDataValidationRule,
  validationList,
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
  const sheet = workbook.getSheet(workbook.primarySheetId);
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
  const sheet = workbook.getSheet(workbook.primarySheetId);
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
  const sheet = workbook.getSheet(workbook.primarySheetId);
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
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.drawings.push({ id: 'd1', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 0 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'p1' });
  assert.throws(() => commands.execute('matrix.transpose', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } }), /drawing anchors/);
});

test('Filter supports compound text/blank/date conditions and rejects out-of-range criteria', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Value' });
  sheet.cells.set(1, 0, { value: 'Alpha' });
  sheet.cells.set(2, 0, { value: '' });
  sheet.cells.set(3, 0, { value: 'Beta' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'custom', join: 'and', conditions: [{ operator: 'contains', value: 'a' }] } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((a, b) => a - b), [2]);
  assert.throws(() => normalizeAutoFilterModel({
    ...sheet.autoFilter!,
    columns: { 1: { column: 1, showButton: true, hiddenButton: false, criterion: { kind: 'custom', join: 'and', conditions: [{ operator: 'equals', value: 'x' }] } } },
  }), /outside/);
});

test('AutoFilter value domain is complete and ignores the current column criterion', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  for (let row = 0; row <= 250; row += 1) {
    sheet.cells.set(row, 0, { value: row === 0 ? 'Name' : `Value-${row}` });
    sheet.cells.set(row, 1, { value: row === 0 ? 'Group' : row % 2 === 0 ? 'keep' : 'drop' });
  }
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 250, startColumn: 0, endColumn: 1 },
    columns: {
      0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['Value-2'], includeBlank: false } },
      1: { column: 1, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['keep'], includeBlank: false } },
    },
  });
  const domain = getAutoFilterValueDomain(sheet, 0);
  assert.equal(domain.length, 125);
  assert.equal(domain.includes('Value-2'), true);
  assert.equal(domain.includes('Value-3'), false);
});

test('AutoFilter evaluates Top10 and dynamic date criteria against canonical rows', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Amount' });
  [10, 5, 20, 1].forEach((value, index) => sheet.cells.set(index + 1, 0, { value }));
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'top10', top: true, percent: false, rank: 2 } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((a, b) => a - b), [2, 4]);

  const today = new Date();
  const isoToday = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  sheet.cells.set(0, 0, { value: 'Date' });
  sheet.cells.set(1, 0, { value: isoToday });
  sheet.cells.set(2, 0, { value: '2000-01-01' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'dynamic', type: 'today' } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [2]);

  assert.throws(() => normalizeAutoFilterModel({
    ...sheet.autoFilter!,
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'dynamic', type: 'attackerUnknown' as 'today' } } },
  }), /UNSUPPORTED_FEATURE: dynamic AutoFilter type "attackerUnknown" is not supported/);
});

test('AutoFilter color and icon criteria use native cell metadata or imported differential style', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Status' });
  sheet.cells.set(1, 0, { value: 'red', style: { background: '#ff0000' } });
  sheet.cells.set(2, 0, { value: 'blue', style: { background: '#0000ff' } });
  sheet.cells.set(3, 0, { value: 'icon', filterMetadata: { icon: { iconSet: '3TrafficLights1', iconId: 2 } } });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#ff0000' } } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [2, 3]);
  sheet.autoFilter.columns[0]!.criterion = { kind: 'icon', iconSet: '3TrafficLights1', iconId: 2 };
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [1, 2]);
});

test('Validation supports custom AST, formula-backed list, time/date, multi-select and non-blocking alerts', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Allowed' });
  sheet.cells.set(1, 0, { value: 'Other' });
  const custom = normalizeDataValidationRule({ id: 'custom', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }], type: 'custom', formula1: '=B1="Allowed"', alertStyle: 'stop' });
  sheet.dataValidations.push(custom);
  assert.equal(validateDataInput(sheet, 0, 1, 'x').blocking, true);
  const list = normalizeDataValidationRule({ id: 'list', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 }], type: 'list', listSource: { kind: 'formula', formula: '=A1:A2' }, multiSelect: true, alertStyle: 'warning' });
  assert.deepEqual(validationList(list, sheet), ['Allowed', 'Other']);
  sheet.dataValidations.push(list);
  assert.equal(validateDataInput(sheet, 0, 2, 'Allowed,Other').blocking, false);
  const time = normalizeDataValidationRule({ id: 'time', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }], type: 'time', alertStyle: 'stop' });
  sheet.dataValidations.push(time);
  assert.equal(validateDataInput(sheet, 0, 3, '12:30').valid, true);
  assert.equal(validateDataInput(sheet, 0, 3, '25:30').blocking, true);
});

test('Text Columns, Split and Flip are one undoable transaction and clear stale output', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'a,b' });
  sheet.cells.set(0, 2, { value: 'stale' });
  let commandEvents = 0;
  commands.onCommand(() => { commandEvents += 1; });
  const textResult = commands.execute('data.textToColumns', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, delimiter: ',', maxColumns: 3 });
  assert.equal(commandEvents, 1);
  assert.equal(textResult.mutationCount, 2);
  assert.equal(sheet.cells.get(0, 2)?.value, null);
  commands.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'a,b');
  commands.execute('data.splitColumn', { sheetId: sheet.id, row: 0, column: 0, delimiter: ',', maxColumns: 2 });
  assert.equal(sheet.cells.get(0, 1)?.value, 'b');
  commands.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'a,b');
  sheet.cells.set(0, 0, { value: 'a' });
  sheet.cells.set(0, 1, { value: 'b' });
  commands.execute('matrix.flip', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, direction: 'horizontal' });
  assert.equal(sheet.cells.get(0, 0)?.value, 'b');
});
