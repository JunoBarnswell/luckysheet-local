import test from 'node:test';
import assert from 'node:assert/strict';
import { WorksheetModel } from '@react-sheets/core-model';
import type { SheetTableModel } from '@react-sheets/core-model';
import {
  buildTotalRowFormula,
  computeSheetTableCellStyle,
  createAutoFilterModelForTable,
  findSheetTableAt,
  isPointInRange,
  mergePresentationStyles,
  planTotalRowToggle,
  planSheetTableAutoExpansion,
  planSheetTableCreation,
  resolveFilterButtonCells,
  resolveAutoFilters,
  resolveActiveAutoFilter,
  resolveFilterOwner,
  validateFilterOwnership,
  subtotalCodeForTotalsFunction,
  tableBodyBounds,
} from './sheet-table-features';

const sampleTable: SheetTableModel = {
  id: 't1',
  sheetId: 's1',
  name: 'Sales',
  range: { sheetId: 's1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
  hasHeaderRow: true,
  hasTotalRow: true,
  showBandedRows: true,
  showBandedColumns: false,
  showFirstColumn: false,
  showLastColumn: false,
  showFilterButton: true,
  autoExpand: 'both',
  columns: [{ id: 'c1', name: 'Product' }, { id: 'c2', name: 'Amount' }],
};

test('isPointInRange and findSheetTableAt locate cells inside a sheet table', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  sheet.sheetTables.push(sampleTable);
  assert.equal(isPointInRange(sampleTable.range, 1, 0), true);
  assert.equal(findSheetTableAt(sheet, 1, 0)?.name, 'Sales');
});

test('computeSheetTableCellStyle styles header, body bands, and total row', () => {
  assert.equal(computeSheetTableCellStyle(sampleTable, 0, 0)?.background, '#4472C4');
  assert.equal(computeSheetTableCellStyle(sampleTable, 1, 0)?.background, '#FFFFFF');
  assert.equal(computeSheetTableCellStyle(sampleTable, 2, 0)?.background, '#D9E1F2');
  assert.equal(computeSheetTableCellStyle(sampleTable, 3, 0)?.bold, true);
  assert.deepEqual(tableBodyBounds(sampleTable), { startRow: 1, endRow: 2 });
});

test('createAutoFilterModelForTable uses the table range', () => {
  const filter = createAutoFilterModelForTable(sampleTable);
  assert.deepEqual(filter.range, sampleTable.range);
  assert.equal(filter.sheetId, 's1');
});

test('planSheetTableCreation preserves body rows when headers are disabled', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  let id = 0;
  const plan = planSheetTableCreation({
    sheetId: 's1',
    range: { sheetId: 's1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    name: 'BodyOnly',
    hasHeaderRow: false,
    nextId: (prefix) => `${prefix}-${id++}`,
    readCell: () => 'must-not-be-used',
  }, sheet);
  assert.deepEqual(plan.table.columns.map((column) => column.name), ['Column1', 'Column2']);
  assert.equal(plan.table.hasHeaderRow, false);
  assert.equal(plan.table.showFilterButton, false);
  assert.equal(plan.table.autoFilter, undefined);
});

test('planSheetTableAutoExpansion grows contiguous rows and columns atomically', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const table: SheetTableModel = {
    ...sampleTable,
    hasTotalRow: false,
    range: { sheetId: 's1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    columns: [{ id: 'c1', name: 'Product' }, { id: 'c2', name: 'Amount' }],
  };
  sheet.sheetTables.push(table);
  const rowPlan = planSheetTableAutoExpansion(sheet, { sheetId: 's1', startRow: 3, endRow: 4, startColumn: 0, endColumn: 1 }, (prefix) => `${prefix}-row`);
  assert.equal(rowPlan.length, 1);
  assert.equal(rowPlan[0]?.next.range.endRow, 4);
  const columnPlan = planSheetTableAutoExpansion(sheet, { sheetId: 's1', startRow: 0, endRow: 2, startColumn: 2, endColumn: 2 }, (prefix) => `${prefix}-column`);
  assert.equal(columnPlan.length, 1);
  assert.equal(columnPlan[0]?.next.range.endColumn, 2);
  assert.equal(columnPlan[0]?.next.columns.length, 3);
});

test('mergePresentationStyles keeps later layers on top', () => {
  const merged = mergePresentationStyles({ background: '#ffffff' }, { bold: true, background: '#D9E1F2' });
  assert.deepEqual(merged, { background: '#D9E1F2', bold: true });
});

test('buildTotalRowFormula maps totals functions to SUBTOTAL codes', () => {
  assert.equal(subtotalCodeForTotalsFunction('sum'), 109);
  assert.equal(buildTotalRowFormula('Sales', 'Amount', 'sum'), '=SUBTOTAL(109,Sales[Amount])');
  assert.equal(buildTotalRowFormula('Sales', 'Unit Price', 'average'), '=SUBTOTAL(101,Sales[Unit Price])');
  assert.equal(buildTotalRowFormula('Sales', 'Product', 'none'), null);
});

test('planTotalRowToggle expands range and writes total row formulas', () => {
  const table: SheetTableModel = {
    ...sampleTable,
    hasTotalRow: false,
    range: { sheetId: 's1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    columns: [
      { id: 'c1', name: 'Product', totalsFunction: 'none' },
      { id: 'c2', name: 'Amount', totalsFunction: 'sum' },
    ],
  };
  const plan = planTotalRowToggle(table, true);
  assert.equal(plan.nextTable.hasTotalRow, true);
  assert.equal(plan.nextTable.range.endRow, 3);
  assert.equal(plan.values[0]?.[0]?.value, 'Total');
  assert.equal(plan.values[0]?.[1]?.formula, '=SUBTOTAL(109,Sales[Amount])');
});

test('resolveFilterButtonCells targets table header row when filter matches table', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  sheet.sheetTables.push(sampleTable);
  sheet.autoFilter = createAutoFilterModelForTable(sampleTable);
  sheet.autoFilter!.columns[0]!.criterion = { kind: 'values', values: ['A'], includeBlank: false };
  assert.deepEqual(resolveFilterButtonCells(sheet), [
    { row: 0, column: 0, active: true, sorted: false },
    { row: 0, column: 1, active: false, sorted: false },
  ]);
});

test('Worksheet and Table AutoFilter ownership is singular for overlapping ranges', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const table = { ...sampleTable, autoFilter: createAutoFilterModelForTable(sampleTable) };
  sheet.sheetTables.push(table);
  assert.equal(resolveActiveAutoFilter(sheet)?.range.startRow, 0);
  assert.deepEqual(resolveFilterOwner(sheet), { kind: 'table', tableId: table.id });
  sheet.autoFilter = createAutoFilterModelForTable(sampleTable);
  assert.throws(() => resolveActiveAutoFilter(sheet), /cannot overlap/);
});

test('frontend owner validation rejects same, contained, and partial overlaps while allowing disjoint ranges', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const table = { ...sampleTable, autoFilter: createAutoFilterModelForTable(sampleTable) };
  sheet.sheetTables.push(table);
  for (const range of [
    { sheetId: 's1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    { sheetId: 's1', startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 },
    { sheetId: 's1', startRow: 2, endRow: 5, startColumn: 1, endColumn: 2 },
  ]) {
    const candidate = createAutoFilterModelForTable({ ...sampleTable, range });
    assert.throws(() => validateFilterOwnership(sheet, candidate, { kind: 'worksheet' }), /overlap/);
  }
  const disjointRange = { sheetId: 's1', startRow: 10, endRow: 13, startColumn: 0, endColumn: 1 };
  const disjoint = createAutoFilterModelForTable({ ...sampleTable, range: disjointRange });
  assert.doesNotThrow(() => validateFilterOwnership(sheet, disjoint, { kind: 'worksheet' }));
});

test('frontend table owner validation rejects a worksheet overlap before table filter commit', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const worksheetRange = { sheetId: 's1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 };
  sheet.autoFilter = createAutoFilterModelForTable({ ...sampleTable, range: worksheetRange });
  const table = { ...sampleTable, id: 't2', range: { sheetId: 's1', startRow: 2, endRow: 6, startColumn: 0, endColumn: 1 } };
  sheet.sheetTables.push(table);
  const candidate = createAutoFilterModelForTable(table);
  assert.throws(() => validateFilterOwnership(sheet, candidate, { kind: 'table', tableId: table.id }), /overlap/);
});

test('multiple non-overlapping Table AutoFilters retain independent owners and buttons', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const first = { ...sampleTable, autoFilter: createAutoFilterModelForTable(sampleTable) };
  const second = {
    ...sampleTable,
    id: 't2',
    name: 'Other',
    range: { sheetId: 's1', startRow: 10, endRow: 13, startColumn: 0, endColumn: 1 },
  };
  second.autoFilter = createAutoFilterModelForTable(second);
  sheet.sheetTables.push(first, second);
  assert.equal(resolveAutoFilters(sheet).length, 2);
  assert.deepEqual(resolveFilterButtonCells(sheet), [
    { row: 0, column: 0, active: false, sorted: false }, { row: 0, column: 1, active: false, sorted: false },
    { row: 10, column: 0, active: false, sorted: false }, { row: 10, column: 1, active: false, sorted: false },
  ]);
});
