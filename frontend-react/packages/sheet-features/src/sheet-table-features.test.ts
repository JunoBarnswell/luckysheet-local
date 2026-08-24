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
  resolveFilterButtonCells,
  resolveActiveAutoFilter,
  resolveFilterOwner,
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
  showFilterButton: true,
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
  assert.deepEqual(resolveFilterButtonCells(sheet), [
    { row: 0, column: 0 },
    { row: 0, column: 1 },
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
