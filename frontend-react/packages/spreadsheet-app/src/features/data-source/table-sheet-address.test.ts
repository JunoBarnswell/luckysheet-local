import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { resolveTableSheetCellAddress, resolveWorkbookViewCell } from './table-sheet-address';
import { buildCanvasSheetSnapshot } from '../../ui-snapshot';

function fixture() {
  const workbook = new WorkbookModel('table-address-proof', 'Proof');
  const source = workbook.getSheet(workbook.primarySheetId);
  [['Name', 'Quantity', 'Amount'], ['Alpha', 2, 20], ['Beta', 3, 60]].forEach((line, row) => line.forEach((value, column) => source.cells.set(row, column, { value })));
  workbook.addTable({ id: 'orders', name: 'Orders', sourceSheetId: source.id, sourceRange: { sheetId: source.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }, rowCount: 2, fields: [{ id: 'name', name: 'Name', ordinal: 0, type: 'text' }, { id: 'quantity', name: 'Quantity', ordinal: 1, type: 'number' }, { id: 'amount', name: 'Amount', ordinal: 2, type: 'number' }], blockSize: 1024, blocks: [], revision: 0 });
  const view = workbook.addAdvancedSheet({ id: 'view', name: 'View', kind: 'table-sheet', tableSheet: { viewId: 'orders', columns: [{ fieldId: 'amount', caption: 'Total' }, { fieldId: 'name', caption: 'Customer' }], grouping: [], sortState: [{ fieldId: 'amount', direction: 'desc' }] } });
  const engine = new FormulaEngine({ defaultSheetId: source.id, sheetOrder: workbook.getSheets().map(({ id, name }) => ({ id, name })) });
  source.cells.forEach((cell, row, column) => engine.setValue({ sheetId: source.id, row, column }, cell.value ?? null));
  return { workbook, source, view, engine };
}

test('visible fields, sorting, header and write target resolve through one address owner', () => {
  const { workbook, source, view, engine } = fixture();
  const canvas = buildCanvasSheetSnapshot(workbook, view, engine, true);
  assert.equal(canvas.getCell(0, 0)?.value, 'Total');
  assert.equal(canvas.getCell(1, 0)?.value, '60');
  assert.equal(canvas.getCell(1, 1)?.value, 'Beta');
  const address = resolveTableSheetCellAddress(workbook, view, engine, 1, 0)!;
  assert.deepEqual(address, { sheetId: source.id, row: 2, column: 2 });
  const commands = new CommandRuntime(workbook);
  registerSheetCommands(commands);
  commands.execute('sheet.range.set', { sheetId: address.sheetId, startRow: address.row, startColumn: address.column, values: [[{ value: 10 }]] });
  engine.setValue(address, 10);
  assert.equal(source.cells.get(2, 2)?.value, 10);
  assert.equal(canvas.getCell(1, 1)?.value, 'Alpha');
  const restored = WorkbookModel.fromSnapshot(workbook.snapshot());
  assert.equal(buildCanvasSheetSnapshot(restored, restored.getSheet(view.id), engine, true).getCell(1, 1)?.value, 'Alpha');
});

test('source cell formulas project as calculated values and affect the shared sort mapping', () => {
  const { workbook, source, view, engine } = fixture();
  source.cells.set(1, 2, { formula: '=B2*100', value: null });
  engine.setFormula({ sheetId: source.id, row: 1, column: 2 }, '=B2*100');
  assert.equal(buildCanvasSheetSnapshot(workbook, view, engine, true).getCell(1, 0)?.value, '200');
  assert.deepEqual(resolveTableSheetCellAddress(workbook, view, engine, 1, 0), { sheetId: source.id, row: 1, column: 2 });
  const resolved = resolveWorkbookViewCell(workbook, view, engine, 1, 0, (owner, row, column) => owner.cells.get(row, column));
  assert.equal(resolved.owner.id, source.id);
  assert.equal(resolved.cell?.formula, '=B2*100');
  assert.equal(resolved.writable, true);
  const header = resolveWorkbookViewCell(workbook, view, engine, 0, 0, (owner, row, column) => owner.cells.get(row, column));
  assert.equal(header.cell?.value, 'Total');
  assert.equal(header.writable, false);
});

test('unsupported calculated fields fail before a view mutation is committed', () => {
  const { workbook, view, engine } = fixture();
  const commands = new CommandRuntime(workbook);
  registerSheetCommands(commands);
  const before = workbook.snapshot();
  const definition = structuredClone(view.tableSheet!);
  definition.columns[0]!.formula = '=[Quantity]*2';
  definition.columns[0]!.type = 'formula';
  assert.throws(() => commands.execute('tableSheet.update', { sheetId: view.id, definition }), /UNSUPPORTED_FEATURE/);
  assert.deepEqual(workbook.snapshot(), before);
  view.tableSheet = definition;
  assert.throws(() => resolveTableSheetCellAddress(workbook, view, engine, 1, 0), /UNSUPPORTED_FEATURE/);
});

test('unknown field and unavailable block address are rejected observably', () => {
  const { workbook, view, engine } = fixture();
  view.tableSheet!.columns[0]!.fieldId = 'missing';
  assert.throws(() => resolveTableSheetCellAddress(workbook, view, engine, 1, 0), /TABLE_SHEET_ADDRESS_INVALID/);
  view.tableSheet!.columns[0]!.fieldId = 'amount';
  delete workbook.getTable('orders').sourceRange;
  assert.throws(() => resolveTableSheetCellAddress(workbook, view, engine, 1, 0), /UNSUPPORTED_FEATURE/);
});
