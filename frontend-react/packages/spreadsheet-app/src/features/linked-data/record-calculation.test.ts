import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel, canonicalRecordFieldFormula, assertRecordRelationship, type WorkbookTableModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerRecordCommands } from './commands';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { synchronizeRecordCalculations, recordSheetTables } from './record-calculation';
import { createRemoteReadySessionFixture } from '../../session-test-fixtures';
import { hydrateRuntime } from '../../runtime';

function fixture() {
  const workbook = new WorkbookModel('records', 'Records');
  const orders = workbook.getSheet(workbook.primarySheetId); orders.name = 'Orders';
  const customers = workbook.addSheet('customers-sheet', 'Customers');
  [['ID', 'Customer', 'Quantity', 'Price', 'Amount'], ['o1', 'c1', 2, 10, null], ['o2', 'c1', 3, 20, null], ['o3', 'c2', 4, 5, null]].forEach((row, r) => row.forEach((value, c) => orders.cells.set(r, c, { value })));
  [['ID', 'Name', 'Amounts', 'Total'], ['c1', 'Alpha', null, null], ['c2', 'Beta', null, null]].forEach((row, r) => row.forEach((value, c) => customers.cells.set(r, c, { value })));
  const table = (id: string, sheetId: string, names: string[], count: number): WorkbookTableModel => ({ id, name: id, sourceSheetId: sheetId, sourceRange: { sheetId, startRow: 0, endRow: count, startColumn: 0, endColumn: names.length - 1 }, rowCount: count, recordIdFieldId: `${id}-id`, fields: names.map((name, ordinal) => ({ id: `${id}-${name.toLowerCase()}`, name, ordinal, type: ordinal < 2 ? 'text' : 'number' })), blockSize: 1024, blocks: [], revision: 0 });
  const orderTable = table('Orders', orders.id, ['ID', 'Customer', 'Quantity', 'Price', 'Amount'], 3);
  orderTable.fields[4]!.calculation = { kind: 'formula', formula: canonicalRecordFieldFormula(orderTable, '=[Quantity]*[Price]') };
  workbook.addTable(orderTable); workbook.addTable(table('Customers', customers.id, ['ID', 'Name', 'Amounts', 'Total'], 2));
  const relation = { id: 'orders-customers', fromTableId: 'Orders', fromFieldId: 'Orders-customer', toTableId: 'Customers', toFieldId: 'Customers-id', cardinality: 'many-to-one' as const };
  assertRecordRelationship(workbook, relation); workbook.dataModel.relationships.set(relation.id, relation);
  workbook.getTable('Customers').fields[2]!.calculation = { kind: 'lookup', relationshipId: relation.id, targetFieldId: 'Orders-amount', direction: 'reverse' };
  workbook.getTable('Customers').fields[3]!.calculation = { kind: 'rollup', relationshipId: relation.id, targetFieldId: 'Orders-amount', direction: 'reverse', aggregate: 'SUM' };
  const engine = new FormulaEngine({ defaultSheetId: orders.id, sheetOrder: workbook.getSheets().map(({ id, name }) => ({ id, name })) });
  engine.setSheetTables(recordSheetTables(workbook));
  for (const sheet of workbook.getSheets()) sheet.cells.forEach((cell, row, column) => engine.setValue({ sheetId: sheet.id, row, column }, cell.value ?? null));
  synchronizeRecordCalculations(engine, workbook);
  return { workbook, orders, customers, engine };
}

test('Record formulas, reverse Lookup and Rollup share one dependency graph and native persistence', async () => {
  const { workbook, orders, engine } = fixture();
  await engine.recalculateAsync();
  assert.equal(engine.getRecordFieldResult('Orders', 'o1', 'Orders-amount')?.value, 20);
  assert.deepEqual(engine.getRecordFieldResult('Customers', 'c1', 'Customers-amounts')?.value, [[20], [60]]);
  assert.equal(engine.getRecordFieldResult('Customers', 'c1', 'Customers-total')?.value, 80);
  engine.setValue({ sheetId: orders.id, row: 1, column: 2 }, 5);
  assert.equal(engine.getRecordFieldResult('Customers', 'c1', 'Customers-total')?.value, 110);
  assert.deepEqual(WorkbookModel.fromSnapshot(workbook.snapshot()).getDataModel(), workbook.getDataModel());
  assert.equal(orders.cells.getWithoutHydration(1, 4)?.formula, undefined);
});
test('stable Record/Field writes reject missing relationship targets without modifying the model', () => {
  const { workbook, orders } = fixture();
  const commands = new CommandRuntime(workbook); registerRecordCommands(commands);
  const before = workbook.snapshot();
  assert.throws(() => commands.execute('record.set', { tableId: 'Orders', recordId: 'o1', fieldId: 'Orders-customer', value: 'missing' }), /missing/i);
  assert.deepEqual(workbook.snapshot(), before);
  assert.throws(() => commands.execute('record.set', { tableId: 'Orders', recordId: 'o1', fieldId: 'Orders-amount', value: 123 }), /read.only/i);
  commands.execute('record.set', { tableId: 'Orders', recordId: 'o2', fieldId: 'Orders-quantity', value: 7 });
  assert.equal(orders.cells.getWithoutHydration(2, 2)?.value, 7);
});

test('3D endpoint deletion has reversible formula facts across canonical command undo and redo', () => {
  const workbook = new WorkbookModel('lifecycle', 'Lifecycle');
  const first = workbook.getSheet(workbook.primarySheetId); first.name = 'Jan';
  workbook.addSheet('feb', 'Feb'); workbook.addSheet('mar', 'Mar');
  const summary = workbook.addSheet('summary', 'Summary');
  summary.cells.set(0, 0, { value: null, formula: '=SUM(Jan:Mar!A1)' });
  summary.cells.set(1, 0, { value: null, formula: '=Jan!A1' });
  const commands = new CommandRuntime(workbook); registerSheetCommands(commands);
  commands.execute('sheet.remove', { id: first.id });
  assert.equal(summary.cells.get(0, 0)?.formula, '=SUM(Feb:Mar!A1)');
  assert.equal(summary.cells.get(1, 0)?.formula, '=#REF!');
  commands.undo();
  assert.equal(summary.cells.get(0, 0)?.formula, '=SUM(Jan:Mar!A1)');
  assert.equal(summary.cells.get(1, 0)?.formula, '=Jan!A1');
  commands.redo();
  assert.equal(summary.cells.get(0, 0)?.formula, '=SUM(Feb:Mar!A1)');
});

test('a committed deletion ACK recalculates surviving owners without addressing the removed sheet', async () => {
  const app = createRemoteReadySessionFixture();
  try {
    const runtime = app['runtime'];
    const first = runtime.model.getSheet(runtime.model.primarySheetId); first.name = 'Jan';
    first.cells.set(0, 0, { value: 10 });
    runtime.model.addSheet('feb', 'Feb').cells.set(0, 0, { value: 20 });
    runtime.model.addSheet('mar', 'Mar').cells.set(0, 0, { value: 30 });
    runtime.model.addSheet('summary', 'Summary').cells.set(0, 0, { value: null, formula: '=SUM(Jan:Mar!A1)' });
    hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
    let facts: import('@react-sheets/command-runtime').MutationInfo | undefined;
    const detach = runtime.commands.onMutation((mutation, source) => { if (source === 'command') facts = mutation; });
    const result = runtime.commands.execute('sheet.remove', { id: first.id }); detach();
    assert.ok(facts);
    runtime.commands.applyCommittedStructuralPatches(result.operationId, [{ ...facts, structuralDefinedNameOwnerDeltas: [], structuralRangeOwnerDeltas: [] }], 1);
    await app.waitForFormulaCalculation();
    assert.equal(runtime.formula.getCellResult({ sheetId: 'summary', row: 0, column: 0 })?.value, 50);
  } finally { app.dispose(); }
});

 test('runtime loading preserves calculated field owners over empty physical source cells', async () => {
  const { workbook } = fixture();
  const app = createRemoteReadySessionFixture();
  try {
    const runtime = app['runtime'];
    hydrateRuntime(runtime, { snapshot: workbook.snapshot(), revision: 0 });
    await app.waitForFormulaCalculation();
    assert.equal(runtime.formula.getRecordFieldResult('Orders', 'o2', 'Orders-amount')?.value, 60);
    assert.equal(runtime.formula.getRecordFieldResult('Customers', 'c1', 'Customers-total')?.value, 80);
  } finally { app.dispose(); }
});
