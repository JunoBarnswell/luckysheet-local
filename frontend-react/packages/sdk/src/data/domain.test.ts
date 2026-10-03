import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookSession } from '@react-sheets/spreadsheet-app';
import { validateSheetTableModel } from '@react-sheets/sheet-features';
import { WorkbookModel } from '@react-sheets/core-model';
import { hydrateRuntime } from '../../../spreadsheet-app/src/runtime';
import { DataDomain } from './domain';

function fixture(values: unknown[][]) {
  const session = new WorkbookSession();
  const sheetId = session.getActiveSheetId();
  session.runCommand('sheet.range.set', { sheetId, startRow: 0, startColumn: 0, values: values.map(row => row.map(value => ({ value }))) });
  session.selectAddress(`A1:B${values.length}`);
  const domain = new DataDomain(session);
  return { session, actions: domain.actions, domain, sheet: session['runtime'].model.getSheet(sheetId) };
}

test('SDK split writes through the canonical transaction and undo restores empty destinations', async () => {
  const { session, actions, sheet } = fixture([['a,b,c', null], ['1,2,3', null]]);
  try {
    const before = sheet.snapshot();
    assert.deepEqual(await actions.textToColumns({ maxColumns: 3 }), { status: 'applied' });
    assert.equal(sheet.cells.get(0, 2)?.value, 'c');
    assert.equal(sheet.cells.get(1, 2)?.value, '3');
    session.undo();
    assert.deepEqual(sheet.snapshot(), before);
    session.redo(); assert.equal(sheet.cells.get(0, 2)?.value, 'c');
  } finally { session.dispose(); }
});

test('SDK subtotal writes formulas and outline in one history entry with exact undo', async () => {
  const { session, actions, sheet } = fixture([['Group', 'Amount'], ['East', 10], ['East', 5], ['West', 7]]);
  try {
    const before = sheet.snapshot();
    assert.deepEqual(await actions.subtotal(), { status: 'applied' });
    assert.equal(sheet.cells.get(6, 1)?.formula, '=SUBTOTAL(9,B2:B3)');
    await session.waitForFormulaCalculation();
    assert.equal(session['runtime'].formula.getCellValue({ sheetId: sheet.id, row: 6, column: 1 }), 15);
    assert.equal(sheet.outline?.groups.length, 2);
    session.undo(); assert.deepEqual(sheet.snapshot(), before);
    session.redo(); await session.waitForFormulaCalculation();
    assert.equal(session['runtime'].formula.getCellValue({ sheetId: sheet.id, row: 7, column: 1 }), 7);
  } finally { session.dispose(); }
});

test('SDK data rejects invalid inputs and permissions without partial writes', async () => {
  const { session, actions, sheet } = fixture([['Group', 'Amount'], ['East', 10]]);
  try {
    const before = sheet.snapshot();
    for (const result of [await actions.subtotal({ valueColumn: 99 }), await actions.textToColumns({ delimiter: '' }), await actions.removeDuplicates({ columns: [99] })]) {
      assert.equal(result.status, 'rejected');
      if (result.status === 'rejected') { assert.equal(result.error.operation.startsWith('data.'), true); assert.ok(result.error.recovery); }
      assert.deepEqual(sheet.snapshot(), before);
    }
    session['permission'].applyServerAccess({ unitId: session.getUiSnapshot().unitId, role: 'viewer', accessRevision: 1, regions: [] });
    const denied = await actions.toggleFilter();
    assert.equal(denied.status, 'rejected');
    if (denied.status === 'rejected') assert.equal(denied.error.code, 'FORBIDDEN');
    assert.deepEqual(sheet.snapshot(), before);
  } finally { session.dispose(); }
});

test('SDK table filter clear retains its owner and buttons; offline structural toggle refuses', async () => {
  const { session, actions, sheet } = fixture([['Group', 'Amount'], ['East', 10], ['West', 7]]);
  try {
    const sheetId = sheet.id;
    const range = { sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
    const seed = validateSheetTableModel({ id: 't', name: 'Sales', sheetId, range, hasHeaderRow: true, hasTotalRow: false, showFilterButton: true, autoExpand: 'none', showBandedRows: false, showBandedColumns: false, showFirstColumn: false, showLastColumn: false,
      columns: [{ id: 'group', name: 'Group' }, { id: 'amount', name: 'Amount' }],
      autoFilter: undefined }, sheet);
    seed.autoFilter!.columns[0]!.criterion = { kind: 'values', values: ['East'], includeBlank: false };
    sheet.sheetTables.push(seed);
    session['refresh']();
    assert.deepEqual(await actions.clearFilter(), { status: 'applied' });
    const table = sheet.sheetTables[0]!;
    assert.equal(table.showFilterButton, true);
    assert.ok(table.autoFilter);
    assert.equal(table.autoFilter.columns[0]?.criterion, undefined);
    session.undo(); assert.ok(sheet.sheetTables[0]!.autoFilter?.columns[0]?.criterion);
    const beforeToggle = sheet.snapshot();
    const toggle = await actions.toggleFilter();
    assert.equal(toggle.status, 'rejected');
    if (toggle.status === 'rejected') assert.match(toggle.error.message, /STRUCTURAL_PLANNER_OFFLINE/);
    assert.deepEqual(sheet.snapshot(), beforeToggle);
    assert.equal(sheet.autoFilter, undefined);
  } finally { session.dispose(); }
});

test('SDK retired data actions reject without invoking a disposed session', async () => {
  const { session, actions, domain } = fixture([['Group', 'Amount'], ['East', 10]]);
  domain.dispose(); session.dispose();
  const result = await actions.subtotal();
  assert.equal(result.status, 'rejected');
  if (result.status === 'rejected') assert.equal(result.error.code, 'RUNTIME_DISPOSED');
});

test('SDK subtotal targets an explicit inactive worksheet and consumes the one formula engine', async () => {
  const { session, actions } = fixture([['Group', 'Amount'], ['East', 10]]);
  try {
    // Arrange the second worksheet at the canonical snapshot boundary. Actual
    // worksheet creation requires Java and is covered in the browser UAT.
    const prepared = WorkbookModel.fromSnapshot(session['runtime'].model.snapshot());
    prepared.addSheet('other', 'Other');
    hydrateRuntime(session['runtime'], { snapshot: prepared.snapshot(), revision: 0 });
    session.runCommand('sheet.range.set', { sheetId: 'other', startRow: 0, startColumn: 0,
      values: [[{ value: 'Group' }, { value: 'Amount' }], [{ value: 'East' }, { value: 2 }], [{ value: 'East' }, { value: 4 }], [{ value: 'East' }, { value: '7' }]] });
    session.selectSheet(session['runtime'].model.primarySheetId);
    for (const [functionName, expected] of [['COUNTA', 3], ['STDEV', 1.4142135623731], ['VARP', 1], ['PRODUCT', 8]] as const) {
      const result = await actions.subtotal({ range: { sheetId: 'other', address: 'A1:B4' }, functionName });
      assert.equal(result.status, 'applied', JSON.stringify(result)); await session.waitForFormulaCalculation();
      assert.equal(session['runtime'].formula.getCellValue({ sheetId: 'other', row: 6, column: 1 }), expected);
      session.undo();
    }
  } finally { session.dispose(); }
});

test('SDK subtotal rejects malformed function, visibility and explicit address without a partial transaction', async () => {
  const { session, actions, sheet } = fixture([['Group', 'Amount'], ['East', 10]]);
  try {
    const before = sheet.snapshot();
    const invalidInputs: unknown[] = [{ functionName: 'MEDIAN' }, { functionName: null }, { groupColumn: null }, { valueColumn: null }, { excludeHiddenRows: 1 }, { range: { sheetId: sheet.id, address: 'B2:A1' } },
      { range: { sheetId: 'missing', address: 'A1:B2' } }, { range: { sheetId: sheet.id, address: 'A0:B2' } }, { range: { sheetId: sheet.id, address: 'A1:B1048576' } }];
    for (const input of invalidInputs) {
      const result = await actions.subtotal(input as Parameters<typeof actions.subtotal>[0]);
      assert.equal(result.status, 'rejected');
      if (result.status === 'rejected') { assert.equal(result.error.code, 'INVALID_ARGUMENT'); assert.equal(result.error.operation, 'data.subtotal'); assert.equal(result.error.object?.workbookId, session.getUiSnapshot().unitId); }
      assert.deepEqual(sheet.snapshot(), before);
    }
  } finally { session.dispose(); }
});

test('SDK subtotal excludes manual hidden rows on request while ordinary subtotal still reads them', async () => {
  const { session, actions, sheet } = fixture([['Group', 'Amount'], ['East', 2], ['East', 4]]);
  try {
    session.runCommand('sheet.rows.visibility.set', { sheetId: sheet.id, rows: [2], hidden: true });
    const before = sheet.snapshot();
    assert.deepEqual(await actions.subtotal({ excludeHiddenRows: true }), { status: 'applied' });
    await session.waitForFormulaCalculation();
    assert.equal(sheet.cells.get(5, 1)?.formula, '=SUBTOTAL(109,B2:B3)');
    assert.equal(session['runtime'].formula.getCellValue({ sheetId: sheet.id, row: 5, column: 1 }), 2);
    session.undo(); assert.deepEqual(sheet.snapshot(), before);
    assert.deepEqual(await actions.subtotal(), { status: 'applied' }); await session.waitForFormulaCalculation();
    assert.equal(session['runtime'].formula.getCellValue({ sheetId: sheet.id, row: 5, column: 1 }), 6);
  } finally { session.dispose(); }
});
