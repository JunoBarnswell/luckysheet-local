import assert from 'node:assert/strict';
import test from 'node:test';
import { getWorkbookObjectPort, WorkbookSession } from '@react-sheets/spreadsheet-app';
import { Workbook } from './workbook';
import { SdkError } from '../error';
function fixture(id: string, scope: object = {}) {
  const session = new WorkbookSession({ unitId: id });
  const workbook = new Workbook(getWorkbookObjectPort(session), scope, () => session.dispose());
  return { session, workbook, sheet: session['runtime'].model.getSheet(session.getActiveSheetId()) };
}
const invalid = (cause: unknown) => cause instanceof SdkError && cause.code === 'INVALID_ARGUMENT' && Boolean(cause.object?.workbookId) && Boolean(cause.recovery);
test('object cells own explicit addresses, preserve literal values/styles and share the canonical formula engine', async () => {
  const { session, workbook, sheet } = fixture('object-1');
  try {
    const worksheet = workbook.worksheets.at(0);
    assert.equal(workbook.worksheets.byId(worksheet.id), worksheet);
    assert.equal(workbook.worksheets.byName(worksheet.name.toLowerCase()), worksheet);
    const cell = worksheet.cells.get('$D$8');
    assert.equal(cell, worksheet.cells.get('d8'));
    sheet.cells.set(7, 3, { value: 9, style: { bold: true } });
    session.selectAddress('A1');
    await cell.setValue('=1+2');
    assert.equal((await cell.read()).value, '=1+2');
    assert.equal((await cell.read()).formula, undefined);
    assert.equal(sheet.cells.get(7, 3)?.style?.bold, true);
    await cell.setFormula('=SUM(1,2)');
    assert.equal((await cell.read()).calculatedValue, 3);
    assert.equal(sheet.cells.get(7, 3)?.formula, '=SUM(1,2)');
    assert.deepEqual([...worksheet.ranges.get('D8:E9').cells()].map(cell => cell.address), ['D8', 'E8', 'D9', 'E9']);
    for (const forbidden of ['session', 'runtime', 'model', 'dispatch', 'api', 'token', 'getAccessToken']) assert.equal(forbidden in workbook, false);
    assert.equal(Object.isFrozen(await cell.read()), true);
  } finally { workbook.close(); }
});
test('invalid object addresses and values reject without changes', async () => {
  const { session, workbook } = fixture('object-invalid');
  try {
    const sheet = workbook.worksheets.at(0), before = session['runtime'].model.snapshot();
    for (const address of ['A0', 'A1048577', 'XFE1', 'A$$1', 'Sheet1!A1', 'A1:B2']) assert.throws(() => sheet.cells.get(address), invalid);
    assert.throws(() => sheet.ranges.get('B2:A1'), invalid);
    assert.throws(() => workbook.worksheets.at(-1), invalid);
    await assert.rejects(sheet.cells.get('D8').setValue(Infinity), invalid);
    await assert.rejects(sheet.cells.get('D8').setFormula('SUM(1,2)'), invalid);
    assert.deepEqual(session['runtime'].model.snapshot(), before);
  } finally { workbook.close(); }
});
test('object permission checks reject viewer writes and hidden-range reads without mutating canonical cells', async () => {
  const { session, workbook, sheet } = fixture('object-permission');
  try {
    const cell = workbook.worksheets.at(0).cells.get('D8');
    const before = session['runtime'].model.snapshot();
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 1, regions: [] });
    await assert.rejects(cell.setValue(4), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 2,
      regions: [{ range: { sheetId: sheet.id, startRow: 7, endRow: 7, startColumn: 3, endColumn: 3 }, access: 'hidden' }] });
    await assert.rejects(cell.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), before);
  } finally { workbook.close(); }
});
test('closing one workbook retires its cell and range handles while another stays usable', async () => {
  const scope = {}, first = fixture('first', scope), second = fixture('second', scope);
  try {
    const cell = first.workbook.worksheets.at(0).cells.get('D8');
    const range = first.workbook.worksheets.at(0).ranges.get('A1:B2');
    first.workbook.close(); first.workbook.close();
    await assert.rejects(cell.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    assert.throws(() => [...range.cells()], (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    await second.workbook.worksheets.at(0).cells.get('D8').setValue(42);
    assert.equal((await second.workbook.worksheets.at(0).cells.get('D8').read()).value, 42);
  } finally { second.workbook.close(); }
});
test('cross-workbook formulas refresh authorized revisions and clear revoked inputs', async () => {
  const scope = {}, source = fixture('object-source', scope), target = fixture('object-target', scope);
  try {
    const runtime = target.session['runtime'];
    let revision = 1, denied = false;
    runtime.api.getExternalLinkInputs = async (_unit, linkId) => {
      if (denied) throw Object.assign(new Error('Source access revoked'), { status: 403, code: 'FORBIDDEN' });
      const binding = runtime.model.dataModel.externalLinks.get(linkId)!;
      return { binding: structuredClone(binding), snapshot: source.session['runtime'].model.snapshot(), subject: 'object-owner', sourceRevision: revision, accessRevision: 0, blockedRanges: [] };
    };
    const sourceSheet = source.workbook.worksheets.at(0);
    await sourceSheet.cells.get('D8').setValue(10);
    await sourceSheet.cells.get('D9').setValue(20);
    await target.workbook.externalLinks.bind(source.workbook, 'Source.xlsx');
    const cell = target.workbook.worksheets.at(0).cells.get('D8');
    await cell.setFormula(`=SUM('[Source.xlsx]${sourceSheet.name}'!D8:D9)`);
    await target.workbook.externalLinks.refresh();
    assert.equal((await cell.read()).calculatedValue, 30);
    await sourceSheet.cells.get('D9').setValue(40); revision++;
    assert.equal((await target.workbook.externalLinks.refresh())[0]?.sourceRevision, 2);
    assert.equal((await cell.read()).calculatedValue, 50);
    denied = true;
    assert.equal((await target.workbook.externalLinks.refresh())[0]?.state, 'denied');
    assert.equal(runtime.formula.getExternalCalculationLinks()[0]?.cells.length, 0);
    assert.equal(((await cell.read()).calculatedValue as { code: string }).code, '#BLOCKED!');
    assert.equal((await cell.read()).formula, `=SUM('[Source.xlsx]${sourceSheet.name}'!D8:D9)`);
  } finally { source.workbook.close(); target.workbook.close(); }
});
test('cross-scope, self and invalid-token binding refuse before touching target definitions', async () => {
  const first = fixture('bind-first'), second = fixture('bind-second');
  try {
    const before = first.session['runtime'].model.snapshot();
    for (const [source, token] of [[second.workbook, 'Source.xlsx'], [first.workbook, 'Self.xlsx'], [first.workbook, '../bad']] as const) {
      await assert.rejects(first.workbook.externalLinks.bind(source, token), invalid);
      assert.deepEqual(first.session['runtime'].model.snapshot(), before);
    }
  } finally { first.workbook.close(); second.workbook.close(); }
});
