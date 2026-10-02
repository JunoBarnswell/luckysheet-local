import assert from 'node:assert/strict';
import test from 'node:test';
import { FormulaEngine, rewriteSheetLifecycleFormula, type ExternalCalculationLink, assertFormulaCalculationSnapshot } from './index';

const before = ['First', 'Middle', 'Last', 'Outside'].map(name => ({ id: name, name }));
test('3D endpoints shrink on deletion and crossing movement; external syntax is opaque', () => {
  const after = before.filter(sheet => sheet.id !== 'First');
  const change = { before, after, sheetId: 'First', kind: 'delete' as const };
  assert.equal(rewriteSheetLifecycleFormula('=SUM(First:Last!A1)', change), '=SUM(Middle:Last!A1)');
  assert.equal(rewriteSheetLifecycleFormula('=SUM(Last:First!A1)', change), '=SUM(Last:Middle!A1)');
  assert.equal(rewriteSheetLifecycleFormula('=First!A1+SUM(First:Last!A1)', change), '=#REF!+SUM(Middle:Last!A1)');
  assert.equal(rewriteSheetLifecycleFormula('=[Book.xlsx]First!A1', change), '=[Book.xlsx]First!A1');
  assert.equal(rewriteSheetLifecycleFormula('=SUM(First:Last!A1)', { before, after: [before[1]!, before[2]!, before[3]!, before[0]!], sheetId: 'First', kind: 'move' }), '=SUM(Middle:Last!A1)');
});

const link = (revision: number, value: number): ExternalCalculationLink => ({ id: 'link', token: 'Source.xlsx', sourceUnitId: 'source', subject: 'reader', sourceRevision: revision, accessRevision: 1, state: 'connected', sheets: [{ id: 'stable-sales', name: 'Sales', rowCount: 1000, columnCount: 10 }], cells: [{ address: { sheetId: 'stable-sales', row: 1, column: 1 }, value }] });
test('external revisions calculate in the canonical engine and survive Worker snapshots', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Target' });
  engine.setFormula('A1', '=SUM([Source.xlsx]Sales!B2:B1000)');
  engine.applyExternalCalculationLinks([link(1, 10)]);
  await engine.recalculateAsync();
  assert.equal(engine.getCellResult('A1')?.value, 10);
  const snapshot = engine.exportCalculationSnapshot();
  assertFormulaCalculationSnapshot(snapshot);
  const worker = FormulaEngine.fromCalculationSnapshot(snapshot);
  worker.applyExternalCalculationLinks([link(2, 25)]);
  await worker.recalculateAsync();
  assert.equal(worker.getCellResult('A1')?.value, 25);
  worker.applyExternalCalculationLinks([link(1, 2)]);
  await worker.recalculateAsync();
  assert.equal(worker.getCellResult('A1')?.value, 25);
});
test('authorization revocation clears external values and hidden ranges return BLOCKED', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Target' });
  engine.setFormula('A1', '=[Source.xlsx]Sales!B2');
  engine.applyExternalCalculationLinks([link(1, 10)]);
  await engine.recalculateAsync();
  engine.applyExternalCalculationLinks([{ ...link(1, 10), accessRevision: 2, state: 'denied', sheets: [], cells: [] }]);
  await engine.recalculateAsync();
  assert.equal((engine.getCellResult('A1')?.value as { code: string }).code, '#BLOCKED!');
  assert.equal(engine.getExternalCalculationLinks()[0]?.cells.length, 0);
  assert.throws(() => engine.applyExternalCalculationLinks([{ ...link(1, 10), state: 'denied' }]), /REVOKED_CACHE/);
  assert.throws(() => engine.applyExternalCalculationLinks([{ ...link(1, 10), blockedRanges: [{ sheetId: 'stable-sales', startRow: -1, endRow: 1, startColumn: 0, endColumn: 2 }] }]), /BLOCKED_RANGE/);
});

test('external source rebinding and sparse conditional reads retain authorized semantics', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Target' });
  engine.setFormula('A1', '=SUMIF([Source.xlsx]Sales!B2:B10,">15")');
  engine.applyExternalCalculationLinks([link(5, 25)]);
  await engine.recalculateAsync();
  assert.equal(engine.getCellResult('A1')?.value, 25);
  engine.applyExternalCalculationLinks([{ ...link(1, 40), sourceUnitId: 'new-source' }]);
  await engine.recalculateAsync();
  assert.equal(engine.getCellResult('A1')?.value, 40);
  engine.applyExternalCalculationLinks([{ ...link(2, 40), sourceUnitId: 'new-source', blockedRanges: [{ sheetId: 'stable-sales', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }] }]);
  await engine.recalculateAsync();
  assert.equal((engine.getCellResult('A1')?.value as { code: string }).code, '#BLOCKED!');
});

test('Record owner conflicts are rejected even when the retained owner map is unchanged', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Source' });
  const owner = { tableId: 'Orders', recordId: 'o1', fieldId: 'amount', address: { sheetId: 'Source', row: 1, column: 4 } };
  engine.setRecordFormulaOwners([owner]);
  assert.throws(() => engine.setRecordFormulaOwners([owner, owner]), /OWNER_CONFLICT/);
  assert.deepEqual(engine.getRecordFormulaOwners(), [owner]);
});
