import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';
import { isFormulaError } from './values';
import { isSpillChild, resolveSpill, spillValueAt } from './spill-resolver';

test('resolveSpill detects blockers and ok states', () => {
  const ok = resolveSpill({
    sheetId: 'Sheet1',
    anchor: { row: 0, column: 0 },
    values: [[1, 2], [3, 4]],
    rowCount: 10,
    columnCount: 10,
    isOccupied: () => false,
  });
  assert.equal(ok.state, 'ok');
  assert.equal(ok.range.endRow, 1);
  assert.equal(ok.range.endColumn, 1);

  const blocked = resolveSpill({
    sheetId: 'Sheet1',
    anchor: { row: 0, column: 0 },
    values: [[1, 2], [3, 4]],
    rowCount: 10,
    columnCount: 10,
    isOccupied: (row, column) => row === 0 && column === 1,
  });
  assert.equal(blocked.state, 'blocked');
  assert.deepEqual(blocked.blocker, { row: 0, column: 1 });
});

test('resolveSpill checks static blocker ranges without materializing every covered cell', () => {
  const blocked = resolveSpill({
    sheetId: 'Sheet1',
    anchor: { row: 0, column: 0 },
    values: [[1, 2], [3, 4]],
    rowCount: 10,
    columnCount: 10,
    isOccupied: () => false,
    blockedRanges: [{ startRow: 0, endRow: 1, startColumn: 1, endColumn: 3 }],
  });
  assert.equal(blocked.state, 'blocked');
  assert.deepEqual(blocked.blocker, { row: 0, column: 1 });

  const anchorOnly = resolveSpill({
    sheetId: 'Sheet1',
    anchor: { row: 0, column: 0 },
    values: [[1, 2], [3, 4]],
    rowCount: 10,
    columnCount: 10,
    isOccupied: () => false,
    blockedRanges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
  });
  assert.equal(anchorOnly.state, 'ok');
});

test('isSpillChild excludes anchor cell', () => {
  const spill = {
    sheetId: 'Sheet1',
    anchor: { row: 0, column: 0 },
    range: { sheetId: 'Sheet1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    values: [[1, 2], [3, 4]],
    state: 'ok' as const,
  };
  assert.equal(isSpillChild(spill, 0, 0), false);
  assert.equal(isSpillChild(spill, 1, 1), true);
});

test('FormulaEngine tracks spill ranges and child values', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', {
    rowCount: 20,
    columnCount: 20,
    isOccupied: () => false,
  });
  engine.setFormula('A1', '=SEQUENCE(2,2,1,1)');
  const spills = engine.getSpillsForSheet('Sheet1');
  assert.equal(spills.length, 1);
  assert.equal(spills[0]?.state, 'ok');
  assert.equal(engine.getCellValue({ sheetId: 'Sheet1', row: 0, column: 0 }), 1);
  assert.equal(engine.getSpillValueAt('Sheet1', 1, 1), 4);
  assert.equal(engine.getCellValue({ sheetId: 'Sheet1', row: 1, column: 1 }), 4);
});

test('FormulaEngine returns #SPILL! when spill area is blocked', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('B1', 'blocker');
  engine.setSpillEnvironment('Sheet1', {
    rowCount: 20,
    columnCount: 20,
    isOccupied: (row, column) => row === 0 && column === 1,
  });
  const result = engine.setFormula('A1', '=SEQUENCE(2,2,1,1)').value;
  assert.ok(isFormulaError(result));
  if (!isFormulaError(result)) throw new Error('expected spill error');
  assert.equal(result.code, '#SPILL!');
});

test('blocked spill results never project matrix values into child cells', () => {
  const spill = resolveSpill({
    sheetId: 'Sheet1',
    anchor: { row: 0, column: 0 },
    values: [[1, 2], [3, 4]],
    rowCount: 10,
    columnCount: 10,
    isOccupied: () => false,
    blockedRanges: [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }],
  });

  const anchor = spillValueAt(spill, 0, 0);
  assert.ok(isFormulaError(anchor));
  if (!isFormulaError(anchor)) throw new Error('expected #SPILL! at the formula anchor');
  assert.equal(anchor.code, '#SPILL!');
  assert.equal(spillValueAt(spill, 0, 1), undefined);
  assert.equal(spillValueAt(spill, 1, 0), undefined);
});

test('calculation snapshots preserve spill geometry and exclude the spill formula own projection', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', {
    rowCount: 20,
    columnCount: 20,
    isOccupied: () => false,
    getBlockedRanges: () => [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }],
  });
  engine.setFormula('A1', '=SEQUENCE(2,2,1,1)');
  const snapshot = engine.exportCalculationSnapshot();
  assert.deepEqual(snapshot.spillSpaces[0]?.blockedRanges, [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }]);
  assert.equal(snapshot.spillSpaces[0]?.spills[0]?.state, 'blocked');

  const restored = FormulaEngine.fromCalculationSnapshot(snapshot);
  await restored.recalculateAsync([{ sheetId: 'Sheet1', row: 0, column: 0 }]);
  const restoredSpill = restored.getSpillsForSheet('Sheet1')[0];
  assert.equal(restoredSpill?.state, 'blocked');
  assert.deepEqual(restoredSpill?.blocker, { row: 0, column: 1 });
});

test('recalculating a persisted spill does not treat its old children as blockers', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', { rowCount: 20, columnCount: 20, isOccupied: () => false });
  engine.setFormula('A1', '=SEQUENCE(2,2,1,1)');

  const restored = FormulaEngine.fromCalculationSnapshot(engine.exportCalculationSnapshot());
  await restored.recalculateAsync([{ sheetId: 'Sheet1', row: 0, column: 0 }]);

  assert.equal(restored.getSpillsForSheet('Sheet1')[0]?.state, 'ok');
  assert.equal(restored.getCellValue('B2'), 4);
});

test('affected spill owners do not block one another with stale ranges during one recalculation', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('C1', 2);
  engine.setValue('D3', 2);
  engine.setFormula('B1', '=SEQUENCE(C1,1)');
  engine.setFormula('A3', '=SEQUENCE(1,D3)');
  assert.equal(engine.getSpillsForSheet('Sheet1').every(({ state }) => state === 'ok'), true);

  engine.setRecalculationMode('manual');
  engine.setValue('C1', 3);
  engine.setValue('D3', 1);
  await engine.recalculateAsync([
    { sheetId: 'Sheet1', row: 0, column: 2 },
    { sheetId: 'Sheet1', row: 2, column: 3 },
  ]);

  const growingSpill = engine.getSpillsForSheet('Sheet1').find(({ anchor }) => anchor.row === 0 && anchor.column === 1);
  assert.equal(growingSpill?.state, 'ok');
  assert.equal(engine.getCellValue('B3'), 3);
});

test('calculation snapshot rejects a persisted spill without an authored formula anchor', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', { rowCount: 20, columnCount: 20, isOccupied: () => false });
  engine.setFormula('A1', '=SEQUENCE(2,2,1,1)');
  const snapshot = engine.exportCalculationSnapshot();
  const malformed = {
    ...snapshot,
    cells: snapshot.cells.filter(({ address }) => address.row !== 0 || address.column !== 0),
  };

  assert.throws(() => FormulaEngine.fromCalculationSnapshot(malformed), /spill has no formula anchor/i);
});
