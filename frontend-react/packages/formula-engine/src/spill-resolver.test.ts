import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';
import { isFormulaError } from './values';
import { isSpillChild, resolveSpill } from './spill-resolver';

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
