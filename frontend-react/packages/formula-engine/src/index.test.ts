import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FormulaEngine,
  RangeIndex,
  formatCellAddress,
  isFormulaError,
  lexFormula,
  parseCellAddress,
  parseFormula,
  type CellAddress,
  type FormulaValue,
} from './index';

test('lexer and parser produce a precedence-aware AST without executable code', () => {
  const tokens = lexFormula('=Sheet1!$A$1:Sheet1!B2');
  assert.deepEqual(tokens.map((token) => token.kind), [
    'identifier',
    'bang',
    'identifier',
    'colon',
    'identifier',
    'bang',
    'identifier',
    'eof',
  ]);

  const ast = parseFormula('=1 + 2 * (3 - 1)');
  assert.equal(ast.type, 'binary-expression');
  if (ast.type !== 'binary-expression') throw new Error('Expected binary AST');
  assert.equal(ast.operator, '+');
  assert.equal(ast.right.type, 'binary-expression');
  assert.equal(parseFormula('"line\\nvalue"').type, 'string-literal');
});

test('A1 addresses support zero-based engine coordinates and qualified sheets', () => {
  assert.deepEqual(parseCellAddress("'Annual Plan'!$C$4"), { sheetId: 'Annual Plan', row: 3, column: 2 });
  assert.equal(formatCellAddress({ sheetId: 'Sheet1', row: 0, column: 0 }), 'A1');
  assert.equal(formatCellAddress({ sheetId: 'Annual Plan', row: 3, column: 2 }, true), "'Annual Plan'!C4");
});

test('RangeIndex tracks direct and rectangular dependencies', () => {
  const index = new RangeIndex();
  const owner = address('Sheet1', 0, 3);
  index.set(owner, [
    { kind: 'cell', address: address('Sheet1', 0, 0) },
    {
      kind: 'range',
      start: address('Sheet1', 1, 0),
      end: address('Sheet1', 2, 2),
    },
  ]);

  assert.deepEqual(index.getDependents(address('Sheet1', 2, 1)), [owner]);
  assert.deepEqual(index.getDependents(address('Sheet1', 4, 4)), []);
  assert.equal(index.getDependencies(owner).length, 2);
  assert.equal(index.remove(owner), true);
  assert.deepEqual(index.getDependents(address('Sheet1', 0, 0)), []);
});

test('FormulaEngine evaluates arithmetic, strings, SUM ranges, and exposes dependencies', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 2);
  engine.setValue('A2', 3);
  engine.setValue('A3', 'ignored');

  const result = engine.setFormula('C1', '=SUM(A1:A3) * 2');
  assert.equal(result.value, 10);
  assert.deepEqual(result.dependencies, [
    {
      kind: 'range',
      start: { sheetId: 'Sheet1', row: 0, column: 0 },
      end: { sheetId: 'Sheet1', row: 2, column: 0 },
    },
  ]);
  assert.deepEqual(engine.getDependents('A2').map((cell) => formatCellAddress(cell, true)), ['Sheet1!C1']);

  engine.setValue('A2', 4);
  assert.equal(engine.getCellValue('C1'), 12);
  assert.equal(engine.getCellValue('A3'), 'ignored');
});

test('FormulaEngine recalculates a dependency chain and updates replaced dependencies', () => {
  const engine = new FormulaEngine();
  engine.setFormula('B1', '=A1 + 1');
  engine.setFormula('C1', '=B1 * 2');
  engine.setValue('A1', 4);
  assert.equal(engine.getCellValue('B1'), 5);
  assert.equal(engine.getCellValue('C1'), 10);

  engine.setValue('D1', 10);
  engine.setFormula('E1', '=A1');
  engine.setFormula('E1', '=D1');
  engine.setValue('A1', 100);
  assert.equal(engine.getCellValue('E1'), 10);
  engine.setValue('D1', 11);
  assert.equal(engine.getCellValue('E1'), 11);

  const report = engine.clearCell('D1');
  assert.deepEqual(report.recalculated.map((cell) => formatCellAddress(cell, true)), ['Sheet1!E1']);
  assert.equal(engine.getCellValue('E1'), null);
});

test('FormulaEngine returns explicit errors and propagates them through formulas', () => {
  const engine = new FormulaEngine();
  const division = engine.setFormula('A1', '=1 / 0').value;
  const propagated = engine.setFormula('B1', '=A1 + 1').value;
  const stringArithmetic = engine.setFormula('C1', '="text" + 1').value;
  const unknownFunction = engine.setFormula('D1', '=NO_SUCH_FUNCTION(1)').value;
  const parseError = engine.setFormula('E1', '=1 +').value;

  assertError(division, '#DIV/0!');
  assertError(propagated, '#DIV/0!');
  assertError(stringArithmetic, '#VALUE!');
  assertError(unknownFunction, '#NAME?');
  assertError(parseError, '#PARSE!');
});

test('FormulaEngine supports qualified references and detects cycles', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue({ sheetId: 'Sheet2', row: 0, column: 0 }, 7);
  assert.equal(engine.setFormula('A1', '=Sheet2!A1 + 1').value, 8);

  const firstCycle = engine.setFormula('B1', '=C1 + 1').value;
  const secondCycle = engine.setFormula('C1', '=B1 + 1').value;
  assert.equal(firstCycle, 1);
  assertError(secondCycle, '#CYCLE!');
  assertError(engine.getCellValue('B1'), '#CYCLE!');
});

function address(sheetId: string, row: number, column: number): CellAddress {
  return { sheetId, row, column };
}

function assertError(value: FormulaValue, code: '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#PARSE!' | '#CYCLE!'): void {
  assert.equal(isFormulaError(value), true);
  if (!isFormulaError(value)) throw new Error('Expected a formula error');
  assert.equal(value.code, code);
}
