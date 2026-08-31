import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FormulaEngine,
  FUNCTION_CAPABILITY_MATRIX,
  getFunctionCapability,
  RangeIndex,
  formatFormula,
  formatCellAddress,
  isFormulaError,
  lexFormula,
  parseCellAddress,
  parseFormula,
  offsetAst,
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

test('spill and implicit-intersection references round-trip and offset through the AST', () => {
  assert.deepEqual(lexFormula('=SUM(A2#)').map((token) => token.kind), [
    'identifier', 'left-paren', 'identifier', 'spill-operator', 'right-paren', 'eof',
  ]);
  assert.equal(formatFormula(parseFormula('=SUM(A2#)')), '=SUM(A2#)');
  assert.equal(formatFormula(offsetAst(parseFormula('=SUM(A2#)'), 1, 1)), '=SUM(B3#)');
  assert.equal(formatFormula(offsetAst(parseFormula('=@A2#'), 1, 1)), '=@B3#');
});

test('AST formatter preserves explicit grouping and qualified sheet names', () => {
  const ast = parseFormula("=('Annual Plan'!$A$1+A1)*B1");
  assert.equal(formatFormula(ast), "=('Annual Plan'!$A$1+A1)*B1");
  assert.equal(formatFormula(offsetAst(parseFormula('=A1+$B$1'), 2, 3)), '=D3+$B$1');
});

test('structural deletion invalidates references instead of clamping them', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 10);
  engine.setValue('A2', 20);
  engine.setFormula('B1', '=A2+1');
  const report = engine.remapStructure('Sheet1', { axis: 'row', at: 1, count: 1, op: 'delete' });

  assert.equal(formatFormula(parseFormula('=#REF!')), '=#REF!');
  assert.equal(engine.getCellResult('B1')?.formula, '=#REF!+1');
  assertError(engine.getCellValue('B1'), '#REF!');
  assert.equal(report.recalculated.some((address) => address.row === 0 && address.column === 1), true);
  assert.deepEqual(engine.getDependencies('B1'), []);
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

test('FormulaEngine evaluates comprehensive math, logical, text, and lookup functions', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 10);
  engine.setValue('A2', 20);
  engine.setValue('A3', 30);
  engine.setValue('B1', 'Apple');
  engine.setValue('B2', 'Banana');
  engine.setValue('B3', 'Cherry');

  // IF, AVERAGE, COUNT, MIN, MAX
  assert.equal(engine.setFormula('C1', '=IF(A1 > 5, "YES", "NO")').value, 'YES');
  assert.equal(engine.setFormula('C2', '=AVERAGE(A1:A3)').value, 20);
  assert.equal(engine.setFormula('C3', '=COUNT(A1:A3)').value, 3);
  assert.equal(engine.setFormula('C4', '=MIN(A1:A3)').value, 10);
  assert.equal(engine.setFormula('C5', '=MAX(A1:A3)').value, 30);

  // VLOOKUP, INDEX, MATCH
  assert.equal(engine.setFormula('D1', '=VLOOKUP(20, A1:B3, 2, FALSE)').value, 'Banana');
  assert.equal(engine.setFormula('D2', '=INDEX(B1:B3, 3, 1)').value, 'Cherry');
  assert.equal(engine.setFormula('D3', '=MATCH("Banana", B1:B3, 0)').value, 2);

  // Text functions: CONCAT, LEFT, RIGHT, UPPER, TEXTJOIN
  assert.equal(engine.setFormula('E1', '=CONCAT("Hello", " ", "World")').value, 'Hello World');
  assert.equal(engine.setFormula('E2', '=UPPER(B1)').value, 'APPLE');
  assert.equal(engine.setFormula('E3', '=LEFT(B2, 3)').value, 'Ban');
  assert.equal(engine.setFormula('E4', '=TEXTJOIN(", ", TRUE, B1:B3)').value, 'Apple, Banana, Cherry');

  // Precedence, percent, concat, date, information
  assert.equal(engine.setFormula('F1', '=2 + 3 * 4 ^ 2').value, 50);
  assert.equal(engine.setFormula('F2', '=50% + 10%').value, 0.6);
  assert.equal(engine.setFormula('F3', '="Total: " & A1').value, 'Total: 10');
  assert.equal(engine.setFormula('F4', '=YEAR(DATE(2025, 6, 15))').value, 2025);
  assert.equal(engine.setFormula('F5', '=MONTH(DATE(2025, 6, 15))').value, 6);
  assert.equal(engine.setFormula('F6', '=DAY(DATE(2025, 6, 15))').value, 15);
  assert.equal(engine.setFormula('F7', '=ISNUMBER(A1)').value, true);
  assert.equal(engine.setFormula('F8', '=ISTEXT(B1)').value, true);
  assert.equal(engine.setFormula('F9', '=ISBLANK(Z99)').value, true);
});

test('FormulaEngine resolves spill references as dynamic ranges', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', { rowCount: 20, columnCount: 10, isOccupied: () => false });
  engine.setFormula('A2', '=SEQUENCE(2,1,1,1)');
  const dependent = engine.setFormula('C2', '=SUM(A2#)');
  assert.equal(dependent.value, 3);
  assert.deepEqual(dependent.dependencies, [{ kind: 'cell', address: { sheetId: 'Sheet1', row: 1, column: 0 } }]);
  assert.equal(engine.setFormula('D2', '=@A2#').value, 1);
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

test('FormulaEngine evaluates only the selected lazy branch and tracks dynamic references', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 1);
  assert.equal(engine.setFormula('B1', '=IF(A1,10,1/0)').value, 10);
  assert.equal(engine.setFormula('C1', '=IFERROR(1/0,42)').value, 42);
  engine.setValue('D1', 'A1');
  engine.setFormula('E1', '=INDIRECT(D1)');
  assert.equal(engine.getCellValue('E1'), 1);
  assert.equal(engine.getDependents('A1').some((address) => address.column === 4), true);
  engine.setValue('A1', 7);
  assert.equal(engine.getCellValue('E1'), 7);
});

test('formula capability matrix is derived from executable functions and marks unsupported entries', () => {
  assert.equal(FUNCTION_CAPABILITY_MATRIX.some((entry) => entry.id === 'INDIRECT' && entry.status === 'native'), true);
  assert.equal(getFunctionCapability('ROMAN').status, 'unsupported');
  assert.equal(getFunctionCapability('NO_SUCH_FUNCTION').status, 'unsupported');
});

test('calculation sessions fail closed when a persistent Worker is unavailable', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  assert.throws(() => engine.createCalculationSessionPort(), /CALCULATION_WORKER_UNAVAILABLE/);
});

test('FormulaEngine supports qualified references and detects cycles', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue({ sheetId: 'Sheet2', row: 0, column: 0 }, 7);
  assert.equal(engine.setFormula('A1', '=Sheet2!A1 + 1').value, 8);

  const firstCycle = engine.setFormula('B1', '=C1 + 1').value;
  const secondCycle = engine.setFormula('C1', '=B1 + 1').value;
  assert.equal(firstCycle, 1);
  assertError(secondCycle, '#NUM!');
  assertError(engine.getCellValue('B1'), '#NUM!');
});

function address(sheetId: string, row: number, column: number): CellAddress {
  return { sheetId, row, column };
}

function assertError(
  value: FormulaValue,
  code: '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#PARSE!' | '#SPILL!',
): void {
  assert.equal(isFormulaError(value), true);
  if (!isFormulaError(value)) throw new Error('Expected a formula error');
  assert.equal(value.code, code);
}
