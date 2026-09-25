import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FormulaEngine,
  RangeIndex,
  collectFormulaDependencies,
  collectFormulaReferenceNodes,
  mapAstMovedReferences,
  mapAstStructuralReferences,
  formatFormula,
  formatCellAddress,
  isFormulaError,
  lexFormula,
  parseCellAddress,
  parseFormula,
  offsetAst,
  type CellAddress,
  type FormulaValue,
  type MoveRangeReferenceTransform,
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

test('structural reference index selects affected owners and retains invalid formulas', () => {
  const index = new RangeIndex([{ id: 'Sheet1', name: 'Sheet1' }]);
  const affected = address('Sheet1', 10, 3);
  const unaffected = address('Sheet1', 11, 3);
  const invalid = address('Sheet1', 20, 3);
  index.set(affected, [{
    kind: 'range',
    start: address('Sheet1', 2, 0),
    end: address('Sheet1', 5, 2),
  }]);
  index.set(unaffected, [{ kind: 'cell', address: address('Sheet1', 1, 0) }]);
  index.set(invalid, [], true);

  assert.deepEqual(index.getStructuralDependents('Sheet1', 'row', 3), [affected]);
  assert.deepEqual(index.getRangeDependents('Sheet1', { startRow: 4, endRow: 6, startColumn: 2, endColumn: 3 }), [affected]);
  assert.throws(
    () => index.getRangeDependents('Sheet1', { startRow: 6, endRow: 4, startColumn: 0, endColumn: 1 }),
    /Reference range query bounds are invalid/,
  );
  assert.deepEqual(index.getInvalidFormulaOwners(), [invalid]);

  index.set(invalid, [{ kind: 'cell', address: address('Sheet1', 4, 0) }]);
  assert.deepEqual(index.getInvalidFormulaOwners(), []);
});

test('structural-only formula sources coexist with calculation owners and retain invalid source failures', () => {
  const index = new RangeIndex([{ id: 'Sheet1', name: 'Sheet1' }]);
  const owner = address('Sheet1', 8, 5);
  const calculationTarget = address('Sheet1', 1, 0);
  const structuralTarget = address('Sheet1', 4, 2);
  index.set(owner, [{ kind: 'cell', address: calculationTarget }]);
  index.setStructuralReference(owner, 'structural:barcode', [{ kind: 'cell', address: structuralTarget }]);
  index.setStructuralReference(owner, 'structural:formula-provenance', [{ kind: 'cell', address: structuralTarget }]);

  assert.deepEqual(index.getDependents(calculationTarget), [owner]);
  assert.deepEqual(index.getDependents(structuralTarget), []);
  assert.deepEqual(index.getStructuralDependents('Sheet1', 'row', 4), [owner]);
  assert.deepEqual(index.getStructuralReferenceOwnersInRange('Sheet1', {
    startRow: owner.row, endRow: owner.row, startColumn: owner.column, endColumn: owner.column,
  }), [
    { address: owner, sourceId: 'structural:barcode' },
    { address: owner, sourceId: 'structural:formula-provenance' },
  ]);

  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', sheetOrder: [{ id: 'Sheet1', name: 'Sheet1' }] });
  engine.setStructuralFormulaReference(owner, 'structural:barcode', '=A1+');
  assert.deepEqual(engine.dependencies.getInvalidFormulaOwners(), [owner]);
  engine.setStructuralFormulaReference(owner, 'structural:barcode', '=A1');
  assert.deepEqual(engine.dependencies.getInvalidFormulaOwners(), []);
  assert.deepEqual(engine.dependencies.getStructuralDependents('Sheet1', 'row', 0), [owner]);
});

test('formula-rule reference owners are indexed spatially and remove their recorded failures', () => {
  const index = new RangeIndex([{ id: 'Sheet1', name: 'Sheet1' }]);
  const context = address('Sheet1', 0, 0);
  const owner = { sheetId: 'Sheet1', ruleKind: 'data-validation' as const, ruleId: 'dv-1', field: 'formula1' };
  index.setFormulaRuleReference(owner, collectFormulaReferenceNodes(parseFormula('=A6')), context);

  assert.deepEqual(index.getStructuralFormulaRuleDependents('Sheet1', 'row', 5), [owner]);
  assert.deepEqual(index.getRangeFormulaRuleDependents('Sheet1', {
    startRow: 5, endRow: 5, startColumn: 0, endColumn: 0,
  }), [owner]);
  assert.deepEqual(index.getFormulaRuleReferenceFailures(), []);

  const invalidOwner = { ...owner, ruleId: 'dv-invalid' };
  index.setFormulaRuleReference(invalidOwner, [], context, 'invalid-formula');
  assert.deepEqual(index.getFormulaRuleReferenceFailures(), [{ owner: invalidOwner, reason: 'invalid-formula' }]);
  assert.equal(index.removeFormulaRuleReference(invalidOwner), true);
  assert.deepEqual(index.getFormulaRuleReferenceFailures(), []);
  assert.equal(index.removeFormulaRuleReference(owner), true);
  assert.deepEqual(index.getStructuralFormulaRuleDependents('Sheet1', 'row', 5), []);

  const opaqueIdOwner = { ...owner, ruleId: ' dv-padded ' };
  index.setFormulaRuleReference(opaqueIdOwner, collectFormulaReferenceNodes(parseFormula('=A6')), context);
  assert.deepEqual(index.getStructuralFormulaRuleDependents('Sheet1', 'row', 5), [opaqueIdOwner]);
});

test('reference index preserves exact canonical worksheet IDs', () => {
  const index = new RangeIndex([
    { id: 'sheet-1', name: 'INTEREST' },
    { id: 'sheet-2', name: 'Owner' },
  ]);
  const owner = address('sheet-2', 8, 5);
  index.set(owner, [{ kind: 'cell', address: address('sheet-1', 4, 2) }]);

  assert.deepEqual(index.getDependents(address('sheet-1', 4, 2)), [owner]);
});

test('reference index never remaps a canonical ID that collides with another display name', () => {
  const index = new RangeIndex([
    { id: 'owner-id', name: 'Owner' },
    { id: 'Target', name: 'Other' },
    { id: 'target-id', name: 'Target' },
  ]);
  const owner = address('owner-id', 8, 5);
  index.set(owner, [{ kind: 'cell', address: address('Target', 4, 2) }]);

  assert.deepEqual(index.getDependents(address('Target', 4, 2)), [owner]);
  assert.deepEqual(index.getDependents(address('target-id', 4, 2)), []);
});

test('range moves fail closed when whole-axis references would become non-contiguous', () => {
  const move: MoveRangeReferenceTransform = {
    selection: { sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    rowDelta: 2,
    columnDelta: 2,
    ownerSheetId: 'sheet-1',
    targetSheetId: 'sheet-1',
    targetSheetName: 'Sheet1',
    sheetOrder: [{ id: 'sheet-1', name: 'Sheet1' }],
  };
  assert.throws(() => mapAstMovedReferences(parseFormula('=SUM(A:A)'), move), /whole-column reference non-contiguous/);
  assert.throws(() => mapAstMovedReferences(parseFormula('=SUM(1:1)'), move), /whole-row reference non-contiguous/);
  assert.equal(formatFormula(mapAstMovedReferences(parseFormula('=SUM(A:A)'), { ...move, columnDelta: 0 })), '=SUM(A:A)');
});

test('3-D structural references require a resolvable target worksheet identity', () => {
  const formula = parseFormula('=SUM(Sheet1:Sheet2!A1)');
  const sheetOrder = [
    { id: 'sheet-1', name: 'Sheet1' },
    { id: 'sheet-2', name: 'Sheet2' },
    { id: 'sheet-3', name: 'Sheet3' },
  ];
  const shift = { axis: 'row' as const, at: 0, count: 1, op: 'insert' as const };

  assert.equal(formatFormula(mapAstStructuralReferences(formula, {
    shift,
    ownerSheetId: 'sheet-3',
    targetSheetId: 'sheet-3',
    sheetOrder,
  })), '=SUM(Sheet1:Sheet2!A1)');
  assert.throws(() => mapAstStructuralReferences(formula, {
    shift,
    ownerSheetId: 'sheet-3',
    targetSheetId: 'missing-sheet',
    sheetOrder,
  }), /target worksheet identity is unresolved/);
});

test('structural formula references resolve display names before colliding IDs', () => {
  const formula = parseFormula('=End!A1');
  const mapped = mapAstStructuralReferences(formula, {
    shift: { axis: 'row', at: 0, count: 1, op: 'insert' },
    ownerSheetId: 'owner-id',
    targetSheetId: 'End',
    targetSheetName: 'Target',
    sheetOrder: [
      { id: 'owner-id', name: 'Owner' },
      { id: 'End', name: 'Target' },
      { id: 'end-id', name: 'End' },
    ],
  });
  assert.equal(formatFormula(mapped), '=End!A1');
});

test('3-D reference boundaries resolve display names before colliding IDs', () => {
  const formula = parseFormula('=SUM(Start:End!A1)');
  assert.throws(() => mapAstStructuralReferences(formula, {
    shift: { axis: 'row', at: 0, count: 1, op: 'insert' },
    ownerSheetId: 'target-id',
    targetSheetId: 'target-id',
    targetSheetName: 'Target',
    sheetOrder: [
      { id: 'start-id', name: 'Start' },
      { id: 'End', name: 'Other' },
      { id: 'target-id', name: 'Target' },
      { id: 'end-id', name: 'End' },
    ],
  }), /inside a 3D reference/);
});

test('FormulaEngine input-address range index follows value, formula, clear and reset lifecycle', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  const valueAddress = address('Sheet1', 4, 2);
  const formulaAddress = address('Sheet1', 8, 5);
  engine.setValue(valueAddress, 12);
  engine.setFormula(formulaAddress, '=1+1');

  const range = { sheetId: 'Sheet1', startRow: 4, endRow: 8, startColumn: 2, endColumn: 5 };
  assert.deepEqual(engine.getInputAddressesInRange(range), [valueAddress, formulaAddress]);

  const middleAddress = address('Sheet1', 6, 2);
  engine.setValue(middleAddress, 24);
  assert.deepEqual(engine.getInputAddressesInRange(range), [valueAddress, middleAddress, formulaAddress]);
  engine.setFormula(valueAddress, '=2+2');
  assert.deepEqual(engine.getInputAddressesInRange(range), [valueAddress, middleAddress, formulaAddress]);
  engine.clearCell(valueAddress);
  assert.deepEqual(engine.getInputAddressesInRange(range), [middleAddress, formulaAddress]);
  engine.reset();
  assert.deepEqual(engine.getInputAddressesInRange(range), []);
});

test('visibility changes enqueue SUBTOTAL and AGGREGATE formulas without dirtying ordinary formulas', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  const subtotal = address('Sheet1', 0, 0);
  const aggregate = address('Sheet1', 1, 0);
  engine.setDefinedNameModels([{ name: 'VisibleSubtotal', formula: '=SUBTOTAL(9,B1:B3)', scope: 'workbook' }]);
  engine.setFormula(subtotal, '=VisibleSubtotal');
  engine.setFormula(aggregate, '=AGGREGATE(9,5,B1:B3)');
  engine.setFormula(address('Sheet1', 2, 0), '=SUM(B1:B3)');

  engine.notifyVisibilityChanged();

  assert.deepEqual(engine.getPendingRecalculationRoots(), [subtotal, aggregate]);
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

test('calculation task port is versioned and serializable without pretending to be a Worker', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  const port = engine.createCalculationTaskPort();
  const result = await port.submit({
    protocol: 'react-sheets.formula-calculation',
    version: 2,
    taskId: 'task-1',
    kind: 'recalculate',
    revision: 4,
    roots: [{ sheetId: 'Sheet1', row: 0, column: 0 }],
  });

  assert.equal(result.status, 'completed');
  assert.equal(result.protocol, 'react-sheets.formula-calculation');
  assert.equal(result.version, 2);
  assert.equal(result.revision, 4);
  assert.equal(result.report?.results.some((entry) => entry.value === 6), true);
});

test('FormulaEngine supports qualified references and detects cycles', () => {
  const engine = new FormulaEngine({
    defaultSheetId: 'Sheet1',
    sheetOrder: [
      { id: 'Sheet1', name: 'Sheet1' },
      { id: 'Sheet2', name: 'Sheet2' },
    ],
  });
  engine.setValue({ sheetId: 'Sheet2', row: 0, column: 0 }, 7);
  assert.equal(engine.setFormula('A1', '=Sheet2!A1 + 1').value, 8);

  const firstCycle = engine.setFormula('B1', '=C1 + 1').value;
  const secondCycle = engine.setFormula('C1', '=B1 + 1').value;
  assert.equal(firstCycle, 1);
  assertError(secondCycle, '#NUM!');
  assertError(engine.getCellValue('B1'), '#NUM!');
});

test('FormulaEngine evaluates qualified display names against canonical worksheet IDs', () => {
  const engine = new FormulaEngine({
    defaultSheetId: 'owner-id',
    sheetOrder: [
      { id: 'owner-id', name: 'Owner' },
      { id: 'Target', name: 'Other' },
      { id: 'target-id', name: 'Target' },
      { id: 'space-id', name: ' Target ' },
    ],
  });
  const owner = address('owner-id', 0, 1);
  const targetA1 = address('target-id', 0, 0);
  const targetA2 = address('target-id', 1, 0);
  const spacedNameA1 = address('space-id', 0, 0);
  engine.setValue(address('Target', 0, 0), 100);
  engine.setValue(targetA1, 7);
  engine.setValue(targetA2, 11);
  engine.setValue(spacedNameA1, 4);

  assert.equal(engine.setFormula(owner, '=tArGeT!A1+1').value, 8);
  assert.deepEqual(engine.getDependencies(owner), [{ kind: 'cell', address: targetA1 }]);
  assert.deepEqual(engine.getDependents(targetA1), [owner]);
  assert.equal(engine.setFormula(address('owner-id', 1, 1), '=SUM(Target!A1:A2)').value, 18);
  assert.equal(engine.setFormula(address('owner-id', 2, 1), "=' Target '!A1+1").value, 5);
  assertError(engine.setFormula(address('owner-id', 3, 1), '=Missing!A1').value, '#REF!');
});

test('dependency collection rejects qualified names without worksheet identities', () => {
  assert.throws(
    () => collectFormulaDependencies(parseFormula('=Remote!A1'), address('owner-id', 0, 0), { sheetOrder: [] }),
    /Worksheet identity order is required to resolve/,
  );
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
