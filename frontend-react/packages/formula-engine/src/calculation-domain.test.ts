import assert from 'node:assert/strict';
import test from 'node:test';
import { FormulaEngine, formatFormula, parseFormula, renameAstSheetReferences, mapAstStructuralReferences, mapAstMovedReferences, collectNameReferences, isFormulaError, STANDARD_FORMULA_ERRORS } from './index';

test('permission-projected and authored error constants keep their identity through Worker calculation', () => {
  const engine = new FormulaEngine();
  for (const code of STANDARD_FORMULA_ERRORS) {
    const ast = parseFormula(`=${code}`);
    assert.equal(formatFormula(ast), `=${code}`);
    assert.equal((engine.setFormula('A1', `=${code}`).value as { code: string }).code, code);
  }
  engine.setFormula('A1', '=#BLOCKED!');
  assert.equal((engine.setFormula('B1', '=SUM(A1)').value as { code: string }).code, '#BLOCKED!');
  assert.equal(engine.setFormula('B2', '=IFERROR(A1,0)').value, 0);
  const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
  worker.recalculate();
  assert.equal((worker.getCellValue('A1') as { code: string }).code, '#BLOCKED!');
  assert.equal((worker.getCellValue('B1') as { code: string }).code, '#BLOCKED!');
  assert.equal((engine.setFormula('A2', '=#UNKNOWN!').value as { code: string }).code, '#PARSE!');
});

test('qualified ranges retain one worksheet owner across structural changes', () => {
  const identity = { ownerSheetId: 'summary', targetSheetId: 'source', targetSheetName: 'Source', sheetOrder: [{ id: 'source', name: 'Source' }, { id: 'summary', name: 'Summary' }] };
  const ast = parseFormula('=SUM(Source!A1:B2)');
  assert.equal(formatFormula(mapAstStructuralReferences(ast, { ...identity, shift: { axis: 'row', at: 0, count: 1, op: 'insert' } })), '=SUM(Source!A2:B3)');
  const moved = { ...identity, selection: { sheetId: 'source', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, rowDelta: 2, columnDelta: 1 };
  assert.equal(formatFormula(mapAstMovedReferences(ast, moved)), '=SUM(Source!B3:C4)');
  assert.throws(() => mapAstMovedReferences(ast, { ...moved, selection: { ...moved.selection, endRow: 0 } }), /non-contiguous/);
  assert.throws(() => mapAstStructuralReferences(parseFormula('=SUM(Source!A1:Summary!B2)'), { ...identity, shift: { axis: 'row', at: 0, count: 1, op: 'insert' } }), /different worksheets/);
});

test('sparse 3D aggregates consume occupied values without reading rectangular matrices', () => {
  const sheets = [{ id: 'jan', name: 'Jan' }, { id: 'dec', name: 'Dec' }, { id: 'summary', name: 'Summary' }];
  const engine = new FormulaEngine({ defaultSheetId: 'summary', sheetOrder: sheets });
  for (const sheet of sheets) engine.setSpillEnvironment(sheet.id, { rowCount: 1_000_000, columnCount: 20, isOccupied: () => false });
  engine.setValue({ sheetId: 'jan', row: 999_999, column: 0 }, 10);
  engine.setValue({ sheetId: 'dec', row: 0, column: 0 }, 30);
  engine['readRangeMatrix'] = () => { throw new Error('Sparse aggregate materialized blanks'); };
  for (const [name, expected] of [['SUM', 40], ['AVERAGE', 20], ['COUNT', 2], ['COUNTA', 2], ['MIN', 10], ['MAX', 30], ['PRODUCT', 300], ['VAR.P', 100], ['VAR.S', 200], ['STDEV.P', 10]] as const) {
    assert.equal(engine.setFormula('C1', `=${name}(Jan:Dec!A1:A1000000)`).value, expected, name);
  }
  engine.setValue({ sheetId: 'jan', row: 999_999, column: 0 }, 20);
  assert.equal(engine.getCellValue('C1'), 5);
  const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
  worker.recalculate();
  assert.equal(worker.getCellValue('C1'), 5);
});

test('aggregate reference coercion, error counting and empty variance remain distinct', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', true);
  engine.setValue('A2', '10');
  engine.setFormula('A3', '=""');
  engine.setFormula('A4', '=1/0');
  assert.equal(engine.setFormula('B1', '=COUNTA(A1:A100000)').value, 4);
  assert.equal(engine.setFormula('B2', '=COUNT(A1:A100000)').value, 0);
  assert.equal(engine.setFormula('B3', '=SUM(TRUE,"10")').value, 11);
  assert.equal(engine.setFormula('B4', '=SUM(A1:A2)').value, 0);
  assert.equal(engine.setFormula('B7', '=SUM(A1,A2)').value, 0);
  assert.equal(engine.setFormula('B8', '=COUNT(A1,A2)').value, 0);
  assert.equal((engine.setFormula('B5', '=SUM(A1:A4)').value as { code: string }).code, '#DIV/0!');
  assert.equal((engine.setFormula('B6', '=VAR.S(A1:A2)').value as { code: string }).code, '#DIV/0!');
});

test('sparse aggregate includes spill children whose anchor lies outside the range', () => {
  const engine = new FormulaEngine();
  engine.setSpillEnvironment('Sheet1', { rowCount: 100, columnCount: 10, isOccupied: () => false });
  engine.setFormula('A1', '=SEQUENCE(3,1,1,1)');
  assert.equal(engine.setFormula('C1', '=SUM(A2:A100)').value, 5);
  assert.equal(engine.setFormula('C2', '=SUM(A1:A100)').value, 6);
  engine.setFormula('A1', '=SEQUENCE(2,1,10,1)');
  assert.equal(engine.getCellValue('C1'), 11);
  assert.equal(engine.getCellValue('C2'), 21);
});

test('defined names preserve reference provenance and use the same lexical evaluator', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', true);
  engine.setValue('A2', '10');
  engine.setValue('A3', 20);
  engine.setDefinedNameModels([
    { name: 'Flag', scope: 'workbook', formula: '=$A$1' },
    { name: 'Inputs', scope: 'workbook', formula: '=$A$1:$A$3' },
    { name: 'Twice', scope: 'workbook', formula: '=LET(factor,2,LAMBDA(x,x*factor))' },
    { name: 'Again', scope: 'workbook', formula: '=Twice' },
  ]);
  assert.equal(engine.setFormula('B1', '=SUM(Flag)').value, 0);
  assert.equal(engine.setFormula('B2', '=SUM(Inputs)').value, 20);
  assert.equal(engine.setFormula('B3', '=Again(5)').value, 10);
  engine.setValue('A3', 30);
  assert.equal(engine.getCellValue('B2'), 30);
  const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
  worker.recalculate();
  assert.equal(worker.getCellValue('B1'), 0);
  assert.equal(worker.getCellValue('B2'), 30);
  assert.equal(worker.getCellValue('B3'), 10);
  const result = engine.setFormula('B4', '=Flag(1)').value;
  assert.equal(isFormulaError(result) ? result.code : result, '#VALUE!');
});

test('LET uses lexical scope and retains cross-sheet dependencies in a Worker snapshot', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'summary', sheetOrder: [{ id: 'source', name: 'Source' }, { id: 'summary', name: 'Summary' }] });
  engine.setValue({ sheetId: 'source', row: 0, column: 0 }, 10);
  assert.equal(engine.setFormula('A1', '=LET(total,Source!A1,total*2)').value, 20);
  engine.setValue({ sheetId: 'source', row: 0, column: 0 }, 12);
  assert.equal(engine.getCellValue('A1'), 24);
  assert.equal(engine.setFormula('A2', '=LET(total,2,LET(total,3,total)+total)').value, 5);
  assert.deepEqual(collectNameReferences(parseFormula('=LET(total,Rate,total*2)')), ['RATE']);
  const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
  worker.recalculate();
  assert.equal(worker.getCellValue('A1'), 24);
});

test('LAMBDA supports expression calls, lexical capture, named calls and incremental updates', () => {
  const engine = new FormulaEngine();
  assert.equal(engine.setFormula('A1', '=LAMBDA(price,price*1.13)(100)').value, 113);
  assert.equal(engine.setFormula('A2', '=LET(amount,5,fn,LAMBDA(x,x+amount),LET(amount,9,fn(2)))').value, 7);
  engine.setValue('B1', 2);
  engine.setDefinedNameModels([{ name: 'SCALED', scope: 'workbook', formula: '=LAMBDA(x,x*$B$1)' }]);
  assert.equal(engine.setFormula('A3', '=SCALED(4)').value, 8);
  engine.setValue('B1', 3);
  assert.equal(engine.getCellValue('A3'), 12);
  engine.setDefinedNameModels([{ name: 'SCALED', scope: 'workbook', formula: '=LAMBDA(x,x*$B$1+1)' }]);
  assert.equal(engine.getCellValue('A3'), 13);
  const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
  worker.recalculate();
  assert.equal(worker.getCellValue('A3'), 13);
  assert.equal(formatFormula(parseFormula('=LAMBDA(x,x+1)(2)')), '=LAMBDA(x,x+1)(2)');
  assert.equal(formatFormula(parseFormula('=(1+2)(3)')), '=(1+2)(3)');
});

test('invalid lexical and callable operations return observable typed errors', () => {
  const engine = new FormulaEngine();
  for (const [formula, code] of [
    ['=LET(1,2,3)', '#VALUE!'], ['=LAMBDA(x,x)(1,2)', '#VALUE!'], ['=LAMBDA(x,x)', '#CALC!'], ['=LAMBDA(x,x,x)(1,2)', '#VALUE!'], ['=2(3)', '#VALUE!'],
  ] as const) {
    const result = engine.setFormula('A1', formula!).value;
    assert.equal(isFormulaError(result) ? result.code : result, code, formula);
  }
  engine.setDefinedNameModels([{ name: 'RECURSE', scope: 'workbook', formula: '=LAMBDA(x,RECURSE(x))' }]);
  const result = engine.setFormula('A1', '=RECURSE(1)').value;
  assert.equal(isFormulaError(result) ? result.code : result, '#NUM!');
});

test('quoted 3D and external qualifiers round trip through the canonical AST', () => {
  for (const formula of ["=SUM('Jan Data:Dec Data'!B2)", "=SUM('[Source.xlsx]Sales Data'!B2:B4)", '=[source-unit-id]Sales!B2']) {
    const ast = parseFormula(formula);
    assert.deepEqual(parseFormula(formatFormula(ast)), ast);
  }
  const ast = parseFormula("=SUM('Jan Data:Dec Data'!B2)+SUM('[Source.xlsx]Jan Data'!B2)");
  assert.equal(formatFormula(renameAstSheetReferences(ast, 'Jan Data', 'January')), "=SUM('January:Dec Data'!B2)+SUM('[Source.xlsx]Jan Data'!B2)");
});
