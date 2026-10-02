import assert from 'node:assert/strict';
import test from 'node:test';
import { FormulaEngine, isFormulaError, type FormulaValue } from './index';
import { findLookupIndex } from './functions/lookup-engine';
const code = (value: FormulaValue) => isFormulaError(value) ? value.code : value;

test('million-row consumers keep sparse references and include blanks, errors, names and spill children', () => {
  const engine = new FormulaEngine();
  engine.setSpillEnvironment('Sheet1', { rowCount: 1_048_576, columnCount: 20, isOccupied: () => false });
  engine.setValue('A1', 'x'); engine.setValue('B1', 2);
  engine.setValue('A1000000', 'x'); engine.setValue('B1000000', 6);
  engine.setDefinedNameModels([{ name: 'Labels', scope: 'workbook', formula: '=$A$1:$A$1000000' }, { name: 'Amounts', scope: 'workbook', formula: '=$B$1:$B$1000000' }]);
  engine['readRangeMatrix'] = () => { throw new Error('Consumer allocated a rectangular matrix'); };
  for (const [formula, expected] of [
    ['=SUMIFS(Amounts,Labels,"x")', 8], ['=SUMIF(Labels,"x",Amounts)', 8], ['=AVERAGEIFS(Amounts,Labels,"x")', 4],
    ['=COUNTIFS(Labels,"x")', 2], ['=COUNTIF(Labels,"")', 999998], ['=COUNTIF(Labels,"<>")', 2],
    ['=COUNTBLANK(B1:B1000000)', 999998], ['=MINIFS(Amounts,Labels,"x")', 2], ['=MAXIFS(Amounts,Labels,"x")', 6],
    ['=SUBTOTAL(9,Amounts)', 8], ['=AGGREGATE(9,5,Amounts)', 8], ['=SUMPRODUCT(Amounts,Amounts)', 40],
    ['=INDEX(Amounts,1000000)', 6], ['=ROWS(Amounts)', 1000000], ['=MEDIAN(Amounts)', 4],
    ['=SUM(INDEX(B1:B1000000,0,1))', 8], ['=COUNTBLANK(B:B)', 1048574], ['=ROWS(A:A)', 1048576], ['=COLUMNS(1:1)', 16384],
  ] as const) assert.equal(engine.setFormula('D1', formula).value, expected, formula);
  engine.setFormula('B1000000', '=#BLOCKED!');
  assert.equal(code(engine.setFormula('D1', '=SUMIFS(Amounts,Labels,"x")').value), '#BLOCKED!');
  assert.equal(engine.setFormula('D1', '=AGGREGATE(9,7,Amounts)').value, 2);
  assert.equal(code(engine.setFormula('D1', '=SUMPRODUCT(Amounts,Amounts)').value), '#BLOCKED!');
  engine.setFormula('A1', '=#BLOCKED!');
  assert.equal(code(engine.setFormula('D1', '=COUNTIFS(Labels,"x")').value), '#BLOCKED!');
  assert.equal(code(engine.setFormula('D1', '=SUMIFS(Amounts,Labels,"x")').value), '#BLOCKED!');
});

test('condition ranges project from the target top-left and invalidate cells beyond authored geometry', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 'x'); engine.setValue('A2', 'x'); engine.setValue('A3', 'x');
  engine.setValue('B1', 2); engine.setValue('B2', 4); engine.setValue('B3', 6);
  engine.setDefinedNameModels([{ name: 'Labels', scope: 'workbook', formula: '=$A$1:$A$3' }, { name: 'Origin', scope: 'workbook', formula: '=$B$1' }]);
  assert.equal(engine.setFormula('D1', '=SUMIF(Labels,"x",Origin)').value, 12);
  assert.equal(engine.setFormula('D2', '=AVERAGEIF(A1:A3,"x",B1)').value, 4);
  engine.setValue('B3', 12);
  assert.equal(engine.getCellValue('D1'), 18); assert.equal(engine.getCellValue('D2'), 6);
  assert.equal(code(engine.setFormula('D3', '=SUMIFS(B1:B3,A1:A2,"x")').value), '#VALUE!');
  engine.setFormula('B3', '=1/0');
  assert.equal(code(engine.getCellValue('D1')), '#DIV/0!');
  assert.equal(code(engine.getCellValue('D2')), '#DIV/0!');
});

test('conditional extrema include a matching empty target and preserve error priority', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 'x'); engine.setValue('A2', 'x'); engine.setValue('B1', 4);
  assert.equal(engine.setFormula('D1', '=MINIFS(B1:B3,A1:A3,"x")').value, 0);
  engine.setValue('B1', -4);
  assert.equal(engine.setFormula('D2', '=MAXIFS(B1:B3,A1:A3,"x")').value, 0);
  assert.equal(engine.setFormula('D3', '=COUNTIF(B1:B3,"")').value, 2);
  engine.setValue('B2', 0);
  assert.equal(engine.getCellValue('D3'), 1);
  engine.setValue('A1', 'a*');
  assert.equal(engine.setFormula('D4', '=COUNTIF(A1:A3,"a~*")').value, 1);
});

test('binary lookup consumes a million-element vector in logarithmic reads in both directions', () => {
  for (const search of [2, -2]) {
    let reads = 0;
    const vector = { length: 1_000_000, at(index: number) { reads++; return search === 2 ? index * 2 : (999999 - index) * 2; } };
    assert.equal(findLookupIndex(1_000_000, vector, 0, search), search === 2 ? 500000 : 499999);
    assert.ok(reads <= 23, String(reads));
    assert.equal(findLookupIndex(1_000_001, vector, -1, search), search === 2 ? 500000 : 499999);
    assert.equal(findLookupIndex(1_000_001, vector, 1, search), search === 2 ? 500001 : 499998);
  }
  assert.equal(findLookupIndex(2, [1,2,2,3], 0, -1), 2);
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setValue('A2', 4); engine.setValue('A3', 6);
  engine.setValue('B1', 'x'); engine.setValue('B2', 'y'); engine.setValue('B3', 'z');
  assert.equal(engine.setFormula('D1', '=XLOOKUP(4,A1:A3,B1:B3,"missing",0,2)').value, 'y');
  assert.equal(code(engine.setFormula('D2', '=XLOOKUP(4,A1:B3,B1:B3)').value), '#VALUE!');
  assert.equal(code(engine.setFormula('D3', '=XLOOKUP(4,A1:A3,B1:B2)').value), '#VALUE!');
});

test('deep reverse-address chains calculate iteratively and preserve SCC rejection', () => {
  const engine = new FormulaEngine();
  const length = 6000;
  engine.synchronizeInputs(Array.from({ length }, (_, row) => ({ address: { sheetId: 'Sheet1', row, column: 0 }, input: row === length - 1 ? { kind: 'value' as const, value: 1 } : { kind: 'formula' as const, formula: `=A${row + 2}+1` } })));
  engine.recalculate();
  assert.equal(engine.getCellValue('A1'), length);
  engine.setValue(`A${length}`, 2); assert.equal(engine.getCellValue('A1'), length + 1);
  engine.setFormula(`A${length}`, '=A1'); assert.equal(code(engine.getCellValue('A1')), '#NUM!');
  engine.setValue(`A${length}`, 3); assert.equal(engine.getCellValue('A1'), length + 2);
});

test('body-only formula edits reuse topology and clean prerequisites while changed edges invalidate it', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setValue('D1', 3);
  engine.setFormula('B1', '=SUM(A1:A100000)'); engine.setFormula('C1', '=B1+D1');
  const generation = engine.getFormulaTopologyRevision();
  const original = engine['createEvaluationContext'].bind(engine);
  let sourceEvaluations = 0;
  engine['createEvaluationContext'] = (...args) => { if (args[0].address.column === 1) sourceEvaluations++; return original(...args); };
  engine.setValue('D1', 4); assert.equal(engine.getCellValue('C1'), 6); assert.equal(sourceEvaluations, 0);
  engine.setFormula('C1', '=B1+D1+1'); assert.equal(engine.getFormulaTopologyRevision(), generation); assert.equal(engine.getCellValue('C1'), 7);
  engine.setFormula('C1', '=B1+A1'); assert.ok(engine.getFormulaTopologyRevision() > generation); assert.equal(engine.getCellValue('C1'), 4);
});

test('lazy branches, scalar broadcasting and error fallbacks retain array shape', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setValue('A2', -4);
  engine.setValue('B1', true); engine.setValue('B2', false);
  engine['readRangeMatrix'] = ((original) => (range, ...args) => { if (range.end.row > 1000) throw new Error('Unselected branch was consumed'); return original(range, ...args); })(engine['readRangeMatrix'].bind(engine));
  assert.equal(engine.setFormula('D1', '=IFERROR(2,SUMPRODUCT(A1:A1000000))').value, 2);
  assert.equal(engine.setFormula('D2', '=CHOOSE(1,2,SUMPRODUCT(A1:A1000000))').value, 2);
  assert.deepEqual(engine.setFormula('D3', '=ABS(A1:A2)').value, [[2],[4]]);
  assert.deepEqual(engine.setFormula('D4', '=IF(B1:B2,A1:A2,7)').value, [[2],[7]]);
  assert.deepEqual(engine.setFormula('D5', '=IFERROR(A1:A2/0,7)').value, [[7],[7]]);
  assert.equal(code(engine.setFormula('D6', '=AND(FALSE,#BLOCKED!)').value), '#BLOCKED!');
  assert.equal(code(engine.setFormula('D7', '=OR(TRUE,#BLOCKED!)').value), '#BLOCKED!');
  assert.equal(code(engine.setFormula('D8', '=LEN(#N/A)').value), '#N/A');
  assert.equal(engine.setFormula('D9', '=ISBLANK("")').value, false);
});

test('dynamic arrays filter columns, reject bad masks and use Excel padding and slicing', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 1); engine.setValue('B1', 2); engine.setValue('A2', 3); engine.setValue('B2', 4);
  engine.setValue('A3', true); engine.setValue('B3', false);
  assert.deepEqual(engine.setFormula('D1', '=FILTER(A1:B2,A3:B3)').value, [[1],[3]]);
  engine.setFormula('A3', '=#N/A'); assert.equal(code(engine.setFormula('D2', '=FILTER(A1:B2,A3:B3)').value), '#N/A');
  assert.deepEqual(engine.setFormula('D3', '=TAKE(A1:B2,2,-1)').value, [[2],[4]]);
  assert.equal(code(engine.setFormula('D4', '=TAKE(A1:B2,0)').value), '#CALC!');
  assert.equal(code(engine.setFormula('D5', '=DROP(A1:B2,2)').value), '#CALC!');
  const stacked = engine.setFormula('D6', '=HSTACK(A1:A2,B1)').value;
  assert.ok(Array.isArray(stacked)); assert.equal(code(stacked[1]![1]!), '#N/A');
  assert.equal(code(engine.setFormula('D7', '=SORT(A1:B2,3)').value), '#VALUE!');
  assert.equal(code(engine.setFormula('D8', '=SORTBY(A1:B2,A1)').value), '#VALUE!');
  assert.equal(code(engine.setFormula('D9', '=RANDARRAY(2,2,5,1)').value), '#VALUE!');
});

test('reference consumers support names, unions, intersections, R1C1 and correct OFFSET dimensions', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setValue('A2', 4); engine.setValue('A3', 6);
  engine.setDefinedNameModels([{ name: 'Inputs', scope: 'workbook', formula: '=$A$1:$A$3' }]);
  assert.equal(engine.setFormula('D1', '=SUM((A1,A2))').value, 6);
  assert.equal(engine.setFormula('D2', '=SUM(A1:A3 A2:A3)').value, 10);
  assert.equal(engine.setFormula('D3', '=SUM(OFFSET(Inputs,0,0))').value, 12);
  assert.equal(engine.setFormula('D4', '=INDIRECT("R2C1",FALSE)').value, 4);
  assert.equal(engine.setFormula('D5', '=INDIRECT("R[-3]C[-3]",FALSE)').value, 4);
  assert.equal(code(engine.setFormula('D6', '=INDIRECT("1+2")').value), '#REF!');
  assert.equal(code(engine.setFormula('D7', '=OFFSET(A1,0,0,0)').value), '#REF!');
  assert.equal(engine.setFormula('D8', '=ADDRESS(2,3,3)').value, '$C2');
  assert.equal(engine.setFormula('D9', '=ADDRESS(2,3,4,FALSE)').value, 'R[2]C[3]');
});

test('volatile references invalidate downstream formulas and cleared manual inputs remain dirty roots', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setFormula('B1', '=INDIRECT("A1")'); engine.setFormula('C1', '=B1*3');
  engine.setValue('A1', 4); assert.equal(engine.getCellValue('C1'), 12);
  engine.setRecalculationMode('manual'); engine.clearCell('A1'); engine.recalculate();
  assert.equal(engine.getCellValue('C1'), 0);
});

test('data-table scenarios do not overwrite canonical formula results or nested overrides', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setValue('A2', 4); engine.setValue('D1', 10); engine.setFormula('B1', '=A1*D1');
  assert.deepEqual(engine.setFormula('Z1', '=SJS.TABLE(B1,A1:A2,D1)').value, [[4],[8]]);
  assert.equal(engine.getCellValue('B1'), 20); assert.equal(engine.getCellValue('D1'), 10);
});


test('shape queries retain computed control dependencies and can recover from a value cycle', () => {
  const engine = new FormulaEngine();
  assert.equal(code(engine.setFormula('A1', '=SUM(A1)').value), '#NUM!');
  assert.equal(engine.setFormula('A1', '=ROWS(A1)').value, 1);
  assert.equal(code(engine.setFormula('A1', '=ABS(A1)').value), '#NUM!');
  assert.equal(engine.setFormula('A1', '=COLUMN(A1)').value, 1);
  engine.setValue('A2', 2);
  engine.setFormula('A1', '=ROWS(SEQUENCE(A2))');
  assert.equal(engine.getCellValue('A1'), 2);
  engine.setValue('A2', 3); assert.equal(engine.getCellValue('A1'), 3);
  engine.setFormula('B1', '=ROWS(INDEX(A:A,A2,0))');
  assert.equal(engine.getCellValue('B1'), 1);
  engine.setFormula('A2', '=B1'); assert.equal(code(engine.getCellValue('B1')), '#NUM!');
});
