import assert from 'node:assert/strict';
import { before, test } from 'node:test';
import { initializeNodeKernel } from '../../kernel-client/src/node';
import { kernelInvoke, KernelInvocationError } from '../../kernel-client/src/index';
import { FormulaEngine, formatFormula, parseFormula } from './index';

before(async () => { await initializeNodeKernel(); });

function createBinding(unitId: string): FormulaEngine {
  kernelInvoke('create', { unitId, name: 'Formula binding test', sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 100, columnCount: 26, metadata: {} }] });
  return new FormulaEngine({ unitId, revision: () => 0, defaultSheetId: 'sheet-1' });
}

test('Rust evaluates arithmetic and range formulas against the canonical workbook', () => {
  const engine = createBinding('formula-binding-success');
  assert.equal(engine.evaluateFormula('=SUM(1,2,3)*2', 'A1'), 12);
  assert.equal(engine.evaluateAst(parseFormula('=1+2*3'), { sheetId: 'sheet-1', row: 0, column: 0 }), 7);
  assert.equal(engine.evaluateFormula('=SUM(A2:A10)', 'A1'), 0);
  kernelInvoke('close', { unitId: engine.unitId });
});

test('revision mismatch is observable and never falls back to TypeScript evaluation', () => {
  const live = createBinding('formula-binding-rejection');
  const stale = new FormulaEngine({ unitId: live.unitId, revision: () => 1, defaultSheetId: 'sheet-1' });
  assert.throws(() => stale.evaluateFormula('=1+2', 'A1'), KernelInvocationError);
  kernelInvoke('close', { unitId: live.unitId });
});

test('the binding rejects missing identity and invalid revision', () => {
  assert.throws(() => new FormulaEngine({ unitId: '', defaultSheetId: 'sheet-1', revision: () => 0 }), { code: 'FORMULA_BINDING_INVALID' });
  const invalid = new FormulaEngine({ unitId: 'unknown', defaultSheetId: 'sheet-1', revision: () => -1 });
  assert.throws(() => invalid.evaluateFormula('=1', 'A1'), { code: 'REVISION_INVALID' });
});

test('editor AST round trips without evaluating a formula in the host', () => {
  assert.equal(formatFormula(parseFormula('=SUM($A$1:B2)')), '=SUM($A$1:B2)');
});
