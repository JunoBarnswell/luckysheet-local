import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRegistry, CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import {
  FormulaAuditController,
  getFormulaDependents,
  getFormulaPrecedents,
  registerFormulaAuditCommands,
  scanFormulaErrors,
} from './index';

function createFixture(): { workbook: WorkbookModel; formula: FormulaEngine; sheetId: string } {
  const workbook = new WorkbookModel('wb-formula-audit', 'Formula audit');
  const sheetId = workbook.primarySheetId;
  const formula = new FormulaEngine({ defaultSheetId: sheetId });
  formula.setValue('A1', 2);
  formula.setValue('A2', 3);
  formula.setFormula('B1', '=SUM(A1:A2)*2');
  formula.setFormula('C1', '=B1+1');
  formula.setFormula('D1', '=1/0');
  return { workbook, formula, sheetId };
}

describe('formula audit projections', () => {
  it('projects precedents and dependents from the engine dependency graph', () => {
    const { formula, sheetId } = createFixture();
    const precedents = getFormulaPrecedents(formula, { sheetId, row: 0, column: 1 });
    assert.equal(precedents.length, 1);
    assert.equal(precedents[0]?.target.kind, 'range');

    const dependents = getFormulaDependents(formula, { sheetId, row: 0, column: 1 });
    assert.deepEqual(dependents.map((arrow) => arrow.formulaCell.column), [2]);
    assert.equal(dependents[0]?.direction, 'dependent');
  });

  it('scans calculation errors, shows formulas and evaluates every AST step', () => {
    const { formula, sheetId } = createFixture();
    const controller = new FormulaAuditController(formula);
    const errors = scanFormulaErrors(formula, { sheetId });
    assert.deepEqual(errors.map((error) => error.code), ['#DIV/0!']);

    const shown = controller.setShowFormulas(true);
    assert.equal(shown.formulas.length, 3);
    const evaluation = controller.evaluateStep({ sheetId, row: 0, column: 1 });
    assert.ok(evaluation);
    assert.equal(evaluation?.value, 10);
    assert.ok((evaluation?.steps.length ?? 0) >= 4);
    assert.equal(controller.removeArrows().arrows.length, 0);
  });

  it('exposes audit commands through CommandRuntime without fake mutations', () => {
    const { workbook, formula, sheetId } = createFixture();
    const controller = new FormulaAuditController(formula);
    const runtime = new CommandRuntime(workbook, new CommandRegistry({ requireMutationMetadata: true }));
    const commandIds = registerFormulaAuditCommands(runtime.registry, controller);
    assert.equal(commandIds.length, 7);

    runtime.execute('formula.audit.precedents.show', { address: { sheetId, row: 0, column: 1 } });
    assert.equal(controller.getProjection().arrows.length, 1);
    assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
    runtime.execute('formula.audit.arrows.remove', {});
    assert.equal(controller.getProjection().arrows.length, 0);
    runtime.execute('formula.audit.errors.scan', { sheetId });
    assert.equal(controller.getProjection().errors.length, 1);
    runtime.execute('formula.calculation.mode.set', { mode: 'manual' });
    assert.equal(formula.getRecalculationMode(), 'manual');
  });
});
