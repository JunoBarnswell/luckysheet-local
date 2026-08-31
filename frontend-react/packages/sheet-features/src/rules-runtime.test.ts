import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { createResolvedCellReader, createRuleRuntime, resolveValidationRule } from './rules-runtime';

test('ResolvedCellReader keeps authored, calculated and visibility projections separate', () => {
  const workbook = new WorkbookModel('resolved', 'Resolved');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 2 });
  sheet.cells.set(0, 1, { value: null, formula: '=A1*3' });
  const formula = new FormulaEngine({ defaultSheetId: sheet.id });
  formula.setValue('A1', 2);
  formula.setFormula('B1', '=A1*3');
  const cells = createResolvedCellReader({ workbook, formula });
  const resolved = cells.resolve(sheet.id, 0, 1);
  assert.equal(resolved.authoredValue, null);
  assert.equal(resolved.authoredFormula, '=A1*3');
  assert.equal(resolved.calculatedValue, 6);
  assert.equal(resolved.displayValue, '6');
  assert.equal(resolved.visibility.manualHidden, false);
});

test('overlapping validation owners fail close while conditional rules remain indexed', () => {
  const workbook = new WorkbookModel('rules', 'Rules');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };
  sheet.dataValidations.push(
    { id: 'dv-a', sheetId: sheet.id, ranges: [range], type: 'whole', operator: 'greaterThan', formula1: '0' },
    { id: 'dv-b', sheetId: sheet.id, ranges: [range], type: 'whole', operator: 'lessThan', formula1: '10' },
  );
  assert.throws(() => resolveValidationRule(sheet, 0, 0), /RULE_OWNER_AMBIGUOUS/);
  sheet.dataValidations.splice(1, 1);
  const runtime = createRuleRuntime(sheet);
  assert.equal(runtime.resolve(0, 0).validation?.id, 'dv-a');
  assert.equal(runtime.resolve(2, 0).validation, undefined);
});
