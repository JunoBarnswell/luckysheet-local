import assert from 'node:assert/strict';
import test from 'node:test';
import { registerSheetCommands } from './index';
import { createResolvedCellReader, resolveValidationRule } from './rules-runtime';
import { openCanonicalTestRuntime } from '../../core-model/src/canonical-test-runtime.test';

test('ResolvedCellReader keeps canonical authored values separate from calculated values and visibility', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('resolved-canonical');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet('sheet-1');
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 0, column: 0, value: { value: 2 } });
    const resolved = createResolvedCellReader({ workbook }).resolve(sheet.id, 0, 0);
    assert.equal(resolved.authoredValue, 2);
    assert.equal(resolved.calculatedValue, 2);
    assert.equal(resolved.displayValue, '2');
    assert.equal(resolved.visibility.manualHidden, false);
  } finally {
    close();
  }
});

test('overlapping validation owners fail closed without mutating canonical rule indexes', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('rules-rejection');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet('sheet-1');
    const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };
    sheet.dataValidations.push(
      { id: 'dv-a', sheetId: sheet.id, ranges: [range], type: 'whole', operator: 'greaterThan', formula1: '0' },
      { id: 'dv-b', sheetId: sheet.id, ranges: [range], type: 'whole', operator: 'lessThan', formula1: '10' },
    );
    assert.throws(() => resolveValidationRule(sheet, 0, 0), /RULE_OWNER_AMBIGUOUS/);
    assert.equal(sheet.dataValidations.length, 2);
  } finally {
    close();
  }
});
