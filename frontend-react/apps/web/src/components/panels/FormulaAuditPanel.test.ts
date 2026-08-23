import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createFormulaError, type FormulaDependency } from '@react-sheets/formula-engine';
import type { FormulaAuditProjection } from '@react-sheets/spreadsheet-app';
import {
  formatFormulaAuditAddress,
  formatFormulaAuditDependency,
  formatFormulaAuditValue,
  resolveFormulaAuditPanelState,
} from './FormulaAuditPanel';

const selectedCell = { sheetId: 'Sheet 1', row: 1, column: 1 };

function projection(overrides: Partial<FormulaAuditProjection> = {}): FormulaAuditProjection {
  return {
    selectedCell,
    arrows: [],
    showFormulas: false,
    formulas: [],
    errors: [],
    ...overrides,
  };
}

describe('FormulaAuditPanel view contract', () => {
  it('formats addresses and typed dependency ranges for visible rows', () => {
    const cell: FormulaDependency = { kind: 'cell', address: selectedCell };
    const range: FormulaDependency = {
      kind: 'range',
      start: { sheetId: 'Sheet 1', row: 0, column: 0 },
      end: { sheetId: 'Sheet 1', row: 2, column: 2 },
    };

    assert.equal(formatFormulaAuditAddress(selectedCell), "'Sheet 1'!B2");
    assert.equal(formatFormulaAuditDependency(cell), "'Sheet 1'!B2");
    assert.equal(formatFormulaAuditDependency(range), "'Sheet 1'!A1:'Sheet 1'!C3");
  });

  it('keeps formula errors and matrices readable without losing their type', () => {
    assert.equal(formatFormulaAuditValue(12), '12');
    assert.equal(formatFormulaAuditValue([[1, 'North'], [true, null]]), '1, North; true, ');
    assert.equal(formatFormulaAuditValue(createFormulaError('#DIV/0!', 'Division by zero')), '#DIV/0! Division by zero');
  });

  it('derives loading, empty, error, and ready panel states from explicit state and projection', () => {
    assert.equal(resolveFormulaAuditPanelState(undefined), 'empty');
    assert.equal(resolveFormulaAuditPanelState(null), 'empty');
    assert.equal(resolveFormulaAuditPanelState(projection()), 'ready');
    assert.equal(resolveFormulaAuditPanelState(undefined, 'loading'), 'loading');
    assert.equal(resolveFormulaAuditPanelState(undefined, 'error'), 'error');
    assert.equal(resolveFormulaAuditPanelState(projection(), 'empty'), 'empty');
  });
});

