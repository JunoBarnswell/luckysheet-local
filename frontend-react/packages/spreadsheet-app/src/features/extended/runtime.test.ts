import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { runDataTable, runGoalSeek, runScenario, summarizeGoalSeekResult } from './runtime';

describe('extended runtime', () => {
  it('runs goal seek against formula dependencies', () => {
    const workbook = new WorkbookModel('wb-goal', 'Goal Seek');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const sheetId = workbook.primarySheetId;
    runtime.execute('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { formula: '=B1*2' },
    });
    runtime.execute('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 1,
      value: { value: 10 },
    });

    const formula = new FormulaEngine({ defaultSheetId: sheetId });
    formula.setFormula({ sheetId, row: 0, column: 0 }, '=B1*2');
    formula.setValue({ sheetId, row: 0, column: 1 }, 10);
    formula.recalculate();

    const result = runGoalSeek(workbook, formula, sheetId, {
      setCell: { row: 0, column: 0 },
      toValue: 100,
      byChangingCell: { row: 0, column: 1 },
    });
    assert.equal(result.status, 'converged');
    assert.equal(result.changingCellValue, 50);
    assert.match(summarizeGoalSeekResult(result), /converged/i);
  });

  it('runs one-variable column input data table', () => {
    const workbook = new WorkbookModel('wb-table', 'Data Table');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const sheetId = workbook.primarySheetId;
    runtime.execute('sheet.cell.set', { sheetId, row: 1, column: 0, value: { formula: '=B1*2' } });
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 5 } });
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 2, value: { value: 10 } });
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 3, value: { value: 15 } });

    const formula = new FormulaEngine({ defaultSheetId: sheetId });
    formula.setFormula({ sheetId, row: 1, column: 0 }, '=B1*2');
    formula.setValue({ sheetId, row: 0, column: 1 }, 5);
    formula.recalculate();

    const result = runDataTable(workbook, formula, sheetId, {
      columnInputCell: { row: 0, column: 1 },
      tableRange: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 3 },
    });
    assert.equal(result.status, 'completed');
    assert.equal(result.filledCells, 3);
    assert.deepEqual(result.writes, [
      { row: 1, column: 1, value: 10 },
      { row: 1, column: 2, value: 20 },
      { row: 1, column: 3, value: 30 },
    ]);
    assert.equal(workbook.getSheet(sheetId).cells.get(1, 1), undefined);
  });

  it('runs scenario analysis with result cells', () => {
    const workbook = new WorkbookModel('wb-scenario', 'Scenario');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const sheetId = workbook.primarySheetId;
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 0, value: { formula: '=B1*2' } });
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 10 } });

    const formula = new FormulaEngine({ defaultSheetId: sheetId });
    formula.setFormula({ sheetId, row: 0, column: 0 }, '=B1*2');
    formula.setValue({ sheetId, row: 0, column: 1 }, 10);
    formula.recalculate();

    const result = runScenario(workbook, formula, sheetId, {
      id: 'best-case',
      name: 'Best Case',
      changingCells: [{ row: 0, column: 1, value: 25 }],
      resultCells: [{ row: 0, column: 0 }],
    });
    assert.equal(result.status, 'completed');
    assert.equal(result.outputs[0]?.value, 50);
    assert.equal(workbook.getSheet(sheetId).cells.get(0, 1)?.value, 10);
  });
});
