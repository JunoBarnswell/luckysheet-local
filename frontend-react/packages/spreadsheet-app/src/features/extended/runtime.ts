import type { WorkbookModel } from '@react-sheets/core-model';
import type { FormulaEngine } from '@react-sheets/formula-engine';
import {
  CapabilityRegistry,
  type CapabilityDescriptor,
  type PlatformCapability,
} from './index';
import type {
  GoalSeekParams,
  GoalSeekResult,
  ScenarioDefinition,
  ScenarioResult,
  DataTableParams,
  DataTableResult,
  DataTableCellWrite,
  ScenarioCellOutput,
} from './what-if';

export interface ExtendedSnapshot {
  capabilities: CapabilityDescriptor[];
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
}

function normalizeGoalSeekValue(value: number, tolerance: number): number {
  if (!Number.isFinite(value)) return value;
  const nearestInt = Math.round(value);
  if (Math.abs(value - nearestInt) <= tolerance * 100) return nearestInt;
  const decimals = Math.max(0, Math.ceil(-Math.log10(tolerance)));
  const factor = 10 ** Math.min(decimals, 10);
  return Math.round(value * factor) / factor;
}

function readCellValue(
  formula: FormulaEngine,
  sheetId: string,
  row: number,
  column: number,
): number | string | boolean | null {
  const value = formula.getCellValue({ sheetId, row, column });
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') return value;
  if (value == null) return null;
  if (Array.isArray(value)) {
    const first = Array.isArray(value[0]) ? value[0]?.[0] : value[0];
    if (typeof first === 'number' || typeof first === 'string' || typeof first === 'boolean') return first;
    return first == null ? null : String(first);
  }
  return String(value);
}

function readNumericCellValue(
  formula: FormulaEngine,
  sheetId: string,
  row: number,
  column: number,
): number {
  const value = formula.getCellValue({ sheetId, row, column });
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (Array.isArray(value)) {
    const first = Array.isArray(value[0]) ? value[0]?.[0] : value[0];
    if (typeof first === 'number' && Number.isFinite(first)) return first;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function syncScalarCellValue(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  sheetId: string,
  row: number,
  column: number,
  value: number | string | boolean | null,
): void {
  const sheet = workbook.getSheet(sheetId);
  const address = { sheetId, row, column };
  const existing = sheet.cells.get(row, column);
  if (existing?.formula) {
    sheet.cells.set(row, column, { ...existing, value: value as number | string | boolean });
    formula.setValue(address, value);
    return;
  }
  if (value == null) {
    sheet.cells.delete(row, column);
    formula.clearCell(address);
    return;
  }
  sheet.cells.set(row, column, { value });
  formula.setValue(address, value);
}

function readSheetScalar(
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  column: number,
): number | string | boolean | null {
  const cell = workbook.getSheet(sheetId).cells.get(row, column);
  if (!cell) return null;
  if (cell.value != null) return cell.value as number | string | boolean;
  if (cell.formula) return cell.formula;
  return null;
}

export function runGoalSeek(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  sheetId: string,
  params: GoalSeekParams,
): GoalSeekResult {
  const sheet = workbook.getSheet(sheetId);
  const maxIterations = params.maxIterations ?? 64;
  const tolerance = params.tolerance ?? 1e-6;
  const originalChanging = sheet.cells.get(params.byChangingCell.row, params.byChangingCell.column);
  const originalValue = originalChanging?.value;

  let low = -1_000_000;
  let high = 1_000_000;
  let bestGuess = 0;
  let bestDelta = Number.POSITIVE_INFINITY;
  let iterations = 0;

  let converged = false;

  const writeGuess = (guess: number): number => {
    const changingAddress = {
      sheetId,
      row: params.byChangingCell.row,
      column: params.byChangingCell.column,
    };
    sheet.cells.set(params.byChangingCell.row, params.byChangingCell.column, { value: guess });
    formula.setValue(changingAddress, guess);
    formula.recalculate({ sheetId, row: params.setCell.row, column: params.setCell.column });
    return readNumericCellValue(formula, sheetId, params.setCell.row, params.setCell.column);
  };

  const restoreChangingCell = (): void => {
    const changingAddress = {
      sheetId,
      row: params.byChangingCell.row,
      column: params.byChangingCell.column,
    };
    if (originalChanging) {
      sheet.cells.set(params.byChangingCell.row, params.byChangingCell.column, structuredClone(originalChanging));
      if (originalChanging.formula) {
        formula.setFormula(changingAddress, originalChanging.formula);
      } else {
        formula.setValue(changingAddress, (originalValue ?? null) as number | string | boolean | null);
      }
    } else {
      sheet.cells.delete(params.byChangingCell.row, params.byChangingCell.column);
      formula.clearCell(changingAddress);
    }
    formula.recalculate({ sheetId, row: params.setCell.row, column: params.setCell.column });
  };

  try {
    for (let index = 0; index < maxIterations; index += 1) {
      iterations = index + 1;
      const guess = (low + high) / 2;
      const current = writeGuess(guess);
      if (!Number.isFinite(current)) {
        return {
          kind: 'goal-seek',
          status: 'not-converged',
          iterations,
          message: 'Goal cell does not evaluate to a numeric value',
        };
      }
      const delta = current - params.toValue;
      if (Math.abs(delta) < tolerance) {
        converged = true;
        const changingCellValue = normalizeGoalSeekValue(guess, tolerance);
        writeGuess(changingCellValue);
        return {
          kind: 'goal-seek',
          status: 'converged',
          finalValue: params.toValue,
          changingCellValue,
          iterations,
        };
      }
      if (Math.abs(delta) < bestDelta) {
        bestDelta = Math.abs(delta);
        bestGuess = guess;
      }
      if (delta < 0) low = guess;
      else high = guess;
    }

    const finalValue = writeGuess(bestGuess);
    return {
      kind: 'goal-seek',
      status: 'not-converged',
      finalValue: Number.isFinite(finalValue) ? finalValue : undefined,
      changingCellValue: bestGuess,
      iterations,
      message: `Reached iteration limit with delta ${bestDelta.toFixed(6)}`,
    };
  } finally {
    if (!converged) {
      restoreChangingCell();
    }
  }
}

export function runScenario(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  sheetId: string,
  scenario: ScenarioDefinition,
): ScenarioResult {
  if (scenario.changingCells.length === 0) {
    return {
      kind: 'scenario',
      status: 'failed',
      scenarioId: scenario.id,
      message: 'Scenario has no changing cells',
      outputs: [],
    };
  }

  for (const cell of scenario.changingCells) {
    syncScalarCellValue(workbook, formula, sheetId, cell.row, cell.column, cell.value);
  }
  formula.recalculate();

  const resultCells = scenario.resultCells?.length
    ? scenario.resultCells
    : scenario.changingCells.map((cell) => ({ row: cell.row, column: cell.column }));

  const outputs: ScenarioCellOutput[] = resultCells.map((cell) => ({
    row: cell.row,
    column: cell.column,
    value: readCellValue(formula, sheetId, cell.row, cell.column),
  }));

  return {
    kind: 'scenario',
    status: 'completed',
    scenarioId: scenario.id,
    message: `Scenario "${scenario.name}" applied with ${scenario.changingCells.length} changing cell(s)`,
    outputs,
  };
}

export function runDataTable(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  sheetId: string,
  params: DataTableParams,
): DataTableResult {
  const hasRowInput = Boolean(params.rowInputCell);
  const hasColumnInput = Boolean(params.columnInputCell);
  if (hasRowInput === hasColumnInput) {
    return {
      kind: 'data-table',
      status: 'failed',
      message: 'Data table requires exactly one of rowInputCell or columnInputCell',
      filledCells: 0,
      writes: [],
    };
  }

  const { startRow, startColumn, endRow, endColumn } = params.tableRange;
  if (endRow < startRow || endColumn < startColumn) {
    return {
      kind: 'data-table',
      status: 'failed',
      message: 'Invalid data table range',
      filledCells: 0,
      writes: [],
    };
  }

  const writes: DataTableCellWrite[] = [];

  if (params.rowInputCell) {
    const inputCell = params.rowInputCell;
    const formulaRow = startRow;
    const formulaColumn = startColumn;
    for (let row = startRow + 1; row <= endRow; row += 1) {
      const rawInput = readSheetScalar(workbook, sheetId, row, startColumn);
      if (rawInput == null || typeof rawInput === 'boolean') {
        return {
          kind: 'data-table',
          status: 'failed',
          message: `Missing input value at row ${row + 1}, column ${startColumn + 1}`,
          filledCells: writes.length,
          writes,
        };
      }
      syncScalarCellValue(workbook, formula, sheetId, inputCell.row, inputCell.column, rawInput);
      formula.recalculate({ sheetId, row: formulaRow, column: formulaColumn });
      const result = readCellValue(formula, sheetId, formulaRow, formulaColumn);
      const targetColumn = startColumn + 1;
      syncScalarCellValue(workbook, formula, sheetId, row, targetColumn, result);
      writes.push({ row, column: targetColumn, value: result });
    }
  } else if (params.columnInputCell) {
    const inputCell = params.columnInputCell;
    const formulaRow = startRow + 1;
    const formulaColumn = startColumn;
    for (let column = startColumn + 1; column <= endColumn; column += 1) {
      const rawInput = readSheetScalar(workbook, sheetId, startRow, column);
      if (rawInput == null || typeof rawInput === 'boolean') {
        return {
          kind: 'data-table',
          status: 'failed',
          message: `Missing input value at row ${startRow + 1}, column ${column + 1}`,
          filledCells: writes.length,
          writes,
        };
      }
      syncScalarCellValue(workbook, formula, sheetId, inputCell.row, inputCell.column, rawInput);
      formula.recalculate({ sheetId, row: formulaRow, column: formulaColumn });
      const result = readCellValue(formula, sheetId, formulaRow, formulaColumn);
      syncScalarCellValue(workbook, formula, sheetId, formulaRow, column, result);
      writes.push({ row: formulaRow, column, value: result });
    }
  }

  return {
    kind: 'data-table',
    status: 'completed',
    message: `Data table filled ${writes.length} result cell(s)`,
    filledCells: writes.length,
    writes,
  };
}

export function evaluateCapability(
  registry: CapabilityRegistry,
  id: PlatformCapability,
): { canEnable: boolean; reason?: string } {
  return registry.evaluate(id);
}

export function summarizeGoalSeekResult(result: GoalSeekResult): string {
  if (result.status === 'converged') {
    return `Goal Seek converged in ${result.iterations} iteration(s): target ${result.finalValue}`;
  }
  return result.message ?? `Goal Seek did not converge after ${result.iterations} iteration(s)`;
}

export function summarizeScenarioResult(result: ScenarioResult): string {
  if (result.status === 'completed') {
    return `${result.message}; ${result.outputs.length} output cell(s) evaluated`;
  }
  return result.message;
}

export function summarizeDataTableResult(result: DataTableResult): string {
  if (result.status === 'completed') {
    return result.message;
  }
  return result.message;
}
