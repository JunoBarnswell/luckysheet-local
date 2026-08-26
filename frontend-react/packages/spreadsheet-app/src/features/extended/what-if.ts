import type { CellData, CellValue, WorkbookModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';

export type WhatIfKind = 'goal-seek' | 'scenario';

export interface WhatIfPlanMetadata {
  schema: 'WhatIfPlan';
  kind: WhatIfKind;
  /** Hash of the authoritative workbook state used to build this plan. */
  sourceRevision: string;
  /** Hash of definition + writes; useful for replay/audit, not authorization. */
  planHash: string;
  definition: unknown;
  writeCount: number;
  deterministic: true;
}

export interface GoalSeekParams {
  setCell: { row: number; column: number };
  toValue: number;
  byChangingCell: { row: number; column: number };
  maxIterations?: number;
  tolerance?: number;
}

export interface GoalSeekResult {
  kind: 'goal-seek';
  status: 'converged' | 'not-converged' | 'failed';
  finalValue?: number;
  changingCellValue?: number;
  iterations: number;
  message?: string;
}

export interface ScenarioDefinition {
  id: string;
  name: string;
  changingCells: Array<{ row: number; column: number; value: number | string }>;
  resultCells?: Array<{ row: number; column: number }>;
}

export interface ScenarioCellOutput {
  row: number;
  column: number;
  value: number | string | boolean | null;
}

export interface ScenarioResult {
  kind: 'scenario';
  status: 'completed' | 'failed';
  scenarioId: string;
  message: string;
  outputs: ScenarioCellOutput[];
}

/** A deterministic write plan; applying it is the only side effect of a command. */
export interface WhatIfCellWrite {
  sheetId: string;
  row: number;
  column: number;
  value: CellData;
}

export interface GoalSeekPlan {
  kind: 'goal-seek';
  result: GoalSeekResult;
  writes: WhatIfCellWrite[];
  metadata?: WhatIfPlanMetadata;
}

export interface ScenarioPlan {
  kind: 'scenario';
  result: ScenarioResult;
  writes: WhatIfCellWrite[];
  metadata?: WhatIfPlanMetadata;
}

export type WhatIfPlan = GoalSeekPlan | ScenarioPlan;

type FormulaScalar = CellValue | string;

function scalarValue(value: unknown): FormulaScalar {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) {
    const first = Array.isArray(value[0]) ? value[0]?.[0] : value[0];
    return scalarValue(first ?? null);
  }
  return String(value);
}

function createPlanningFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const firstSheet = workbook.getSheets()[0];
  const engine = new FormulaEngine({ defaultSheetId: firstSheet?.id ?? workbook.primarySheetId });
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula) engine.setFormula(address, cell.formula);
      else engine.setValue(address, (cell.value ?? null) as never);
    });
  }
  engine.setDefinedNameModels(workbook.definedNameModels);
  engine.recalculate();
  return engine;
}

function readFormulaScalar(engine: FormulaEngine, sheetId: string, row: number, column: number): FormulaScalar {
  return scalarValue(engine.getCellValue({ sheetId, row, column }));
}

function isSpillCell(workbook: WorkbookModel, sheetId: string, row: number, column: number): boolean {
  return workbook.getSheet(sheetId).spillRanges.some((spill) => (
    spill.range.startRow <= row
      && row <= spill.range.endRow
      && spill.range.startColumn <= column
      && column <= spill.range.endColumn
  ));
}

function hasArrayResult(engine: FormulaEngine, sheetId: string, row: number, column: number): boolean {
  return Array.isArray(engine.getCellValue({ sheetId, row, column }));
}

function readWorkbookScalar(workbook: WorkbookModel, sheetId: string, row: number, column: number): FormulaScalar {
  const cell = workbook.getSheet(sheetId).cells.get(row, column);
  if (!cell) return null;
  if (cell.value != null) return cell.value;
  return cell.formula ?? null;
}

function cellWrite(sheetId: string, row: number, column: number, value: FormulaScalar): WhatIfCellWrite {
  return { sheetId, row, column, value: { value } };
}

function validCell(sheetId: string, row: number, column: number, workbook: WorkbookModel): boolean {
  const sheet = workbook.getSheet(sheetId);
  return Number.isInteger(row) && Number.isInteger(column) && row >= 0 && row < sheet.rowCount && column >= 0 && column < sheet.columnCount;
}

function normalizeGoalSeekValue(value: number, tolerance: number): number {
  if (!Number.isFinite(value)) return value;
  const nearestInt = Math.round(value);
  if (Math.abs(value - nearestInt) <= tolerance * 100) return nearestInt;
  const decimals = Math.max(0, Math.ceil(-Math.log10(tolerance)));
  const factor = 10 ** Math.min(decimals, 10);
  return Math.round(value * factor) / factor;
}

export function planGoalSeek(workbook: WorkbookModel, sheetId: string, params: GoalSeekParams): GoalSeekPlan {
  const invalid = !validCell(sheetId, params.setCell.row, params.setCell.column, workbook)
    || !validCell(sheetId, params.byChangingCell.row, params.byChangingCell.column, workbook);
  const maxIterations = Number.isFinite(params.maxIterations ?? 64) ? Math.floor(params.maxIterations ?? 64) : 0;
  const tolerance = params.tolerance ?? 1e-6;
  if (invalid || !Number.isFinite(params.toValue) || maxIterations <= 0 || !Number.isFinite(tolerance) || tolerance <= 0) {
    return {
      kind: 'goal-seek',
      result: { kind: 'goal-seek', status: 'failed', iterations: 0, message: 'Invalid Goal Seek parameters' },
      writes: [],
    };
  }

  if (
    isSpillCell(workbook, sheetId, params.setCell.row, params.setCell.column)
    || isSpillCell(workbook, sheetId, params.byChangingCell.row, params.byChangingCell.column)
  ) {
    return {
      kind: 'goal-seek',
      result: { kind: 'goal-seek', status: 'failed', iterations: 0, message: 'Goal Seek cannot write to a spill range' },
      writes: [],
    };
  }

  const engine = createPlanningFormulaEngine(workbook);
  const { setCell, byChangingCell } = params;
  let low = -1_000_000;
  let high = 1_000_000;
  let bestGuess = 0;
  let bestDelta = Number.POSITIVE_INFINITY;
  let iterations = 0;
  let spillDetected = false;
  const writeGuess = (guess: number): number => {
    engine.setValue({ sheetId, row: byChangingCell.row, column: byChangingCell.column }, guess);
    engine.recalculate({ sheetId, row: setCell.row, column: setCell.column });
    if (hasArrayResult(engine, sheetId, setCell.row, setCell.column)) {
      spillDetected = true;
      return Number.NaN;
    }
    const value = readFormulaScalar(engine, sheetId, setCell.row, setCell.column);
    return typeof value === 'number' ? value : Number(value);
  };

  const samples: number[] = [];
  for (let index = 0; index <= 16; index += 1) {
    const guess = low + ((high - low) * index) / 16;
    const value = writeGuess(guess);
    if (!Number.isFinite(value)) {
      return {
        kind: 'goal-seek',
        result: {
          kind: 'goal-seek',
          status: spillDetected ? 'failed' : 'not-converged',
          iterations: 0,
          message: spillDetected ? 'Goal cell produces a spill result; no mutation applied' : 'Goal cell does not evaluate to a numeric value',
        },
        writes: [],
      };
    }
    samples.push(value);
  }
  let direction = 0;
  for (let index = 1; index < samples.length; index += 1) {
    const delta = samples[index]! - samples[index - 1]!;
    if (Math.abs(delta) <= tolerance) continue;
    const nextDirection = Math.sign(delta);
    if (direction !== 0 && nextDirection !== direction) {
      return {
        kind: 'goal-seek',
        result: { kind: 'goal-seek', status: 'not-converged', iterations: 0, message: 'Goal function is non-monotonic; no mutation applied' },
        writes: [],
      };
    }
    direction = nextDirection;
  }

  for (let index = 0; index < maxIterations; index += 1) {
    iterations = index + 1;
    const guess = (low + high) / 2;
    const current = writeGuess(guess);
    if (!Number.isFinite(current)) {
      return {
        kind: 'goal-seek',
        result: {
          kind: 'goal-seek',
          status: spillDetected ? 'failed' : 'not-converged',
          iterations,
          message: spillDetected ? 'Goal cell produces a spill result; no mutation applied' : 'Goal cell does not evaluate to a numeric value',
        },
        writes: [],
      };
    }
    const delta = current - params.toValue;
    if (Math.abs(delta) < tolerance) {
      const changingCellValue = normalizeGoalSeekValue(guess, tolerance);
      writeGuess(changingCellValue);
      return {
        kind: 'goal-seek',
        result: { kind: 'goal-seek', status: 'converged', finalValue: params.toValue, changingCellValue, iterations },
        writes: [cellWrite(sheetId, byChangingCell.row, byChangingCell.column, changingCellValue)],
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
    result: {
      kind: 'goal-seek',
      status: 'not-converged',
      finalValue: Number.isFinite(finalValue) ? finalValue : undefined,
      changingCellValue: bestGuess,
      iterations,
      message: `Reached iteration limit with delta ${bestDelta.toFixed(6)}`,
    },
    writes: [],
  };
}

export function planScenario(workbook: WorkbookModel, sheetId: string, scenario: ScenarioDefinition): ScenarioPlan {
  if (scenario.changingCells.length === 0) {
    return {
      kind: 'scenario',
      result: { kind: 'scenario', status: 'failed', scenarioId: scenario.id, message: 'Scenario has no changing cells', outputs: [] },
      writes: [],
    };
  }
  const seen = new Set<string>();
  const sortedChanges = [...scenario.changingCells].sort((a, b) => a.row - b.row || a.column - b.column);
  for (const cell of sortedChanges) {
    const key = `${cell.row}:${cell.column}`;
    if (!validCell(sheetId, cell.row, cell.column, workbook) || seen.has(key) || isSpillCell(workbook, sheetId, cell.row, cell.column)) {
      return {
        kind: 'scenario',
        result: { kind: 'scenario', status: 'failed', scenarioId: scenario.id, message: 'Scenario contains an invalid, duplicate, or spill changing cell', outputs: [] },
        writes: [],
      };
    }
    seen.add(key);
  }

  const engine = createPlanningFormulaEngine(workbook);
  for (const cell of sortedChanges) engine.setValue({ sheetId, row: cell.row, column: cell.column }, cell.value);
  engine.recalculate();
  const resultCells = scenario.resultCells?.length
    ? scenario.resultCells
    : scenario.changingCells.map((cell) => ({ row: cell.row, column: cell.column }));
  if (resultCells.some((cell) => !validCell(sheetId, cell.row, cell.column, workbook))) {
    return {
      kind: 'scenario',
      result: { kind: 'scenario', status: 'failed', scenarioId: scenario.id, message: 'Scenario contains an invalid result cell', outputs: [] },
      writes: [],
    };
  }
  const outputs = resultCells.map((cell) => ({
    row: cell.row,
    column: cell.column,
    value: readFormulaScalar(engine, sheetId, cell.row, cell.column),
  }));
  if (resultCells.some((cell) => hasArrayResult(engine, sheetId, cell.row, cell.column))) {
    return {
      kind: 'scenario',
      result: { kind: 'scenario', status: 'failed', scenarioId: scenario.id, message: 'Scenario result contains a spill value', outputs: [] },
      writes: [],
    };
  }
  return {
    kind: 'scenario',
    result: {
      kind: 'scenario',
      status: 'completed',
      scenarioId: scenario.id,
      message: `Scenario "${scenario.name}" applied with ${scenario.changingCells.length} changing cell(s)`,
      outputs,
    },
    writes: sortedChanges.map((cell) => cellWrite(sheetId, cell.row, cell.column, cell.value)),
  };
}
