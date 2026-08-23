export type WhatIfKind = 'goal-seek' | 'data-table' | 'scenario';

export interface GoalSeekParams {
  setCell: { row: number; column: number };
  toValue: number;
  byChangingCell: { row: number; column: number };
  maxIterations?: number;
  tolerance?: number;
}

export interface GoalSeekResult {
  kind: 'goal-seek';
  status: 'converged' | 'not-converged' | 'stub';
  finalValue?: number;
  changingCellValue?: number;
  iterations: number;
  message?: string;
}

export interface DataTableParams {
  rowInputCell?: { row: number; column: number };
  columnInputCell?: { row: number; column: number };
  tableRange: { startRow: number; startColumn: number; endRow: number; endColumn: number };
}

export interface DataTableResult {
  kind: 'data-table';
  status: 'stub';
  message: string;
}

export interface ScenarioDefinition {
  id: string;
  name: string;
  changingCells: Array<{ row: number; column: number; value: number | string }>;
}

export interface ScenarioResult {
  kind: 'scenario';
  status: 'stub';
  scenarioId: string;
  message: string;
}

/** What-if 基础 stub — Goal Seek / Data Table / Scenario */
export class WhatIfService {
  goalSeek(params: GoalSeekParams): GoalSeekResult {
    return {
      kind: 'goal-seek',
      status: 'stub',
      iterations: 0,
      message: `Goal Seek stub: set (${params.setCell.row},${params.setCell.column}) to ${params.toValue} by changing (${params.byChangingCell.row},${params.byChangingCell.column})`,
    };
  }

  dataTable(_params: DataTableParams): DataTableResult {
    return {
      kind: 'data-table',
      status: 'stub',
      message: 'Data Table requires formula recalc engine integration (M18 stub)',
    };
  }

  runScenario(scenario: ScenarioDefinition): ScenarioResult {
    return {
      kind: 'scenario',
      status: 'stub',
      scenarioId: scenario.id,
      message: `Scenario "${scenario.name}" stub — ${scenario.changingCells.length} changing cells`,
    };
  }
}
