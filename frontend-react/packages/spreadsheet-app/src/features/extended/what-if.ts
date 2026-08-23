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
  status: 'converged' | 'not-converged' | 'failed';
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

export interface DataTableCellWrite {
  row: number;
  column: number;
  value: number | string | boolean | null;
}

export interface DataTableResult {
  kind: 'data-table';
  status: 'completed' | 'failed';
  message: string;
  filledCells: number;
  writes: DataTableCellWrite[];
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
