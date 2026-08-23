import type { WorkbookModel } from '@react-sheets/core-model';
import type { FormulaEngine } from '@react-sheets/formula-engine';
import {
  CapabilityRegistry,
  type CapabilityDescriptor,
  type PlatformCapability,
} from './index';
import {
  planDataTable,
  planGoalSeek,
  planScenario,
  type GoalSeekParams,
  type GoalSeekResult,
  type ScenarioDefinition,
  type ScenarioResult,
  type DataTableParams,
  type DataTableResult,
} from './what-if';

export interface ExtendedSnapshot {
  capabilities: CapabilityDescriptor[];
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
}

/**
 * Read-only host API.  What-if calculation is delegated to the isolated
 * planner; these functions intentionally do not write WorkbookModel or the
 * caller-owned FormulaEngine.  Hosts apply the returned plan through the
 * registered extended command transaction.
 */
export function runGoalSeek(
  workbook: WorkbookModel,
  _formula: FormulaEngine,
  sheetId: string,
  params: GoalSeekParams,
): GoalSeekResult {
  return planGoalSeek(workbook, sheetId, params).result;
}

export function runScenario(
  workbook: WorkbookModel,
  _formula: FormulaEngine,
  sheetId: string,
  scenario: ScenarioDefinition,
): ScenarioResult {
  return planScenario(workbook, sheetId, scenario).result;
}

export function runDataTable(
  workbook: WorkbookModel,
  _formula: FormulaEngine,
  sheetId: string,
  params: DataTableParams,
): DataTableResult {
  return planDataTable(workbook, sheetId, params).result;
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
  return result.message;
}
