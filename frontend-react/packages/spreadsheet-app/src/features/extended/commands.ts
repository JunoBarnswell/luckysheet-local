import type { CellData, RangeRef } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import { CapabilityRegistry, type PlatformCapability } from './index';
import {
  planDataTable,
  planGoalSeek,
  planScenario,
  type DataTableParams,
  type GoalSeekParams,
  type ScenarioDefinition,
  type WhatIfCellWrite,
  type WhatIfPlan,
} from './what-if';

export interface ExtendedGoalSeekCommandParams extends GoalSeekParams {
  sheetId?: string;
}

export interface ExtendedScenarioCommandParams {
  sheetId?: string;
  scenario: ScenarioDefinition;
}

export interface ExtendedDataTableCommandParams extends DataTableParams {
  sheetId?: string;
}

export interface ExtendedCapabilityEvaluateParams {
  capability: PlatformCapability;
}

function cellRange(sheetId: string, row: number, column: number): RangeRef[] {
  return [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
}

function ensureCellMutations(registry: CommandRegistry): void {
  if (!registry.hasMutation('cell.set')) {
    registry.registerMutation<{ sheetId: string; row: number; column: number; value: CellData }>('cell.set', (item, context) => {
      const params = item.params;
      context.workbook.getSheet(params.sheetId).cells.set(params.row, params.column, structuredClone(params.value));
    });
  }
  if (!registry.hasMutation('cell.restore')) {
    registry.registerMutation<{ row: number; column: number; previous?: CellData }>('cell.restore', (item, context) => {
      const params = item.params;
      const sheet = context.workbook.getSheet(item.sheetId);
      if (params.previous) sheet.cells.set(params.row, params.column, structuredClone(params.previous));
      else sheet.cells.delete(params.row, params.column);
    });
  }
}

function applyWrite(write: WhatIfCellWrite, context: CommandContext): void {
  const sheet = context.workbook.getSheet(write.sheetId);
  const previous = sheet.cells.get(write.row, write.column);
  const affectedRanges = cellRange(write.sheetId, write.row, write.column);
  context.applyMutation({
    id: 'cell.set',
    unitId: context.workbook.unitId,
    sheetId: write.sheetId,
    params: { sheetId: write.sheetId, row: write.row, column: write.column, value: structuredClone(write.value) },
    affectedRanges,
    inverse: [{
      id: 'cell.restore',
      unitId: context.workbook.unitId,
      sheetId: write.sheetId,
      params: { row: write.row, column: write.column, previous: previous ? structuredClone(previous) : undefined },
      affectedRanges,
    }],
    apply: () => sheet.cells.set(write.row, write.column, structuredClone(write.value)),
  });
}

function commandPlanResult(plan: WhatIfPlan, context: CommandContext): CommandResult & { plan: WhatIfPlan } {
  const affectedRanges = plan.writes.map((write) => ({
    sheetId: write.sheetId,
    startRow: write.row,
    endRow: write.row,
    startColumn: write.column,
    endColumn: write.column,
  }));
  for (const write of plan.writes) applyWrite(write, context);
  return {
    operationId: context.operationId,
    mutationCount: plan.writes.length,
    affectedRanges,
    plan,
  };
}

export function registerExtendedCommands(
  registry: CommandRegistry,
  capabilities = new CapabilityRegistry(),
): void {
  ensureCellMutations(registry);

  registry.registerCommand<ExtendedGoalSeekCommandParams>({
    id: 'extended.whatIf.goalSeek',
    execute(params, context): CommandResult & { plan: WhatIfPlan } {
      const sheetId = params.sheetId ?? context.workbook.activeSheetId;
      return commandPlanResult(planGoalSeek(context.workbook, sheetId, params), context);
    },
  });

  registry.registerCommand<ExtendedScenarioCommandParams>({
    id: 'extended.whatIf.scenario',
    execute(params, context): CommandResult & { plan: WhatIfPlan } {
      const sheetId = params.sheetId ?? context.workbook.activeSheetId;
      return commandPlanResult(planScenario(context.workbook, sheetId, params.scenario), context);
    },
  });

  registry.registerCommand<ExtendedDataTableCommandParams>({
    id: 'extended.whatIf.dataTable',
    execute(params, context): CommandResult & { plan: WhatIfPlan } {
      const sheetId = params.sheetId ?? context.workbook.activeSheetId;
      return commandPlanResult(planDataTable(context.workbook, sheetId, params), context);
    },
  });

  registry.registerCommand<ExtendedCapabilityEvaluateParams>({
    id: 'extended.capability.evaluate',
    execute(params, context): CommandResult & { capability: ReturnType<CapabilityRegistry['evaluate']> } {
      const evaluation = capabilities.evaluate(params.capability);
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
        capability: evaluation,
      };
    },
  });
}
