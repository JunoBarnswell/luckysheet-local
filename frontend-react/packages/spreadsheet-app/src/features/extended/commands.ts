import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { PlatformCapability } from './index';
import type { GoalSeekParams, ScenarioDefinition, DataTableParams } from './what-if';

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

export function registerExtendedCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'extended.whatIf.goalSeek',
    execute(params: ExtendedGoalSeekCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [{
          sheetId: params.sheetId ?? context.workbook.activeSheetId,
          startRow: params.setCell.row,
          endRow: params.setCell.row,
          startColumn: params.setCell.column,
          endColumn: params.setCell.column,
        }],
      };
    },
  });

  registry.registerCommand({
    id: 'extended.whatIf.scenario',
    execute(params: ExtendedScenarioCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'extended.whatIf.dataTable',
    execute(params: ExtendedDataTableCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'extended.capability.evaluate',
    execute(params: ExtendedCapabilityEvaluateParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });
}
