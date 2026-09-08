import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { planGoalSeek, planScenario, type GoalSeekParams, type ScenarioDefinition, type WhatIfPlan, type WhatIfPlanMetadata } from './what-if';

export interface ExtendedGoalSeekCommandParams extends GoalSeekParams { sheetId?: string; }
export interface ExtendedScenarioCommandParams { sheetId?: string; scenario: ScenarioDefinition; }

/** What-if is a read-only preview. A later explicit edit command owns writes. */
function commandPlanResult(plan: WhatIfPlan, definition: unknown, context: CommandContext): CommandResult & { plan: WhatIfPlan } {
  const metadata: WhatIfPlanMetadata = {
    schema: 'WhatIfPlan', kind: plan.kind,
    sourceRevision: `${context.workbook.unitId}:${context.workbook.revision}`,
    planHash: hashValue({ definition, writes: plan.writes }),
    definition: structuredClone(definition), writeCount: plan.writes.length, deterministic: true,
  };
  plan.metadata = metadata;
  return {
    operationId: context.operationId, mutationCount: 0,
    affectedRanges: plan.writes.map((write) => ({ sheetId: write.sheetId, startRow: write.row, endRow: write.row, startColumn: write.column, endColumn: write.column })),
    plan,
  };
}

export function registerExtendedCommands(registry: CommandRegistry): void {
  registry.registerCommand<ExtendedGoalSeekCommandParams>({
    id: 'extended.whatIf.goalSeek',
    execute(params, context): CommandResult & { plan: WhatIfPlan } {
      const sheetId = params.sheetId ?? context.workbook.primarySheetId;
      const formula = new FormulaEngine({ unitId: context.workbook.unitId, revision: () => context.workbook.revision, defaultSheetId: sheetId });
      return commandPlanResult(planGoalSeek(context.workbook, formula, sheetId, params), params, context);
    },
  });
  registry.registerCommand<ExtendedScenarioCommandParams>({
    id: 'extended.whatIf.scenario',
    execute(params, context): CommandResult & { plan: WhatIfPlan } {
      const sheetId = params.sheetId ?? context.workbook.primarySheetId;
      const formula = new FormulaEngine({ unitId: context.workbook.unitId, revision: () => context.workbook.revision, defaultSheetId: sheetId });
      return commandPlanResult(planScenario(context.workbook, formula, sheetId, params.scenario), params.scenario, context);
    },
  });
}

function hashValue(value: unknown): string {
  const text = JSON.stringify(value);
  let hash = 2166136261;
  for (let index = 0; index < text.length; index += 1) { hash ^= text.charCodeAt(index); hash = Math.imul(hash, 16777619); }
  return (hash >>> 0).toString(16).padStart(8, '0');
}
