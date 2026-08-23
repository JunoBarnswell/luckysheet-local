import type { CellData, RangeRef } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import {
  checkFacadeExecution,
  validateFacadePlan,
  type FacadeCellOperation,
  type FacadePlan,
} from './dsl';
import { DEFAULT_SANDBOX_POLICY, ScriptSandbox } from './sandbox';

export interface AutomationRunParams {
  plan: FacadePlan;
  label?: string;
}

export interface AutomationPlanResult extends CommandResult {
  plan: {
    schema: 'AutomationPlan';
    kind: 'facade-dsl';
    statements: number;
    operations: number;
    sourceHash: string;
    sourceLength: number;
    cellCount: number;
    affectedRanges: readonly RangeRef[];
    limits: { maxSourceLength: number; maxStatements: number; maxCells: number; maxOperations: number; maxDurationMs: number };
    serializable: true;
  };
}

export interface AutomationCommandOptions {
  sandbox?: ScriptSandbox;
}

interface RecordingState {
  recording: boolean;
}

function cellRange(sheetId: string, row: number, column: number): RangeRef[] {
  return [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
}

function isRecordingMutation(value: unknown): value is { recording: boolean } {
  return typeof value === 'object'
    && value !== null
    && !Array.isArray(value)
    && typeof (value as { recording?: unknown }).recording === 'boolean';
}

function applyPlannedCellOperation(operation: FacadeCellOperation, context: CommandContext): void {
  const sheet = context.workbook.getSheet(operation.sheetId);
  const previous = sheet.cells.get(operation.row, operation.column);
  const affectedRanges = cellRange(operation.sheetId, operation.row, operation.column);

  if (operation.kind === 'clear-cell') {
    if (!previous) return;
    context.applyMutation({
      id: 'cell.restore',
      unitId: context.workbook.unitId,
      sheetId: operation.sheetId,
        params: { sheetId: operation.sheetId, row: operation.row, column: operation.column, previous: undefined },
      affectedRanges,
      inverse: [{
        id: 'cell.restore',
        unitId: context.workbook.unitId,
        sheetId: operation.sheetId,
          params: { sheetId: operation.sheetId, row: operation.row, column: operation.column, previous: structuredClone(previous) },
        affectedRanges,
      }],
      apply: () => sheet.cells.delete(operation.row, operation.column),
    });
    return;
  }

  const value = operation.kind === 'set-cell'
    ? structuredClone(operation.value as CellData)
    : {
      ...(previous ? structuredClone(previous) : { value: null }),
      style: { ...(previous?.style ?? {}), ...operation.style },
    } as CellData;
  context.applyMutation({
    id: 'cell.set',
    unitId: context.workbook.unitId,
    sheetId: operation.sheetId,
    params: { sheetId: operation.sheetId, row: operation.row, column: operation.column, value },
    affectedRanges,
    inverse: [{
      id: 'cell.restore',
      unitId: context.workbook.unitId,
      sheetId: operation.sheetId,
      params: { sheetId: operation.sheetId, row: operation.row, column: operation.column, previous: previous ? structuredClone(previous) : undefined },
      affectedRanges,
    }],
    apply: () => sheet.cells.set(operation.row, operation.column, structuredClone(value)),
  });
}

function applyRecordingState(state: RecordingState, next: boolean, context: CommandContext): void {
  const previous = state.recording;
  context.applyMutation({
    id: 'automation.recording.changed',
    unitId: context.workbook.unitId,
    sheetId: context.workbook.primarySheetId,
    params: { recording: next },
    affectedRanges: [],
    inverse: [{
      id: 'automation.recording.changed',
      unitId: context.workbook.unitId,
      sheetId: context.workbook.primarySheetId,
      params: { recording: previous },
      affectedRanges: [],
    }],
    apply: () => { state.recording = next; },
  });
}

export function registerAutomationCommands(registry: CommandRegistry, options: AutomationCommandOptions = {}): void {
  const state: RecordingState = { recording: false };
  const sandbox = options.sandbox ?? new ScriptSandbox(DEFAULT_SANDBOX_POLICY);
  registry.registerMutation<{ recording: boolean }>({
    id: 'automation.recording.changed',
    handler: (item) => {
      state.recording = item.params.recording;
    },
    metadata: {
      schema: { name: 'AutomationRecordingMutation', validate: isRecordingMutation },
      permission: { capability: 'automation.recording.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['automation.recording.changed'], minCount: 1, maxCount: 1 },
    },
  });

  registry.registerCommand<AutomationRunParams>({
    id: 'automation.run',
    execute(params, context): AutomationPlanResult {
      if (!params || !params.plan) throw new Error('AUTOMATION_PLAN_REQUIRED: a Worker-generated plan is required');
      const sheet = context.workbook.getSheet(context.workbook.primarySheetId);
      const bounds = { sheetId: sheet.id, rowCount: sheet.rowCount, columnCount: sheet.columnCount };
      validateFacadePlan(params.plan, bounds);
      sandbox.assertPlanAllowed(params.plan);
      const deadlineAt = Date.now() + sandbox.getTimeoutMs();
      checkFacadeExecution({ deadlineAt });
      const plan = params.plan;
      for (const operation of plan.operations) {
        checkFacadeExecution({ deadlineAt });
        applyPlannedCellOperation(operation, context);
      }
      // Keep the deadline check inside the CommandRuntime transaction. If the
      // budget expires after the last write, throwing here still triggers the
      // runtime's inverse-mutation rollback.
      checkFacadeExecution({ deadlineAt });
      return {
        operationId: context.operationId,
        mutationCount: plan.operations.length,
        affectedRanges: [...plan.affectedRanges],
        plan: {
          schema: 'AutomationPlan',
          kind: 'facade-dsl',
          statements: plan.statements.length,
          operations: plan.operations.length,
          sourceHash: plan.sourceHash,
          sourceLength: plan.sourceLength,
          cellCount: plan.cellCount,
          affectedRanges: structuredClone(plan.affectedRanges),
          limits: sandbox.getPolicy(),
          serializable: true,
        },
      };
    },
  });

  registry.registerCommand({
    id: 'automation.record.start',
    execute(_params, context): CommandResult {
      if (state.recording) throw new Error('Automation recording is already active');
      applyRecordingState(state, true, context);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand({
    id: 'automation.record.stop',
    execute(_params, context): CommandResult {
      if (!state.recording) throw new Error('Automation recording is not active');
      applyRecordingState(state, false, context);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });
}
