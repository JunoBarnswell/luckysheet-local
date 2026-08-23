import type { CellData, RangeRef } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import {
  buildFacadePlan,
  parseFacadeScript,
  type FacadeCellOperation,
  type FacadeProgram,
} from './dsl';
import { DEFAULT_SANDBOX_POLICY, ScriptSandbox } from './sandbox';

export interface AutomationRunParams {
  source: string;
  label?: string;
  /** Internal parsed form; hosts must still provide source, which is parsed again. */
  program?: FacadeProgram;
}

export interface AutomationPlanResult extends CommandResult {
  plan: {
    schema: 'AutomationPlanV1';
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

function registerCellMutationHandlers(registry: CommandRegistry): void {
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
      params: { row: operation.row, column: operation.column, previous: undefined },
      affectedRanges,
      inverse: [{
        id: 'cell.restore',
        unitId: context.workbook.unitId,
        sheetId: operation.sheetId,
        params: { row: operation.row, column: operation.column, previous: structuredClone(previous) },
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
      params: { row: operation.row, column: operation.column, previous: previous ? structuredClone(previous) : undefined },
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
    sheetId: context.workbook.activeSheetId,
    params: { recording: next },
    affectedRanges: [],
    inverse: [{
      id: 'automation.recording.changed',
      unitId: context.workbook.unitId,
      sheetId: context.workbook.activeSheetId,
      params: { recording: previous },
      affectedRanges: [],
    }],
    apply: () => { state.recording = next; },
  });
}

export function registerAutomationCommands(registry: CommandRegistry, options: AutomationCommandOptions = {}): void {
  registerCellMutationHandlers(registry);
  const state: RecordingState = { recording: false };
  const sandbox = options.sandbox ?? new ScriptSandbox(DEFAULT_SANDBOX_POLICY);
  registry.registerMutation<{ recording: boolean }>('automation.recording.changed', (item) => {
    state.recording = item.params.recording;
  });

  registry.registerCommand<AutomationRunParams>({
    id: 'automation.run',
    execute(params, context): AutomationPlanResult {
      if (!params || typeof params.source !== 'string') throw new Error('Automation source is required');
      if (Object.prototype.hasOwnProperty.call(params, 'program')) {
        throw new Error('Automation command accepts source only; program payload is not serializable');
      }
      // Always parse the source supplied to the command. A caller cannot
      // smuggle an AST with executable JavaScript through an out-of-band hint.
      const program = sandbox.parse(params.source);
      const plan = buildFacadePlan(context.workbook, program);
      sandbox.assertPlanAllowed(plan);
      const deadline = Date.now() + sandbox.getTimeoutMs();
      for (const operation of plan.operations) {
        if (Date.now() > deadline) throw new Error(`Automation exceeded ${sandbox.getTimeoutMs()}ms execution limit`);
        applyPlannedCellOperation(operation, context);
      }
      return {
        operationId: context.operationId,
        mutationCount: plan.operations.length,
        affectedRanges: [...plan.affectedRanges],
        plan: {
          schema: 'AutomationPlanV1',
          kind: 'facade-dsl',
          statements: plan.statements.length,
          operations: plan.operations.length,
          sourceHash: hashSource(params.source),
          sourceLength: program.sourceLength,
          cellCount: program.cellCount,
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

function hashSource(source: string): string {
  // Deterministic, non-secret plan identity. This is metadata only and is not
  // used as an authorization or integrity primitive.
  let hash = 2166136261;
  for (let index = 0; index < source.length; index += 1) {
    hash ^= source.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}
