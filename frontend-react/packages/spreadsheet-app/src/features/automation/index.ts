import type { WorkbookModel, CellData } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import { ScriptSandbox } from './sandbox';
import {
  buildFacadePlan,
  parseA1Range,
  parseFacadeScript,
  type FacadePlan,
  type FacadeProgram,
} from './dsl';

function normalizeFacadeCellValue(value: unknown): CellData {
  if (value != null && typeof value === 'object' && 'value' in (value as object)) {
    return value as CellData;
  }
  return { value: value as CellData['value'] };
}

export interface FacadeRange {
  setValues(values: unknown[][]): void;
  setFontWeight(weight: 'normal' | 'bold'): void;
  clear(): void;
}

export interface SpreadsheetFacade {
  getActiveSheet(): { getName(): string; getRange(a1: string): FacadeRange };
  getWorkbook(): { getName(): string };
}

/** Facade 脚本运行时 — 脚本只允许调 Facade */
export class FacadeScriptRuntime {
  constructor(
    private readonly workbook: WorkbookModel,
    private readonly runtime: CommandRuntime,
  ) {}

  createFacade(): SpreadsheetFacade {
    const workbook = this.workbook;
    const runtime = this.runtime;
    const activeSheetId = () => workbook.activeSheetId;

    const createRange = (a1: string): FacadeRange => ({
      setValues(values: unknown[][]) {
        const ref = parseA1Range(a1);
        runtime.execute('sheet.range.set', {
          sheetId: activeSheetId(),
          startRow: ref.startRow,
          startColumn: ref.startColumn,
          values: values.map((row) => row.map(normalizeFacadeCellValue)),
        });
      },
      setFontWeight(weight: 'normal' | 'bold') {
        const ref = parseA1Range(a1);
        const sheetId = activeSheetId();
        runtime.execute('sheet.style.set', {
          sheetId,
          range: {
            sheetId,
            startRow: ref.startRow,
            endRow: ref.endRow,
            startColumn: ref.startColumn,
            endColumn: ref.endColumn,
          },
          style: { bold: weight === 'bold' },
        });
      },
      clear() {
        const ref = parseA1Range(a1);
        const sheetId = activeSheetId();
        runtime.execute('sheet.range.clear', {
          sheetId,
          range: {
            sheetId,
            startRow: ref.startRow,
            endRow: ref.endRow,
            startColumn: ref.startColumn,
            endColumn: ref.endColumn,
          },
        });
      },
    });

    return {
      getActiveSheet() {
        const sheet = workbook.getSheet(activeSheetId());
        return {
          getName: () => sheet.name,
          getRange: createRange,
        };
      },
      getWorkbook() {
        return { getName: () => workbook.name };
      },
    };
  }

  /**
   * Parse, validate, and execute a Facade DSL program.  Parsing and range
   * validation are completed before the first mutation.  When the feature
   * command is registered, CommandRuntime owns the entire transaction;
   * otherwise the already validated plan is applied through existing sheet
   * commands for the standalone kernel use case.
   */
  runScript(source: string, sandbox: ScriptSandbox): ScriptRunResult {
    const started = Date.now();
    try {
      const program = sandbox.parse(source);
      const plan = buildFacadePlan(this.workbook, program);
      const result = this.executePlan(source, program, plan);
      return {
        ok: true,
        durationMs: Date.now() - started,
        mutationCount: result.mutationCount,
        plan,
      };
    } catch (error) {
      return { ok: false, durationMs: Date.now() - started, error: error instanceof Error ? error.message : String(error) };
    }
  }

  /** Public deterministic plan API for hosts that own transaction execution. */
  planScript(source: string, sandbox = new ScriptSandbox()): FacadePlan {
    return buildFacadePlan(this.workbook, sandbox.parse(source));
  }

  private executePlan(source: string, program: FacadeProgram, plan: FacadePlan): { mutationCount: number } {
    if (this.runtime.registry.hasCommand('automation.run')) {
      const result = this.runtime.execute('automation.run', { source, program });
      return { mutationCount: result.mutationCount };
    }

    const commandIds = new Set<string>();
    for (const operation of plan.operations) {
      commandIds.add(operation.kind === 'set-cell' ? 'sheet.cell.set' : operation.kind === 'set-style' ? 'sheet.style.set' : 'sheet.range.clear');
    }
    for (const commandId of commandIds) {
      if (!this.runtime.registry.hasCommand(commandId)) throw new Error(`Automation requires registered command: ${commandId}`);
    }

    let mutationCount = 0;
    for (const statement of plan.statements) {
      const sheetId = this.workbook.activeSheetId;
      if (statement.kind === 'set-values') {
        const result = this.runtime.execute('sheet.range.set', {
          sheetId,
          startRow: statement.range.startRow,
          startColumn: statement.range.startColumn,
          values: statement.values.map((row) => row.map(normalizeFacadeCellValue)),
        });
        mutationCount += result.mutationCount;
      } else if (statement.kind === 'set-font-weight') {
        const result = this.runtime.execute('sheet.style.set', {
          sheetId,
          range: {
            sheetId,
            startRow: statement.range.startRow,
            endRow: statement.range.endRow,
            startColumn: statement.range.startColumn,
            endColumn: statement.range.endColumn,
          },
          style: { bold: statement.weight === 'bold' },
        });
        mutationCount += result.mutationCount;
      } else {
        const result = this.runtime.execute('sheet.range.clear', {
          sheetId,
          range: {
            sheetId,
            startRow: statement.range.startRow,
            endRow: statement.range.endRow,
            startColumn: statement.range.startColumn,
            endColumn: statement.range.endColumn,
          },
        });
        mutationCount += result.mutationCount;
      }
    }
    return { mutationCount };
  }
}

export interface ScriptRunResult {
  ok: boolean;
  durationMs: number;
  mutationCount?: number;
  plan?: FacadePlan;
  error?: string;
}

export { CommandRecorder, type RecordedStatement } from './command-recorder';
export { ScriptSandbox, DEFAULT_SANDBOX_POLICY, type SandboxPolicy } from './sandbox';
export { registerAutomationCommands } from './commands';
export {
  buildFacadePlan,
  parseA1Range,
  parseAndBuildFacadePlan,
  parseFacadeScript,
  DEFAULT_FACADE_DSL_LIMITS,
  type A1Range,
  type FacadeCellOperation,
  type FacadeDslLimits,
  type FacadePlan,
  type FacadeProgram,
  type FacadeStatement,
} from './dsl';
