import type { WorkbookModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import { ScriptSandbox } from './sandbox';
import {
  buildFacadePlan,
  parseFacadeScript,
  type FacadePlan,
  type FacadeProgram,
} from './dsl';

/** Facade 脚本运行时 — 脚本只允许调 Facade */
export class FacadeScriptRuntime {
  constructor(
    private readonly workbook: WorkbookModel,
    private readonly runtime: CommandRuntime,
  ) {}

  /**
   * Parse, validate, and execute a Facade DSL program.  Parsing and range
   * validation are completed before the first mutation. CommandRuntime owns
   * the complete transaction; there is deliberately no statement-by-
   * statement fallback path.
   */
  runScript(source: string, sandbox: ScriptSandbox): ScriptRunResult {
    const started = Date.now();
    try {
      const program = sandbox.parse(source);
      const plan = buildFacadePlan(this.workbook, program);
      sandbox.assertPlanAllowed(plan);
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
    const plan = buildFacadePlan(this.workbook, sandbox.parse(source));
    sandbox.assertPlanAllowed(plan);
    return plan;
  }

  private executePlan(source: string, _program: FacadeProgram, _plan: FacadePlan): { mutationCount: number } {
    if (this.runtime.registry.hasCommand('automation.run')) {
      const result = this.runtime.execute('automation.run', { source });
      return { mutationCount: result.mutationCount };
    }
    throw new Error('Automation command is not registered; script execution is unavailable');
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
export * from './runtime';
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
