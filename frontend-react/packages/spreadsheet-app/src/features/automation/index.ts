import type { WorkbookModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import { ScriptSandbox } from './sandbox';
import {
  validateFacadePlan,
  type FacadePlan,
  type FacadeSheetBounds,
} from './dsl';
import {
  AutomationWorkerClient,
  createAutomationWorker,
  type AutomationWorkerFactory,
} from './automation-worker';

/** Facade 脚本运行时 — 脚本只允许调 Facade */
export class FacadeScriptRuntime {
  constructor(
    private readonly workbook: WorkbookModel,
    private readonly runtime: CommandRuntime,
  ) {}

  /**
   * Synchronous execution is intentionally unavailable. A browser Worker
   * cannot be synchronously joined from the main thread, and this method must
   * not become an accidental main-thread fallback.
   */
  runScript(_source: string, _sandbox: ScriptSandbox): ScriptRunResult {
    return {
      ok: false,
      durationMs: 0,
      error: 'AUTOMATION_WORKER_ASYNC_REQUIRED: use runScriptAsync to execute in a browser Worker',
    };
  }

  async runScriptAsync(source: string, sandbox: ScriptSandbox, options: ScriptRunOptions = {}): Promise<ScriptRunResult> {
    const started = Date.now();
    const sheet = this.workbook.getSheet(this.workbook.primarySheetId);
    const bounds: FacadeSheetBounds = {
      sheetId: sheet.id,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
    };
    const worker = new AutomationWorkerClient(options.workerFactory ?? createAutomationWorker);
    try {
      const workerResult = await worker.submit(source, bounds, {
        limits: sandbox.getLimits(),
        maxOperations: sandbox.getPolicy().maxOperations,
        maxDurationMs: sandbox.getTimeoutMs(),
      }, options.signal);
      if (workerResult.status === 'cancelled') {
        return { ok: false, durationMs: Date.now() - started, error: 'Automation execution cancelled' };
      }
      if (workerResult.status === 'failed') {
        return { ok: false, durationMs: Date.now() - started, error: `${workerResult.error.code}: ${workerResult.error.message}` };
      }
      validateFacadePlan(workerResult.plan, bounds);
      sandbox.assertPlanAllowed(workerResult.plan);
      const result = this.runtime.execute('automation.run', { plan: workerResult.plan });
      return {
        ok: true,
        durationMs: Date.now() - started,
        mutationCount: result.mutationCount,
        plan: workerResult.plan,
      };
    } catch (error) {
      return { ok: false, durationMs: Date.now() - started, error: error instanceof Error ? error.message : String(error) };
    } finally {
      worker.dispose();
    }
  }

  /** Public deterministic plan API for hosts that own transaction execution. */
  planScript(source: string, sandbox = new ScriptSandbox()): FacadePlan {
    void source;
    void sandbox;
    throw new Error('AUTOMATION_WORKER_ASYNC_REQUIRED: use planScriptAsync to plan in a browser Worker');
  }

  async planScriptAsync(source: string, sandbox = new ScriptSandbox(), options: ScriptRunOptions = {}): Promise<FacadePlan> {
    const sheet = this.workbook.getSheet(this.workbook.primarySheetId);
    const worker = new AutomationWorkerClient(options.workerFactory ?? createAutomationWorker);
    try {
      const result = await worker.submit(source, { sheetId: sheet.id, rowCount: sheet.rowCount, columnCount: sheet.columnCount }, {
        limits: sandbox.getLimits(),
        maxOperations: sandbox.getPolicy().maxOperations,
        maxDurationMs: sandbox.getTimeoutMs(),
      }, options.signal);
      if (result.status === 'cancelled') throw new Error('Automation execution cancelled');
      if (result.status === 'failed') throw new Error(`${result.error.code}: ${result.error.message}`);
      validateFacadePlan(result.plan, { sheetId: sheet.id, rowCount: sheet.rowCount, columnCount: sheet.columnCount });
      sandbox.assertPlanAllowed(result.plan);
      return result.plan;
    } finally {
      worker.dispose();
    }
  }
}

export interface ScriptRunOptions {
  signal?: AbortSignal;
  workerFactory?: AutomationWorkerFactory;
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
  AutomationWorkerClient,
  createAutomationWorker,
  consumeAutomationWorkerRequest,
  isAutomationWorkerCancel,
  isAutomationWorkerRequest,
  isAutomationWorkerResult,
  AUTOMATION_WORKER_PROTOCOL,
  type AutomationWorkerCancel,
  type AutomationWorkerFactory,
  type AutomationWorkerFailure,
  type AutomationWorkerRequest,
  type AutomationWorkerResult,
  type AutomationWorkerSurface,
} from './automation-worker';
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
  type FacadeExecutionControl,
  type FacadePlan,
  type FacadeProgram,
  type FacadeStatement,
} from './dsl';
