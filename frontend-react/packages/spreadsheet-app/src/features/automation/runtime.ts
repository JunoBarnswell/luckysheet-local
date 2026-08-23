import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { WorkbookModel } from '@react-sheets/core-model';
import { CommandRecorder } from './command-recorder';
import { registerAutomationCommands } from './commands';
import { FacadeScriptRuntime, type ScriptRunResult } from './index';
import { ScriptSandbox } from './sandbox';
import type { AutomationWorkerFactory } from './automation-worker';

export interface AutomationRunParams {
  source: string;
  label?: string;
}

export interface AutomationSnapshot {
  recording: boolean;
  recordedScript: string;
  lastResult: ScriptRunResult | null;
  lastRunAt: string | null;
}

export interface AutomationAsyncOptions {
  signal?: AbortSignal;
  workerFactory?: AutomationWorkerFactory;
}

export const SAMPLE_AUTOMATION_SCRIPT = `sheet.getRange('A1').setValues([['Automated']]);
sheet.getRange('A1').setFontWeight('bold');`;

export function createCommandRecorder(): CommandRecorder {
  return new CommandRecorder();
}

export function runAutomationScript(
  workbook: WorkbookModel,
  commands: CommandRuntime,
  source: string,
  sandbox = new ScriptSandbox(),
): ScriptRunResult {
  if (!commands.registry.hasCommand('automation.run')) registerAutomationCommands(commands.registry, { sandbox });
  const runtime = new FacadeScriptRuntime(workbook, commands);
  return runtime.runScript(source.trim(), sandbox);
}

/** Browser Worker path. The synchronous API above deliberately does not fall back. */
export async function runAutomationScriptAsync(
  workbook: WorkbookModel,
  commands: CommandRuntime,
  source: string,
  sandbox = new ScriptSandbox(),
  options: AutomationAsyncOptions = {},
): Promise<ScriptRunResult> {
  if (!commands.registry.hasCommand('automation.run')) registerAutomationCommands(commands.registry, { sandbox });
  const runtime = new FacadeScriptRuntime(workbook, commands);
  return runtime.runScriptAsync(source.trim(), sandbox, options);
}

export function summarizeScriptResult(result: ScriptRunResult): string {
  if (result.ok) return `Script completed in ${result.durationMs}ms`;
  return `Script failed: ${result.error ?? 'unknown error'}`;
}
