import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { WorkbookModel } from '@react-sheets/core-model';
import { CommandRecorder } from './command-recorder';
import { FacadeScriptRuntime, type ScriptRunResult } from './index';
import { ScriptSandbox } from './sandbox';

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
  const runtime = new FacadeScriptRuntime(workbook, commands);
  return runtime.runScript(source.trim(), sandbox);
}

export function summarizeScriptResult(result: ScriptRunResult): string {
  if (result.ok) return `Script completed in ${result.durationMs}ms`;
  return `Script failed: ${result.error ?? 'unknown error'}`;
}
