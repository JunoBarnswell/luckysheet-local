import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { FacadeScriptRuntime, ScriptSandbox } from './index';
import { registerAutomationCommands } from './commands';
import {
  AUTOMATION_WORKER_PROTOCOL,
  consumeAutomationWorkerRequest,
  type AutomationWorkerSurface,
} from './automation-worker';

class ImmediateAutomationWorker implements AutomationWorkerSurface {
  private readonly listeners = new Map<'message' | 'error' | 'messageerror', Set<(event: { readonly data?: unknown; readonly message?: string }) => void>>([
    ['message', new Set()],
    ['error', new Set()],
    ['messageerror', new Set()],
  ]);

  postMessage(message: unknown): void {
    queueMicrotask(() => {
      const result = consumeAutomationWorkerRequest(message);
      for (const listener of this.listeners.get('message')!) listener({ data: result });
    });
  }

  terminate(): void {
    for (const listeners of this.listeners.values()) listeners.clear();
  }

  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void {
    this.listeners.get(type)!.add(listener);
  }

  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void {
    this.listeners.get(type)!.delete(listener);
  }
}

function workerFactory(): AutomationWorkerSurface {
  return new ImmediateAutomationWorker();
}

function planFor(workbook: WorkbookModel, source: string): unknown {
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const result = consumeAutomationWorkerRequest({
    protocol: AUTOMATION_WORKER_PROTOCOL,
    taskId: 'test-plan',
    kind: 'plan',
    source,
    bounds: { sheetId: sheet.id, rowCount: sheet.rowCount, columnCount: sheet.columnCount },
    limits: new ScriptSandbox().getLimits(),
    maxOperations: new ScriptSandbox().getPolicy().maxOperations,
    maxDurationMs: new ScriptSandbox().getTimeoutMs(),
  });
  if (result.status !== 'completed') throw new Error(result.status === 'failed' ? result.error.message : 'worker cancelled');
  return result.plan;
}

describe('Facade automation Worker', () => {
  it('returns a structured error for a malformed Worker request', () => {
    const result = consumeAutomationWorkerRequest({ taskId: 'bad-request' });
    if (result.status !== 'failed') throw new Error('Expected a failed Worker result');
    assert.equal(result.error.code, 'AUTOMATION_WORKER_PROTOCOL');
  });

  it('rejects JavaScript in the Worker and performs no writes', async () => {
    const workbook = new WorkbookModel('automation-dsl', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const result = await new FacadeScriptRuntime(workbook, runtime).runScriptAsync(
      "fetch('https://example.invalid');",
      new ScriptSandbox(),
      { workerFactory: workerFactory },
    );
    assert.equal(result.ok, false);
    assert.match(result.error ?? '', /blocked|DSL|worker/i);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0), undefined);
  });

  it('parses the complete program in the Worker before applying any statement', async () => {
    const workbook = new WorkbookModel('automation-atomic-parse', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const result = await new FacadeScriptRuntime(workbook, runtime).runScriptAsync(
      "sheet.getRange('A1').setValues([['must not write']]); unknown.call();",
      new ScriptSandbox(),
      { workerFactory },
    );
    assert.equal(result.ok, false);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0), undefined);
  });

  it('submits only the Worker plan to one automation transaction', () => {
    const workbook = new WorkbookModel('automation-command', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const result = runtime.execute('automation.run', {
      plan: planFor(workbook, "sheet.getRange('A1').setValues([['A']]); sheet.getRange('A1').setFontWeight('bold');"),
    });
    assert.equal(result.mutationCount, 2);
    assert.equal(runtime.getHistoryDepth().undo, 1);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0)?.style?.bold, true);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0), undefined);
  });

  it('rejects a script before writing when the Worker operation budget is exceeded', async () => {
    const workbook = new WorkbookModel('automation-limit', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry, { sandbox: new ScriptSandbox({
      ...new ScriptSandbox().getPolicy(), maxOperations: 1,
    }) });
    const result = await new FacadeScriptRuntime(workbook, runtime).runScriptAsync(
      "sheet.getRange('A1:A2').clear();",
      new ScriptSandbox({ ...new ScriptSandbox().getPolicy(), maxOperations: 1 }),
      { workerFactory },
    );
    assert.equal(result.ok, false);
    assert.match(result.error ?? '', /exceeds 1 operations/i);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.count(), 0);
  });

  it('rejects source-only and out-of-band AST command payloads', () => {
    const workbook = new WorkbookModel('automation-wire', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    assert.throws(() => runtime.execute('automation.run', { source: "sheet.getRange('A1').clear();" } as never), /PLAN_REQUIRED/i);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.count(), 0);
  });

  it('honors cancellation before a Worker transaction starts', async () => {
    const workbook = new WorkbookModel('automation-cancel', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const controller = new AbortController();
    controller.abort();
    const result = await new FacadeScriptRuntime(workbook, runtime).runScriptAsync(
      "sheet.getRange('A1').setValues([[1]]);",
      new ScriptSandbox(),
      { signal: controller.signal, workerFactory },
    );
    assert.equal(result.ok, false);
    assert.equal(result.error, 'Automation execution cancelled');
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.count(), 0);
  });
});
