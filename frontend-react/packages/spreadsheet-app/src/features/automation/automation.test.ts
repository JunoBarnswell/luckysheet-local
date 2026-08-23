import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { FacadeScriptRuntime, ScriptSandbox } from './index';
import { registerAutomationCommands } from './commands';

describe('Facade automation DSL', () => {
  it('rejects JavaScript and performs no writes', () => {
    const workbook = new WorkbookModel('automation-dsl', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const result = new FacadeScriptRuntime(workbook, runtime).runScript(
      "fetch('https://example.invalid');",
      new ScriptSandbox(),
    );
    assert.equal(result.ok, false);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0), undefined);
  });

  it('parses the complete program before applying any statement', () => {
    const workbook = new WorkbookModel('automation-atomic-parse', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const result = new FacadeScriptRuntime(workbook, runtime).runScript(
      "sheet.getRange('A1').setValues([['must not write']]); unknown.call();",
      new ScriptSandbox(),
    );
    assert.equal(result.ok, false);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0), undefined);
  });

  it('executes a validated program as one CommandRuntime transaction', () => {
    const workbook = new WorkbookModel('automation-command', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const result = runtime.execute('automation.run', {
      source: "sheet.getRange('A1').setValues([['A']]); sheet.getRange('A1').setFontWeight('bold');",
    });
    assert.equal(result.mutationCount, 2);
    assert.equal(runtime.getHistoryDepth().undo, 1);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0)?.style?.bold, true);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.get(0, 0), undefined);
  });

  it('returns a serializable plan with resource limits', () => {
    const workbook = new WorkbookModel('automation-plan', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const result = runtime.execute('automation.run', { source: "sheet.getRange('A1').setValues([[1]]);" }) as { plan?: { schema?: string; sourceHash?: string; serializable?: boolean; limits?: { maxOperations: number } } };
    assert.equal(result.plan?.schema, 'AutomationPlan');
    assert.equal(result.plan?.serializable, true);
    assert.ok(result.plan?.sourceHash);
    assert.ok((result.plan?.limits?.maxOperations ?? 0) > 0);
  });

  it('rejects a script before writing when the operation budget is exceeded', () => {
    const workbook = new WorkbookModel('automation-limit', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry, { sandbox: new ScriptSandbox({
      ...new ScriptSandbox().getPolicy(), maxOperations: 1,
    }) });
    assert.throws(() => runtime.execute('automation.run', { source: "sheet.getRange('A1:A2').clear();" }), /exceeds 1 operations/i);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.count(), 0);
  });

  it('rejects an out-of-band AST payload explicitly', () => {
    const workbook = new WorkbookModel('automation-wire', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    assert.throws(() => runtime.execute('automation.run', { source: "sheet.getRange('A1').clear();", program: () => undefined } as never), /source only|not serializable/i);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.count(), 0);
  });

  it('honors cancellation and timeout before opening a transaction', () => {
    const workbook = new WorkbookModel('automation-cancel', 'DSL');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerAutomationCommands(runtime.registry);
    const controller = new AbortController();
    controller.abort();
    const cancelled = new FacadeScriptRuntime(workbook, runtime).runScript(
      "sheet.getRange('A1').setValues([[1]]);",
      new ScriptSandbox(),
      { signal: controller.signal },
    );
    assert.equal(cancelled.ok, false);
    assert.match(cancelled.error ?? '', /cancel/i);
    const timedOut = runtime.execute.bind(runtime, 'automation.run', { source: "sheet.getRange('A1').setValues([[1]]);", deadlineAt: Date.now() - 1 });
    assert.throws(timedOut, /timed out|execution/i);
    assert.equal(workbook.getSheet(workbook.primarySheetId).cells.count(), 0);
  });
});
