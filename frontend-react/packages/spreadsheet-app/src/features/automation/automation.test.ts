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
    assert.equal(workbook.getSheet(workbook.activeSheetId).cells.get(0, 0), undefined);
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
    assert.equal(workbook.getSheet(workbook.activeSheetId).cells.get(0, 0), undefined);
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
    assert.equal(workbook.getSheet(workbook.activeSheetId).cells.get(0, 0)?.style?.bold, true);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet(workbook.activeSheetId).cells.get(0, 0), undefined);
  });
});
