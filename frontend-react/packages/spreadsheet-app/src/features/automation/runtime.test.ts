import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import {
  runAutomationScript,
  SAMPLE_AUTOMATION_SCRIPT,
  summarizeScriptResult,
} from './runtime';
import { CommandRecorder } from './command-recorder';
import { ScriptSandbox } from './sandbox';

describe('automation runtime', () => {
  it('runs facade scripts through command runtime', () => {
    const model = new WorkbookModel('wb-auto', 'Automation');
    const runtime = new CommandRuntime(model);
    registerSheetCommands(runtime);
    const sheetId = model.activeSheetId;

    const result = runAutomationScript(model, runtime, SAMPLE_AUTOMATION_SCRIPT);
    assert.equal(result.ok, true);
    assert.match(summarizeScriptResult(result), /completed/i);

    const sheet = model.getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'Automated');
    assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);
  });

  it('blocks scripts that violate sandbox policy', () => {
    const model = new WorkbookModel('wb-blocked', 'Blocked');
    const runtime = new CommandRuntime(model);
    registerSheetCommands(runtime);
    const result = runAutomationScript(
      model,
      runtime,
      "fetch('https://example.com');",
      new ScriptSandbox(),
    );
    assert.equal(result.ok, false);
    assert.match(result.error ?? '', /blocked/i);
  });
});

describe('CommandRecorder', () => {
  it('records facade statements from command events', () => {
    const recorder = new CommandRecorder();
    const listener = recorder.createListener();
    recorder.start();
    listener('sheet.cell.set', { row: 0, column: 0, value: { value: 'Hi' } }, {
      operationId: 'op-1',
      mutationCount: 1,
      affectedRanges: [],
    });
    listener('sheet.style.set', {
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      style: { bold: true },
    }, {
      operationId: 'op-2',
      mutationCount: 1,
      affectedRanges: [],
    });
    const statements = recorder.stop();
    assert.equal(statements.length, 2);
    assert.match(recorder.toScript(), /getRange\('A1'\)/);
    assert.match(recorder.toScript(), /setFontWeight\('bold'\)/);
  });

  it('fails explicitly when a command has no serializable Facade form', () => {
    const recorder = new CommandRecorder();
    const listener = recorder.createListener();
    recorder.start();
    assert.throws(() => listener('pivot.refresh', {}, {
      operationId: 'op-3', mutationCount: 1, affectedRanges: [],
    }), /Cannot serialize recorded command/i);
  });
});
