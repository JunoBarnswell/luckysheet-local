import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';
import { SAMPLE_AUTOMATION_SCRIPT } from './automation-bridge';

describe('SpreadsheetApplication automation integration', () => {
  it('runs scripts through automation.run command path', () => {
    const app = new SpreadsheetApplication();
    app.runAutomationScript(SAMPLE_AUTOMATION_SCRIPT);
    const sheet = app.getWorkbook().getSheet(app.getActiveSheetId());
    assert.equal(sheet.cells.get(0, 0)?.value, 'Automated');
    assert.equal(app.getUiSnapshot().lastScriptResult?.ok, true);
    assert.equal(app.getUiSnapshot().activePanel, 'automate');
  });

  it('records commands into facade script text', () => {
    const app = new SpreadsheetApplication();
    app.startAutomationRecording();
    app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(),
      row: 1,
      column: 1,
      value: { value: 'Recorded' },
    });
    const script = app.stopAutomationRecording();
    assert.match(script, /getRange\('B2'\)/);
    assert.equal(app.getUiSnapshot().automationRecording, false);
    assert.match(app.getUiSnapshot().recordedScript, /setValues/);
  });

  it('blocks automation.run for viewers', () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('viewer');
    app.runAutomationScript(SAMPLE_AUTOMATION_SCRIPT);
    assert.equal(app.getUiSnapshot().lastScriptResult, null);
    assert.match(app.getUiSnapshot().notice, /permission|viewer|script/i);
  });
});
