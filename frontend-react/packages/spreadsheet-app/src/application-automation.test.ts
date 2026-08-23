import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';
import { SAMPLE_AUTOMATION_SCRIPT } from './features/automation';
import { consumeAutomationWorkerRequest, type AutomationWorkerSurface } from './features/automation/automation-worker';

class ImmediateAutomationWorker implements AutomationWorkerSurface {
  private readonly listeners = new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>();

  postMessage(message: unknown): void {
    queueMicrotask(() => {
      const result = consumeAutomationWorkerRequest(message);
      for (const listener of this.listeners) listener({ data: result });
    });
  }

  terminate(): void { this.listeners.clear(); }
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void {
    if (type === 'message') this.listeners.delete(listener);
  }
}

describe('WorkbookSession automation integration', () => {
  it('runs scripts through automation.run command path', async () => {
    const app = new WorkbookSession({ automationWorkerFactory: () => new ImmediateAutomationWorker() });
    await app.runAutomationScript(SAMPLE_AUTOMATION_SCRIPT);
    assert.equal(app.getUiSnapshot().lastScriptResult?.ok, true, app.getUiSnapshot().lastScriptResult?.error ?? 'Automation worker did not complete');
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(sheet.cells.get(0, 0)?.value, 'Automated');
    assert.equal(app.getUiSnapshot().activePanel, 'automate');
  });

  it('records commands into facade script text', () => {
    const app = new WorkbookSession();
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

  it('blocks automation.run for viewers', async () => {
    const app = new WorkbookSession();
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);
    await app.runAutomationScript(SAMPLE_AUTOMATION_SCRIPT);
    assert.equal(app.getUiSnapshot().lastScriptResult, null);
    assert.match(app.getUiSnapshot().notice, /permission|viewer|script/i);
  });
});
