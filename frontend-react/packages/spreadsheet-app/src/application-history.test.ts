import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';
import { buildRestoreParams } from './features/history';

describe('WorkbookSession history integration', () => {
  it('does not accept a client-provided snapshot as a restore mutation', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'before' },
    });
    const snapshot = app['runtime'].model.snapshot();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'after' },
    });
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0)?.value, 'after');

    app.restoreFromSnapshot(snapshot, 1, 'test restore');
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0)?.value, 'after');
    assert.match(app.getUiSnapshot().notice, /server-authorized|restore/i);
    assert.equal(app.getUiSnapshot().historyPreviewRevision, null);
  });

  it('undoes session history to a selected entry index', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'step-1' } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 'step-2' } });
    app.undoToHistoryIndex(0);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0)?.value, 'step-1');
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.value, undefined);
  });

  it('requests authoritative recovery when undo observers fail after a live replay', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    let resynchronizations = 0;
    runtime.collab = { requestResynchronization: () => { resynchronizations += 1; } } as never;
    app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(), row: 0, column: 0, value: { value: 'undo target' },
    });
    runtime.commands.onMutation((_mutation, source) => {
      if (source === 'undo') throw new Error('projection observer failed');
    });

    app.undo();

    assert.equal(resynchronizations, 1);
    assert.equal(runtime.commands.isMutationRecoveryRequired, true);
    assert.match(app.getUiSnapshot().notice, /重新同步工作簿/);
    assert.equal(app.canExecute('sheet.cell.set'), false);
  });

  it('requests recovery once when redo fails and later commands see the same lock error', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    let resynchronizations = 0;
    runtime.collab = { requestResynchronization: () => { resynchronizations += 1; } } as never;
    app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(), row: 0, column: 0, value: { value: 'redo target' },
    });
    app.undo();
    runtime.commands.onMutation((_mutation, source) => {
      if (source === 'redo') throw new Error('projection observer failed');
    });

    app.redo();

    assert.equal(resynchronizations, 1);
    assert.equal(runtime.commands.isMutationRecoveryRequired, true);
    assert.match(app.getUiSnapshot().notice, /重新同步工作簿/);
    assert.throws(() => app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(), row: 0, column: 1, value: { value: 'retry while locked' },
    }), /MUTATION_RECOVERY_REQUIRED/);
    assert.equal(resynchronizations, 1);
  });

  it('does not report a successful history-index restore after replay recovery is required', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    let resynchronizations = 0;
    runtime.collab = { requestResynchronization: () => { resynchronizations += 1; } } as never;
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'first' } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 'second' } });
    runtime.commands.onMutation((_mutation, source) => {
      if (source === 'undo') throw new Error('projection observer failed');
    });

    app.undoToHistoryIndex(0);

    assert.equal(resynchronizations, 1);
    assert.equal(runtime.commands.isMutationRecoveryRequired, true);
    assert.match(app.getUiSnapshot().notice, /重新同步工作簿/);
    assert.doesNotMatch(app.getUiSnapshot().notice, /Restored session history/);
  });

  it('requests authoritative recovery when command rollback observers fail', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    let resynchronizations = 0;
    runtime.collab = { requestResynchronization: () => { resynchronizations += 1; } } as never;
    runtime.commands.onMutation((_mutation, source) => {
      if (source === 'command' || source === 'undo') throw new Error('projection observer failed');
    });

    assert.throws(() => app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(), row: 0, column: 0, value: { value: 'command target' },
    }), /MUTATION_RECOVERY_REQUIRED/);

    assert.equal(resynchronizations, 1);
    assert.equal(runtime.commands.isMutationRecoveryRequired, true);
    assert.match(app.getUiSnapshot().notice, /重新同步工作簿/);
  });

  it('blocks history.restore for viewers', () => {
    const app = new WorkbookSession();
    app['permission'].applyServerAccess({ unitId: 'wb-test', role: 'viewer', accessRevision: 1, regions: [] });
    app['permission'].setOnline(true);
    const snapshot = app['runtime'].model.snapshot();
    app.restoreFromSnapshot(snapshot, 0, 'blocked');
    assert.match(app.getUiSnapshot().notice, /viewer|restore|Permission/i);
  });

  it('builds a target-revision restore request without a snapshot payload', () => {
    const params = buildRestoreParams(3, 'reason');
    assert.equal(params.targetRevision, 3);
    assert.equal('snapshot' in params, false);
    assert.equal(params.reason, 'reason');
  });
});
