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

  it('blocks history.restore for viewers', () => {
    const app = new WorkbookSession();
    app['permission'].applyServerAccess('viewer');
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
