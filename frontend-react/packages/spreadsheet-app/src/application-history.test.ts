import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';
import { buildRestoreParams } from './features/history';

describe('SpreadsheetApplication history integration', () => {
  it('restores workbook state through history.restore and rebuilds formulas', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'before' },
    });
    const snapshot = app.getWorkbook().snapshot();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'after' },
    });
    assert.equal(app.getWorkbook().getSheet(sheetId).cells.get(0, 0)?.value, 'after');

    app.restoreFromSnapshot(snapshot, 1, 'test restore');
    assert.equal(app.getWorkbook().getSheet(sheetId).cells.get(0, 0)?.value, 'before');
    assert.equal(app.getUiSnapshot().historyPreviewRevision, null);
  });

  it('undoes session history to a selected entry index', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'step-1' } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 'step-2' } });
    app.undoToHistoryIndex(0);
    assert.equal(app.getWorkbook().getSheet(sheetId).cells.get(0, 0)?.value, 'step-1');
    assert.equal(app.getWorkbook().getSheet(sheetId).cells.get(0, 1)?.value, undefined);
  });

  it('blocks history.restore for viewers', () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('viewer');
    const snapshot = app.getWorkbook().snapshot();
    app.restoreFromSnapshot(snapshot, 0, 'blocked');
    assert.match(app.getUiSnapshot().notice, /viewer|restore|Permission/i);
  });

  it('builds restore params with cloned snapshot payload', () => {
    const snapshot = {
      schema: 'WorkbookSnapshotV1' as const,
      unitId: 'wb-1',
      name: 'Clone',
      activeSheetId: 'sheet-1',
      definedNames: {},
      tables: [],
      sheets: [],
    };
    const params = buildRestoreParams(snapshot, 3, 'reason');
    assert.equal(params.targetRevision, 3);
    assert.notEqual(params.snapshot, snapshot);
    assert.equal(params.reason, 'reason');
  });
});
