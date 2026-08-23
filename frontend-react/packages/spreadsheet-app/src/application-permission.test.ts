import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

describe('SpreadsheetApplication permission integration', () => {
  it('exposes share role and capabilities in ui snapshot', () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('commenter');
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.shareRole, 'commenter');
    assert.equal(snapshot.permissions.comment, true);
    assert.equal(snapshot.permissions.editCell, false);
  });

  it('blocks viewer cell edits and surfaces a notice instead of mutating', () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('viewer');
    app.execute('sheet.cell.set', { sheetId: app.getActiveSheetId(), row: 0, column: 0, value: { value: 'blocked' } });
    assert.equal(app.getWorkbook().getSheet(app.getActiveSheetId()).cells.get(0, 0)?.value, undefined);
    assert.match(app.getUiSnapshot().notice, /viewer|edit-cell|Permission/i);
  });

  it('allows commenter to add comments but not edit cells', () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('commenter');
    assert.equal(app.canExecute('comment.add', { row: 0, column: 0 }), true);
    assert.equal(app.canExecute('sheet.cell.set', { row: 0, column: 0, value: { value: 1 } }), false);
    app.execute('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      primaryRangeIndex: 0,
      primaryRowIndex: 0,
      primaryColumnIndex: 0,
    });
    app.addComment('Visible to reviewers');
    assert.equal(app.getUiSnapshot().selectedSheet.getCell(0, 0)?.commentText, 'Visible to reviewers');
  });

  it('protects and unprotects the current selection through sheet.protect commands', () => {
    const app = new SpreadsheetApplication();
    app.execute('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }],
      primaryRangeIndex: 0,
      primaryRowIndex: 1,
      primaryColumnIndex: 1,
    });
    app.protectSelection(['format']);
    const sheet = app.getWorkbook().getSheet(app.getActiveSheetId());
    assert.equal(sheet.protectionRules.length, 1);
    assert.equal(app.canExecute('sheet.cell.set', { row: 1, column: 1, value: { value: 9 } }), false);
    assert.equal(app.canExecute('sheet.style.set', { range: { sheetId: app.getActiveSheetId(), startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }, style: { bold: true } }), true);
    app.unprotectSelection();
    assert.equal(sheet.protectionRules.length, 0);
    assert.equal(app.canExecute('sheet.cell.set', { row: 1, column: 1, value: { value: 9 } }), true);
  });
});
