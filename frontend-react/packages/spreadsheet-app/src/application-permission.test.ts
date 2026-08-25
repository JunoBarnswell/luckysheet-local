import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession permission integration', () => {
  function applyServerRole(app: WorkbookSession, role: 'owner' | 'editor' | 'commenter' | 'viewer'): void {
    app['permission'].applyServerAccess(role);
    app['permission'].setOnline(true);
  }

  it('exposes only a server-projected role and capabilities in ui snapshot', () => {
    const app = new WorkbookSession();
    applyServerRole(app, 'commenter');
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.shareRole, 'commenter');
    assert.equal(snapshot.permissions.comment, true);
    assert.equal(snapshot.permissions.editCell, false);
  });

  it('blocks viewer cell edits and surfaces a notice instead of mutating', () => {
    const app = new WorkbookSession();
    applyServerRole(app, 'viewer');
    app.dispatch({ commandId: 'sheet.cell.set', params: { sheetId: app.getActiveSheetId(), row: 0, column: 0, value: { value: 'blocked' } } });
    assert.equal(app['runtime'].model.getSheet(app.getActiveSheetId()).cells.get(0, 0)?.value, undefined);
    assert.match(app.getUiSnapshot().notice, /viewer|edit-cell|Permission/i);
  });

  it('allows commenter to add comments but not edit cells', () => {
    const app = new WorkbookSession();
    applyServerRole(app, 'commenter');
    assert.equal(app.canExecute('comment.add', { row: 0, column: 0 }), true);
    assert.equal(app.canExecute('sheet.cell.set', { row: 0, column: 0, value: { value: 1 } }), false);
    app.runCommand('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    app.addComment('Visible to reviewers');
    assert.equal(app.getUiSnapshot().selectedSheet.getCell(0, 0)?.commentText, 'Visible to reviewers');
  });

  it('protects and unprotects the current selection through sheet.protect commands', () => {
    const app = new WorkbookSession();
    app.runCommand('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }],
      primaryRangeIndex: 0,
      activeCell: { row: 1, column: 1 },
      anchorCell: { row: 1, column: 1 },
    });
    app.protectSelection({ formatCells: true });
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(sheet.protectionRules.length, 1);
    assert.equal(app.canExecute('sheet.cell.set', { row: 1, column: 1, value: { value: 9 } }), false);
    assert.equal(app.canExecute('sheet.style.set', { range: { sheetId: app.getActiveSheetId(), startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }, style: { bold: true } }), true);
    app.unprotectSelection();
    assert.equal(sheet.protectionRules.length, 0);
    assert.equal(app.canExecute('sheet.cell.set', { row: 1, column: 1, value: { value: 9 } }), true);
  });

  it('restores the exact protection rule through undo and redo', () => {
    const app = new WorkbookSession();
    app.protectSelection({ formatCells: true });
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const rule = structuredClone(sheet.protectionRules[0]);
    assert.ok(rule);
    app.undo();
    assert.equal(sheet.protectionRules.length, 0);
    app.redo();
    assert.deepEqual(sheet.protectionRules, [rule]);
    app.unprotectSelection();
    app.undo();
    assert.deepEqual(sheet.protectionRules, [rule]);
  });
});
