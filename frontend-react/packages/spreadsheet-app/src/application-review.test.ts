import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { getCellNote } from '@react-sheets/core-model';
import { WorkbookSession } from './workbook-session';
import { findCommentThreadAt, getCellHyperlink } from './features/review';

function selectCell(app: WorkbookSession, row: number, column: number): void {
  const sheetId = app.getActiveSheetId();
  app.runCommand('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }],
    primaryRangeIndex: 0,
    activeCell: { row, column },
    anchorCell: { row, column },
  });
}

describe('WorkbookSession review integration', () => {
  it('addComment routes through comment.add and appears in ui snapshot', () => {
    const app = new WorkbookSession();
    selectCell(app, 1, 2);
    app.addComment('Please review totals');

    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const thread = findCommentThreadAt(sheet, 1, 2);
    assert.ok(thread);
    assert.equal(thread.text, 'Please review totals');

    const cell = app.getUiSnapshot().selectedSheet.getCell(1, 2);
    assert.equal(cell?.commentText, 'Please review totals');
    assert.equal(cell?.hasComment, true);
  });

  it('replyComment and resolveComment use comment.reply and comment.resolve', () => {
    const app = new WorkbookSession();
    selectCell(app, 0, 0);
    app.addComment('Initial thread');
    app.replyComment('Follow-up');
    app.resolveComment();

    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const thread = findCommentThreadAt(sheet, 0, 0);
    assert.equal(thread?.replies.length, 1);
    assert.equal(thread?.resolved, true);

    const cell = app.getUiSnapshot().selectedSheet.getCell(0, 0);
    assert.equal(cell?.comment?.replies?.length, 1);
    assert.equal(cell?.comment?.resolved, true);
  });

  it('removeComment deletes the thread through comment.remove', () => {
    const app = new WorkbookSession();
    selectCell(app, 3, 1);
    app.addComment('Temporary');
    app.removeComment();

    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(findCommentThreadAt(sheet, 3, 1), undefined);
    assert.equal(app.getUiSnapshot().selectedSheet.getCell(3, 1)?.hasComment, false);
  });

  it('setActiveHyperlink and removeHyperlink use typed hyperlink commands with detail payload', () => {
    const app = new WorkbookSession();
    selectCell(app, 2, 2);
    app.setActiveHyperlink({ kind: 'url', url: 'https://example.com/docs' }, 'Documentation');
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(getCellHyperlink(sheet, 2, 2)?.target.kind, 'url');
    assert.equal(getCellHyperlink(sheet, 2, 2)?.tooltip, 'Documentation');
    assert.equal(app.getUiSnapshot().selectedSheet.getCell(2, 2)?.hyperlink, 'https://example.com/docs');

    app.removeHyperlink();
    assert.equal(getCellHyperlink(sheet, 2, 2), undefined);
  });

  it('supports typed email, worksheet, and defined-name targets with atomic undo/redo', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    selectCell(app, 0, 0);
    app['runtime'].model.setDefinedName({ name: 'SalesTotal', formula: '=Sheet1!A1', scope: 'workbook' });
    app.setActiveHyperlink({ kind: 'email', address: 'team@example.com', subject: 'Review' }, 'Email tip');
    assert.equal(getCellHyperlink(app['runtime'].model.getSheet(sheetId), 0, 0)?.target.kind, 'email');
    app.undo();
    assert.equal(getCellHyperlink(app['runtime'].model.getSheet(sheetId), 0, 0), undefined);
    app.redo();
    assert.equal(getCellHyperlink(app['runtime'].model.getSheet(sheetId), 0, 0)?.tooltip, 'Email tip');
    app.setActiveHyperlink({ kind: 'sheet', sheetId, address: 'B2' });
    assert.equal(getCellHyperlink(app['runtime'].model.getSheet(sheetId), 0, 0)?.target.kind, 'sheet');
    app.setActiveHyperlink({ kind: 'name', name: 'SalesTotal' });
    assert.equal(getCellHyperlink(app['runtime'].model.getSheet(sheetId), 0, 0)?.target.kind, 'name');
  });

  it('rejects invalid typed targets before creating a mutation', () => {
    const app = new WorkbookSession();
    selectCell(app, 0, 0);
    assert.throws(() => app.setActiveHyperlink({ kind: 'url', url: 'not-a-url' }), /Invalid hyperlink URL/);
    assert.throws(() => app.setActiveHyperlink({ kind: 'sheet', sheetId: 'missing', address: 'A1' }), /target sheet not found/);
    assert.equal(getCellHyperlink(app['runtime'].model.getSheet(app.getActiveSheetId()), 0, 0), undefined);
  });

  it('addNote and removeNote route through note.set and note.remove', () => {
    const app = new WorkbookSession();
    selectCell(app, 4, 0);
    app.addNote('Audit note');
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    assert.equal(getCellNote(sheet, 4, 0)?.text, 'Audit note');
    assert.equal(app.getUiSnapshot().selectedSheet.getCell(4, 0)?.hasComment, true);
    assert.equal(app.getUiSnapshot().selectedSheet.getCell(4, 0)?.note?.text, 'Audit note');

    app.removeNote();
    assert.equal(getCellNote(sheet, 4, 0), undefined);
  });
});
