import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  buildCellNote,
  buildCommentThread,
  findCommentThreadAt,
  getCellHyperlink,
  resolveHyperlinkDisplay,
  validateHyperlinkTarget,
  threadToCellComment,
} from './helpers';

describe('review helpers', () => {
  it('builds comment threads and maps them to cell comment snapshots', () => {
    const thread = buildCommentThread('sheet-1', 2, 3, 'Alice', 'Please review @bob', 'thread-1');
    assert.equal(thread.mentions?.[0], 'bob');
    const mapped = threadToCellComment(thread);
    assert.equal(mapped.text, 'Please review @bob');
    assert.equal(mapped.replies?.length ?? 0, 0);
  });

  it('finds comment threads by cell coordinates', () => {
    const workbook = new WorkbookModel('wb', 'Review');
    const sheet = workbook.getSheet('sheet-1');
    sheet.review.addThread(buildCommentThread('sheet-1', 1, 1, 'Alice', 'Note', 'thread-1'));
    assert.equal(findCommentThreadAt(sheet, 1, 1)?.id, 'thread-1');
    assert.equal(findCommentThreadAt(sheet, 0, 0), undefined);
  });

  it('validates typed hyperlink targets and resolves serialization only for display/interchange', () => {
    const workbook = new WorkbookModel('wb', 'Review');
    workbook.setDefinedName({ name: 'SalesTotal', formula: '=Sheet1!A1', scope: 'workbook' });
    const email = { kind: 'email' as const, address: 'team@example.com', subject: 'Hello' };
    validateHyperlinkTarget(email, workbook, 'sheet-1');
    const named = { kind: 'name' as const, name: 'SalesTotal' };
    validateHyperlinkTarget(named, workbook, 'sheet-1');
    const namedLink = { id: 'link-2', target: named };
    assert.equal(namedLink.target.kind, 'name');
    const display = resolveHyperlinkDisplay(namedLink);
    assert.equal(display, '#name:SalesTotal');
    assert.throws(() => validateHyperlinkTarget({ kind: 'url', url: 'not-a-url' }, workbook, 'sheet-1'), /Invalid hyperlink URL/);
    assert.throws(() => validateHyperlinkTarget({ kind: 'sheet', sheetId: 'missing', address: 'A1' }, workbook, 'sheet-1'), /sheet not found/);
  });

  it('builds visible cell notes', () => {
    const note = buildCellNote('Alice', 'Check total', 'note-1');
    assert.equal(note.visible, true);
    assert.equal(note.text, 'Check total');
  });
});
