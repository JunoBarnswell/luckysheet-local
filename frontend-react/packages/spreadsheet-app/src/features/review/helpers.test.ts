import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  buildCellNote,
  buildCommentThread,
  findCommentThreadAt,
  getCellHyperlink,
  parseUrlHyperlink,
  resolveHyperlinkDisplay,
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
    sheet.commentThreads.push(buildCommentThread('sheet-1', 1, 1, 'Alice', 'Note', 'thread-1'));
    assert.equal(findCommentThreadAt(sheet, 1, 1)?.id, 'thread-1');
    assert.equal(findCommentThreadAt(sheet, 0, 0), undefined);
  });

  it('parses hyperlink targets and resolves display text', () => {
    const email = parseUrlHyperlink('mailto:team@example.com?subject=Hello', 'link-1');
    assert.equal(email.target.kind, 'email');
    const named = parseUrlHyperlink('#name:SalesTotal', 'link-2');
    assert.equal(named.target.kind, 'name');
    const display = resolveHyperlinkDisplay(named);
    assert.equal(display, '#name:SalesTotal');
  });

  it('builds visible cell notes', () => {
    const note = buildCellNote('Alice', 'Check total', 'note-1');
    assert.equal(note.visible, true);
    assert.equal(note.text, 'Check total');
  });
});
