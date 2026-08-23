import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, getCellNote } from '@react-sheets/core-model';
import { registerReviewFeature, serializeHyperlink } from './commands';

describe('review feature', () => {
  it('keeps notes separate from comment threads and supports hyperlinks', () => {
    const workbook = new WorkbookModel('review-test', 'Review');
    const runtime = new CommandRuntime(workbook);
    const manifest = registerReviewFeature(runtime);

    runtime.execute('note.set', {
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
      note: { id: 'note-1', author: 'Alice', text: 'Check total', createdAt: '2026-01-01', visible: true },
    });
    runtime.execute('comment.add', {
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
      thread: {
        id: 'thread-1',
        sheetId: 'sheet-1',
        row: 0,
        column: 0,
        author: 'Bob',
        text: 'Please review',
        createdAt: '2026-01-02',
        replies: [],
      },
    });
    runtime.execute('hyperlink.set', {
      sheetId: 'sheet-1',
      row: 1,
      column: 0,
      hyperlink: { id: 'h1', target: { kind: 'url', url: 'https://example.com' } },
    });

    const sheet = workbook.getSheet('sheet-1');
    assert.equal(getCellNote(sheet, 0, 0)?.text, 'Check total');
    assert.equal(sheet.commentThreads.length, 1);
    assert.equal(sheet.cells.get(1, 0)?.hyperlinkDetail?.target.kind, 'url');
    assert.match(manifest.ribbon?.find((entry) => entry.id === 'review-hyperlink')?.commandId ?? '', /hyperlink\.set/);
    assert.equal(serializeHyperlink({ id: 'h1', target: { kind: 'name', name: 'SalesTotal' } }), '#name:SalesTotal');
  });
});
