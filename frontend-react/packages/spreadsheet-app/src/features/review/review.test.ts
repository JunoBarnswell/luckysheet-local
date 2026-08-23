import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, getCellNote } from '@react-sheets/core-model';
import { getCellHyperlink, registerReviewFeature, serializeHyperlink } from './commands';

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
    assert.equal(getCellHyperlink(sheet, 1, 0)?.target.kind, 'url');
    assert.match(manifest.ribbon?.find((entry) => entry.id === 'review-hyperlink')?.commandId ?? '', /hyperlink\.set/);
    assert.equal(serializeHyperlink({ id: 'h1', target: { kind: 'name', name: 'SalesTotal' } }), '#name:SalesTotal');
  });

  it('restores review entities through explicit inverse mutations', () => {
    const workbook = new WorkbookModel('review-undo', 'Review undo');
    const runtime = new CommandRuntime(workbook);
    registerReviewFeature(runtime);
    const note = { id: 'note-1', author: 'Alice', text: 'Keep', createdAt: '2026-01-01', visible: false };
    const thread = {
      id: 'thread-1', sheetId: 'sheet-1', row: 0, column: 0, author: 'Bob', text: 'Thread', createdAt: '2026-01-01', replies: [],
    };

    runtime.execute('note.set', { sheetId: 'sheet-1', row: 0, column: 0, note });
    runtime.undo();
    assert.equal(getCellNote(workbook.getSheet('sheet-1'), 0, 0), undefined);

    runtime.execute('comment.add', { sheetId: 'sheet-1', row: 0, column: 0, thread });
    runtime.undo();
    assert.equal(workbook.getSheet('sheet-1').commentThreads.length, 0);

    runtime.execute('comment.add', { sheetId: 'sheet-1', row: 0, column: 0, thread });
    runtime.execute('comment.remove', { sheetId: 'sheet-1', threadId: thread.id });
    runtime.undo();
    assert.equal(workbook.getSheet('sheet-1').commentThreads[0]?.id, thread.id);

    runtime.execute('hyperlink.set', { sheetId: 'sheet-1', row: 1, column: 0, hyperlink: { id: 'h1', target: { kind: 'url', url: 'https://example.com' } } });
    runtime.undo();
    assert.equal(getCellHyperlink(workbook.getSheet('sheet-1'), 1, 0), undefined);
  });

  it('does not create history entries for missing review targets', () => {
    const workbook = new WorkbookModel('review-no-target', 'Review no target');
    const runtime = new CommandRuntime(workbook);
    registerReviewFeature(runtime);

    const result = runtime.execute('comment.remove', { sheetId: 'sheet-1', threadId: 'missing-thread' });
    assert.equal(result.mutationCount, 0);
    assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });

    const hyperlinkResult = runtime.execute('hyperlink.remove', { sheetId: 'sheet-1', row: 1, column: 1 });
    assert.equal(hyperlinkResult.mutationCount, 0);
    assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
  });

  it('replays a comment resolution with the operation timestamp', () => {
    const workbook = new WorkbookModel('review-resolution', 'Review resolution');
    const runtime = new CommandRuntime(workbook);
    registerReviewFeature(runtime);
    runtime.execute('comment.add', {
      sheetId: 'sheet-1', row: 0, column: 0,
      thread: { id: 'thread-1', sheetId: 'sheet-1', row: 0, column: 0, author: 'Alice', text: 'Review', createdAt: '2026-01-01', replies: [] },
    });
    runtime.execute('comment.resolve', {
      sheetId: 'sheet-1', threadId: 'thread-1', resolved: true, resolvedAt: '2026-02-03T04:05:06.000Z',
    });
    assert.equal(workbook.getSheet('sheet-1').commentThreads[0]?.resolvedAt, '2026-02-03T04:05:06.000Z');
    assert.throws(() => runtime.execute('comment.resolve', {
      sheetId: 'sheet-1', threadId: 'thread-1', resolved: true,
    }), /resolvedAt/);
  });
});
