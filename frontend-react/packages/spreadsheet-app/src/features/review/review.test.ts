import assert from 'node:assert/strict';
import test from 'node:test';
import { registerReviewFeature, serializeHyperlink } from './commands';
import { getCellHyperlink } from './helpers';
import { openCanonicalTestRuntime } from '../../../../core-model/src/canonical-test-runtime.test';

async function setup(unitId: string) {
  const fixture = await openCanonicalTestRuntime(unitId, 'Review');
  registerReviewFeature(fixture.runtime);
  return fixture;
}

test('review commands publish typed note, comment and hyperlink mutations', async () => {
  const { workbook, runtime, close } = await setup('review-command-success');
  try {
    const planned: string[] = [];
    runtime.onMutation((mutation) => planned.push(mutation.id));
    await runtime.execute('note.set', {
      sheetId: 'sheet-1', row: 0, column: 0,
      note: { id: 'note-1', author: 'Alice', text: 'Check total', createdAt: '2026-01-01', visible: true },
    });
    await runtime.execute('comment.add', {
      sheetId: 'sheet-1', row: 0, column: 0,
      thread: { id: 'thread-1', sheetId: 'sheet-1', row: 0, column: 0, author: 'Bob', text: 'Please review', createdAt: '2026-01-02', replies: [] },
    });
    await runtime.execute('hyperlink.set', {
      sheetId: 'sheet-1', row: 1, column: 0,
      hyperlink: { id: 'h1', target: { kind: 'url', url: 'https://example.com' } },
    });

    assert.deepEqual(planned, ['note.set', 'comment.add', 'hyperlink.set']);
    const sheet = workbook.getSheet('sheet-1');
    assert.equal(sheet.review.getNoteAt(0, 0)?.text, 'Check total');
    assert.equal(sheet.review.getThread('thread-1')?.text, 'Please review');
    assert.equal(getCellHyperlink(sheet, 1, 0)?.target.kind, 'url');
    assert.equal(serializeHyperlink({ id: 'h1', target: { kind: 'name', name: 'SalesTotal' } }), '#name:SalesTotal');
  } finally {
    close();
  }
});

test('hyperlink commands fail before commit when the target is invalid or absent', async () => {
  const { workbook, runtime, close } = await setup('review-command-rejection');
  try {
    await assert.rejects(() => runtime.execute('hyperlink.set', {
      sheetId: 'sheet-1', row: 0, column: 0,
      hyperlink: { id: 'bad', target: { kind: 'url', url: 'not-a-url' } },
    }), /Invalid hyperlink URL/);
    await assert.rejects(() => runtime.execute('hyperlink.remove', {
      sheetId: 'sheet-1', row: 0, column: 0,
    }), /HYPERLINK_NOT_FOUND/);
    assert.equal(workbook.revision, 0);
    assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
  } finally {
    close();
  }
});
