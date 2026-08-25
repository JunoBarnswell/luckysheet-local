import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerFindReplaceCommands } from './commands';
import { registerReviewCommands } from '../review/commands';
import { WorkbookSession } from '../../workbook-session';

function setup(): { workbook: WorkbookModel; runtime: CommandRuntime } {
  const workbook = new WorkbookModel('find-command', 'Find command');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  registerReviewCommands(runtime);
  registerFindReplaceCommands(runtime);
  return { workbook, runtime };
}

test('find.replace replaces one/all typed cell values in one undoable transaction and replays remotely', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'before' });
  sheet.cells.set(1, 0, { value: 'before' });
  const before = workbook.snapshot();
  const result = runtime.execute('find.replace', {
    sheetId: sheet.id,
    query: 'before',
    replace: '0',
    mode: 'all',
    searchOrder: 'rows',
    scope: 'sheet',
    targets: ['values'],
  });
  assert.equal(result.mutationCount, 1);
  assert.equal(result.event?.payload.count, 2);
  assert.equal(sheet.cells.get(0, 0)?.value, 0);
  assert.equal(sheet.cells.get(1, 0)?.value, 0);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  const entry = runtime.getUndoEntries().at(-1)!;

  const remoteWorkbook = WorkbookModel.fromSnapshot(before);
  const remoteRuntime = new CommandRuntime(remoteWorkbook);
  registerSheetCommands(remoteRuntime);
  registerReviewCommands(remoteRuntime);
  registerFindReplaceCommands(remoteRuntime);
  remoteRuntime.applyRemoteMutations(entry.redo);
  assert.deepEqual(remoteWorkbook.snapshot().sheets, workbook.snapshot().sheets);
  assert.equal(runtime.undo(), true);
  assert.equal(sheet.cells.get(0, 0)?.value, 'before');
  assert.equal(runtime.redo(), true);
  assert.equal(sheet.cells.get(1, 0)?.value, 0);
});

test('find.replace updates canonical notes and comment threads atomically', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.notes.set('2:1', { id: 'note-1', author: 'u', text: 'old note', createdAt: 'now', visible: true });
  sheet.commentThreads.push({ id: 'thread-1', sheetId: sheet.id, row: 3, column: 1, author: 'u', text: 'old comment', createdAt: 'now', replies: [] });
  const result = runtime.execute('find.replace', {
    sheetId: sheet.id,
    query: 'old',
    replace: 'new',
    mode: 'all',
    searchOrder: 'rows',
    scope: 'sheet',
    targets: ['notes', 'comments'],
  });
  assert.equal(result.mutationCount, 1);
  assert.equal(result.event?.payload.count, 2);
  assert.equal(sheet.notes.get('2:1')?.text, 'new note');
  assert.equal(sheet.commentThreads[0]?.text, 'new comment');
  assert.equal(runtime.getHistoryDepth().undo, 1);
  assert.equal(runtime.undo(), true);
  assert.equal(sheet.notes.get('2:1')?.text, 'old note');
  assert.equal(sheet.commentThreads[0]?.text, 'old comment');
});

test('find.replace fails closed before applying any patch for a stale or invalid match', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'foo' });
  sheet.cells.set(1, 0, { formula: '=foo', value: null });
  assert.throws(() => runtime.execute('find.replace', {
    sheetId: sheet.id,
    query: '=foo',
    replace: 'bar',
    mode: 'all',
    searchOrder: 'rows',
    scope: 'sheet',
    targets: ['formulas'],
  }), /Formula replacement/);
  assert.equal(sheet.cells.get(0, 0)?.value, 'foo');
  assert.equal(sheet.cells.get(1, 0)?.formula, '=foo');
  assert.equal(runtime.getHistoryDepth().undo, 0);
  assert.throws(() => runtime.execute('find.replace', {
    sheetId: sheet.id,
    query: 'foo',
    replace: '',
    mode: 'all',
    searchOrder: 'rows',
    scope: 'sheet',
    targets: ['values'],
  }), /Replacement text must not be empty/);
  assert.equal(runtime.getHistoryDepth().undo, 0);
  assert.throws(() => runtime.execute('find.replace', {
    sheetId: sheet.id,
    query: 'foo',
    replace: '0',
    mode: 'one',
    searchOrder: 'rows',
    matchKey: 'sheet-1!9:9:values:',
    targets: ['values'],
  }), /no longer available/);
  assert.equal(runtime.getHistoryDepth().undo, 0);
});

test('WorkbookSession Find Next/Previous/All use transient cursor state and do not create history', () => {
  const app = new WorkbookSession();
  const sheetId = app.getActiveSheetId();
  app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'needle' } });
  app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 0, value: { value: 'needle' } });
  app['runtime'].commands.clearHistory();
  const params = { query: 'needle', searchOrder: 'rows' as const, matchCase: true, entireCell: true, wildcard: false, scope: 'sheet' as const, targets: ['values'] as const };
  assert.equal(app.findNext(params), 1);
  assert.equal(app.getSelection().activeCell.row, 0);
  assert.equal(app.findNext(params), 1);
  assert.equal(app.getSelection().activeCell.row, 1);
  assert.equal(app.findPrevious(params), 1);
  assert.equal(app.getSelection().activeCell.row, 0);
  assert.equal(app.findAll(params), 2);
  assert.equal(app['runtime'].commands.getHistoryDepth().undo, 0);
});
