import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  findAtCursor,
  findCursorFor,
  FindIndex,
  matchesFindText,
  planFind,
  replaceFindText,
} from './find-replace';

function range(sheetId: string, startRow = 0, endRow = 9, startColumn = 0, endColumn = 9) {
  return { sheetId, startRow, endRow, startColumn, endColumn };
}

test('Find planner orders cells and target families deterministically and keeps Find read-only', () => {
  const workbook = new WorkbookModel('find-planner', 'Find planner');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(1, 0, { value: 'needle' });
  sheet.cells.set(0, 1, { value: 'needle' });
  sheet.cells.set(2, 0, { formula: '=needle', value: null });
  sheet.review.setNote(3, 0, { id: 'note-1', author: 'u', text: 'needle note', createdAt: 'now', visible: true });
  sheet.review.addThread({ id: 'comment-1', sheetId: sheet.id, row: 4, column: 0, author: 'u', text: 'needle comment', createdAt: 'now', replies: [] });
  const before = workbook.snapshot();

  const result = planFind(workbook, { sheetId: sheet.id, query: 'needle', searchOrder: 'rows', scope: 'sheet', targets: ['values', 'formulas', 'notes', 'comments'], matchCase: true });
  assert.deepEqual(result.matches.map((match) => `${match.row}:${match.column}:${match.target}`), [
    '0:1:values', '1:0:values', '2:0:formulas', '3:0:notes', '4:0:comments',
  ]);
  assert.deepEqual(workbook.snapshot(), before);
  assert.equal(findAtCursor(result.matches, null, 'next')?.key, result.matches[0]?.key);
  assert.equal(findAtCursor(result.matches, null, 'previous')?.key, result.matches.at(-1)?.key);
  assert.equal(findAtCursor(result.matches, findCursorFor(result.matches[0]!), 'next')?.key, result.matches[1]?.key);
  assert.equal(findAtCursor(result.matches, findCursorFor(result.matches[0]!), 'previous')?.key, result.matches.at(-1)?.key);
});

test('Find planner respects selection, workbook scope, case, entire-cell, and wildcard semantics', () => {
  const workbook = new WorkbookModel('find-options', 'Find options');
  const first = workbook.getSheet(workbook.primarySheetId);
  const second = workbook.addSheet('sheet-2', 'Sheet2', 10, 10);
  first.cells.set(0, 0, { value: 'Alpha beta' });
  first.cells.set(5, 0, { value: 'alpha' });
  second.cells.set(0, 0, { value: 'alpha' });
  assert.equal(planFind(workbook, { sheetId: first.id, query: 'alpha', searchOrder: 'rows', scope: 'sheet', targets: ['values'], matchCase: false }).total, 2);
  assert.equal(planFind(workbook, { sheetId: first.id, query: 'alpha', searchOrder: 'rows', scope: 'sheet', targets: ['values'], matchCase: true }).total, 1);
  assert.equal(planFind(workbook, { sheetId: first.id, query: 'alpha', searchOrder: 'rows', scope: 'sheet', targets: ['values'], entireCell: true }).total, 1);
  assert.equal(planFind(workbook, { sheetId: first.id, query: 'a*', searchOrder: 'rows', scope: 'selection', range: range(first.id, 5, 5, 0, 0), targets: ['values'], wildcard: true }).total, 1);
  assert.equal(planFind(workbook, { sheetId: first.id, query: 'alpha', searchOrder: 'rows', scope: 'workbook', targets: ['values'] }).total, 3);
  assert.equal(planFind(workbook, { sheetId: first.id, query: 'Alpha', searchOrder: 'rows', scope: 'selection', range: range(first.id, 5, 5, 0, 0), targets: ['values'] }).total, 1);
});

test('Find planner supports canonical row-major and column-major search order', () => {
  const workbook = new WorkbookModel('find-order', 'Find order');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 1, { value: 'needle' });
  sheet.cells.set(1, 0, { value: 'needle' });
  const base = { sheetId: sheet.id, query: 'needle', scope: 'sheet' as const, targets: ['values'] as const, matchCase: true };
  assert.deepEqual(planFind(workbook, { ...base, searchOrder: 'rows' }).matches.map((match) => `${match.row}:${match.column}`), ['0:1', '1:0']);
  assert.deepEqual(planFind(workbook, { ...base, searchOrder: 'columns' }).matches.map((match) => `${match.row}:${match.column}`), ['1:0', '0:1']);
});

test('FindIndex searches sparse authored and review entries without logical-row scans', () => {
  const workbook = new WorkbookModel('find-index', 'Find index');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(900000, 2, { value: 'needle' });
  sheet.review.setNote(900001, 2, { id: 'note-index', author: 'u', text: 'needle note', createdAt: 'now', visible: true });
  const index = new FindIndex(workbook);
  const result = index.search({ sheetId: sheet.id, query: 'needle', searchOrder: 'rows', scope: 'sheet', targets: ['values', 'notes'] });
  assert.deepEqual(result.matches.map((match) => match.target), ['values', 'notes']);
  const revision = index.getRevision();
  sheet.cells.set(900000, 2, { value: 'changed' });
  index.rebuildSheet(sheet.id);
  assert.equal(index.getRevision() > revision, true);
  assert.equal(index.search({ sheetId: sheet.id, query: 'needle', searchOrder: 'rows', scope: 'sheet', targets: ['values'] }).total, 0);
});

test('Find matcher replaces literal wildcard matches without replacement-string expansion', () => {
  assert.equal(matchesFindText('A*B', { query: 'A~*B', wildcard: true }), true);
  assert.equal(replaceFindText('ab ab', { query: 'a?', matchCase: false, wildcard: true }, '$1'), '$1 $1');
  assert.equal(replaceFindText('prefix', { query: 'prefix', entireCell: true }, '0'), '0');
  assert.equal(replaceFindText('prefix', { query: 'fix', entireCell: true }, '0'), undefined);
});
