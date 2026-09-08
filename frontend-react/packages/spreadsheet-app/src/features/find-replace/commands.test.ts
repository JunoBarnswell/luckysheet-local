import assert from 'node:assert/strict';
import test from 'node:test';
import { registerSheetCommands, type CellInputInterpretationContext } from '@react-sheets/sheet-features';
import { registerFindReplaceCommands } from './commands';
import { registerReviewCommands } from '../review/commands';
import { openCanonicalTestRuntime, seedCanonicalCells } from '../../../../core-model/src/canonical-test-runtime.test';

const TEST_INPUT_CONTEXT: CellInputInterpretationContext = {
  sourceKind: 'find-replace', cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900',
  referenceDate: { year: 2026, month: 8, day: 27, hour: 0, minute: 0, second: 0, millisecond: 0 },
};

async function setup(unitId: string) {
  const fixture = await openCanonicalTestRuntime(unitId, 'Find command');
  registerSheetCommands(fixture.runtime);
  registerReviewCommands(fixture.runtime);
  registerFindReplaceCommands(fixture.runtime);
  return fixture;
}

test('find.replace plans canonical cell.set mutations without a local reducer', async () => {
  const { workbook, runtime, close } = await setup('find-command-cells');
  try {
    await seedCanonicalCells(runtime, 'sheet-1', [
      { row: 0, column: 0, value: 'before' },
      { row: 1, column: 0, value: 'before' },
    ]);
    const planned: string[] = [];
    runtime.onMutation((mutation) => planned.push(mutation.id));

    const result = await runtime.execute('find.replace', {
      sheetId: 'sheet-1', query: 'before', replace: '0', inputContext: TEST_INPUT_CONTEXT,
      mode: 'all', searchOrder: 'rows', scope: 'sheet', targets: ['values'],
    });

    assert.equal(result.mutationCount, 2);
    assert.equal(result.event?.payload.count, 2);
    assert.deepEqual(planned, ['cell.set', 'cell.set']);
    assert.equal(runtime.registry.hasMutation('find.replaced'), false);
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0)?.value, 0);
    assert.equal(workbook.getSheet('sheet-1').cells.get(1, 0)?.value, 0);
  } finally {
    close();
  }
});

test('find.replace composes canonical review mutations for notes and comments', async () => {
  const { workbook, runtime, close } = await setup('find-command-review');
  try {
    await runtime.execute('note.set', {
      sheetId: 'sheet-1', row: 2, column: 1,
      note: { id: 'note-1', author: 'u', text: 'old note', createdAt: 'now', visible: true },
    });
    await runtime.execute('comment.add', {
      sheetId: 'sheet-1', row: 3, column: 1,
      thread: { id: 'thread-1', sheetId: 'sheet-1', row: 3, column: 1, author: 'u', text: 'old comment', createdAt: 'now', replies: [] },
    });
    const planned: string[] = [];
    runtime.onMutation((mutation) => planned.push(mutation.id));

    const result = await runtime.execute('find.replace', {
      sheetId: 'sheet-1', query: 'old', replace: 'new', inputContext: TEST_INPUT_CONTEXT,
      mode: 'all', searchOrder: 'rows', scope: 'sheet', targets: ['notes', 'comments'],
    });

    assert.equal(result.mutationCount, 2);
    assert.deepEqual(planned, ['note.set', 'comment.update']);
    assert.equal(workbook.getSheet('sheet-1').review.getNoteAt(2, 1)?.text, 'new note');
    assert.equal(workbook.getSheet('sheet-1').review.getThread('thread-1')?.text, 'new comment');
  } finally {
    close();
  }
});

test('find.replace rejects invalid plans before committing any mutation', async () => {
  const { workbook, runtime, close } = await setup('find-command-rejection');
  try {
    await seedCanonicalCells(runtime, 'sheet-1', [
      { row: 0, column: 0, value: 'foo' },
      { row: 1, column: 0, value: null, formula: '=foo' },
    ]);
    const revision = workbook.revision;
    await assert.rejects(() => runtime.execute('find.replace', {
      sheetId: 'sheet-1', query: '=foo', replace: 'bar', inputContext: TEST_INPUT_CONTEXT,
      mode: 'all', searchOrder: 'rows', scope: 'sheet', targets: ['formulas'],
    }), /Formula replacement/);
    await assert.rejects(() => runtime.execute('find.replace', {
      sheetId: 'sheet-1', query: 'foo', replace: '', inputContext: TEST_INPUT_CONTEXT,
      mode: 'all', searchOrder: 'rows', scope: 'sheet', targets: ['values'],
    }), /Replacement text must not be empty/);
    assert.equal(workbook.revision, revision);
    assert.equal(runtime.getHistoryDepth().undo, 2);
  } finally {
    close();
  }
});
