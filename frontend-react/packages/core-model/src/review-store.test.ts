import assert from 'node:assert/strict';
import test from 'node:test';
import { migrateStoredWorkbookSnapshot } from './snapshot';
import { ReviewStore } from './review-store';

const note = (id: string, text = 'note') => ({ id, author: 'u', text, createdAt: '2026-01-01', visible: true });
const thread = (id: string, row: number, column: number) => ({ id, sheetId: 'sheet-1', row, column, author: 'u', text: 'comment', createdAt: '2026-01-01', replies: [] });

test('ReviewStore indexes blank-cell review and rejects identity collisions without mutation', () => {
  const store = new ReviewStore('sheet-1');
  store.setNote(10, 20, note('n1'));
  store.addThread(thread('t1', 10, 20));
  assert.equal(store.getNoteAt(10, 20)?.id, 'n1');
  assert.equal(store.getThreadsAt(10, 20)[0]?.id, 't1');
  assert.throws(() => store.setNote(11, 20, note('n1', 'conflict')), /already belongs/);
  assert.equal(store.getNoteAt(10, 20)?.text, 'note');
  assert.equal(store.noteCount, 1);
  assert.equal(store.threadCount, 1);
});

test('ReviewStore snapshot validation rejects dangling and incompatible indexes', () => {
  const store = new ReviewStore('sheet-1');
  store.setNote(0, 0, note('n1'));
  store.addThread(thread('t1', 1, 1));
  const snapshot = store.toSnapshot();
  delete snapshot.notesById.n1;
  assert.throws(() => ReviewStore.fromSnapshot('sheet-1', snapshot), /missing id/);
  const incompatible = store.toSnapshot();
  incompatible.threadIdsByCell['1:2'] = ['t1'];
  assert.throws(() => ReviewStore.fromSnapshot('sheet-1', incompatible), /incompatible/);
});

test('v7 snapshot migration moves legacy review into the canonical store and removes cell metadata', () => {
  const legacy = {
    schema: 'WorkbookSnapshot',
    version: 7,
    unitId: 'unit-1',
    name: 'Workbook',
    dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
    calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: false, fullCalculationOnLoad: false },
    dataModel: { sources: [], tables: [], relationships: [], views: [] },
    sheets: [{
      kind: 'worksheet', id: 'sheet-1', name: 'Sheet1', rowCount: 20, columnCount: 20, cells: { '3': { '4': { value: null, note: note('cell-note'), comment: thread('cell-thread', 0, 0) } } },
      notes: [{ row: 1, column: 2, note: note('sheet-note') }], commentThreads: [], merges: [], pane: { kind: 'none' }, pivots: [], sparklines: [], drawings: [], drawingPayloads: {},
      defaultRowHeightPx: 20, defaultColumnWidthPx: 64,
    }],
  } as any;
  const migrated = migrateStoredWorkbookSnapshot(legacy);
  const review = migrated.sheets[0]!.review;
  assert.equal(review.notesByCell['1:2'], 'sheet-note');
  assert.equal(review.notesByCell['3:4'], 'cell-note');
  assert.equal(review.threadIdsByCell['3:4']![0], 'cell-thread');
  assert.equal('note' in (migrated.sheets[0]!.cells['3']!['4']!), false);
  assert.equal('comment' in (migrated.sheets[0]!.cells['3']!['4']!), false);
});

test('v7 snapshot migration fails closed on conflicting legacy identities', () => {
  const legacy = {
    schema: 'WorkbookSnapshot', version: 7, unitId: 'unit-1', name: 'Workbook',
    dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
    calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: false, fullCalculationOnLoad: false },
    dataModel: { sources: [], tables: [], relationships: [], views: [] },
    sheets: [{ kind: 'worksheet', id: 'sheet-1', name: 'Sheet1', rowCount: 2, columnCount: 2, cells: {}, notes: [{ row: 0, column: 0, note: note('duplicate', 'first') }, { row: 0, column: 1, note: note('duplicate', 'second') }], commentThreads: [], merges: [], pane: { kind: 'none' }, pivots: [], sparklines: [], drawings: [], drawingPayloads: {}, defaultRowHeightPx: 20, defaultColumnWidthPx: 64 }],
  } as any;
  assert.throws(() => migrateStoredWorkbookSnapshot(legacy), /REVIEW_MIGRATION_CONFLICT/);
});
