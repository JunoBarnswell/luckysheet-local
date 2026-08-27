import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { createFormulaError } from '@react-sheets/formula-engine';
import { registerSheetCommands } from './index';
import './fill-series.test';
import { parseReplacementValue as parseReplacementValueWithContext } from './home-commands';
import type { CellInputInterpretationContext } from './text-input';

const TEST_INPUT_CONTEXT: CellInputInterpretationContext = {
  sourceKind: 'find-replace', cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900',
  referenceDate: { year: 2026, month: 8, day: 27, hour: 0, minute: 0, second: 0, millisecond: 0 },
};
const parseReplacementValue = (text: string) => parseReplacementValueWithContext(text, TEST_INPUT_CONTEXT);

function setup(): { workbook: WorkbookModel; runtime: CommandRuntime } {
  const workbook = new WorkbookModel('home-commands', 'Home Commands');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  return { workbook, runtime };
}

test('AutoSum uses resolved formula results and keeps one reversible remote-replayable operation', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { formula: '=10', value: null });
  sheet.cells.set(1, 0, { formula: '=20', value: null });
  sheet.cells.set(2, 0, { formula: '=30', value: null });
  runtime.setCellValueResolver((_sheet, row, column) => column === 0 ? (row + 1) * 10 : undefined);
  const before = workbook.snapshot();

  runtime.execute('formula.autosum', {
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    target: { row: 3, column: 0 },
  });
  assert.equal(sheet.cells.get(3, 0)?.formula, '=SUM(A1:A3)');
  assert.equal(runtime.getHistoryDepth().undo, 1);
  const entry = runtime.getUndoEntries().at(-1)!;

  const remoteWorkbook = WorkbookModel.fromSnapshot(before);
  const remoteRuntime = new CommandRuntime(remoteWorkbook);
  registerSheetCommands(remoteRuntime);
  remoteRuntime.applyRemoteMutations(entry.redo);
  assert.deepEqual(remoteWorkbook.snapshot().sheets, workbook.snapshot().sheets);

  assert.equal(runtime.undo(), true);
  assert.equal(sheet.cells.get(3, 0), undefined);
  assert.equal(runtime.redo(), true);
  assert.equal(sheet.cells.get(3, 0)?.formula, '=SUM(A1:A3)');
});

test('AutoSum treats zero and FALSE as typed values and rejects errors, self-reference, and unsafe targets', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 0 });
  sheet.cells.set(1, 0, { value: false });
  runtime.setCellValueResolver(() => undefined);
  runtime.execute('formula.autosum', {
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    target: { row: 2, column: 0 },
  });
  assert.equal(sheet.cells.get(2, 0)?.formula, '=SUM(A1:A1)');

  const errorCase = setup();
  const errorSheet = errorCase.workbook.getSheet(errorCase.workbook.primarySheetId);
  errorSheet.cells.set(0, 0, { formula: '=1/0', value: null });
  errorCase.runtime.setCellValueResolver(() => createFormulaError('#DIV/0!', 'division by zero'));
  assert.throws(() => errorCase.runtime.execute('formula.autosum', {
    sheetId: errorSheet.id,
    range: { sheetId: errorSheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 0 },
  }), /no numeric result/);
  assert.equal(errorSheet.cells.get(1, 0), undefined);
  assert.equal(errorCase.runtime.getHistoryDepth().undo, 0);

  const unsafe = setup();
  const unsafeSheet = unsafe.workbook.getSheet(unsafe.workbook.primarySheetId);
  unsafeSheet.cells.set(0, 0, { value: 10 });
  unsafeSheet.cells.set(1, 0, { formula: '=A1', value: null });
  assert.throws(() => unsafe.runtime.execute('formula.autosum', {
    sheetId: unsafeSheet.id,
    range: { sheetId: unsafeSheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 0 },
  }), /not blank/);
  assert.equal(unsafe.runtime.getHistoryDepth().undo, 0);

  const selfReference = setup();
  const selfSheet = selfReference.workbook.getSheet(selfReference.workbook.primarySheetId);
  selfSheet.cells.set(0, 0, { value: 10 });
  assert.throws(() => selfReference.runtime.execute('formula.autosum', {
    sheetId: selfSheet.id,
    range: { sheetId: selfSheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 0 },
  }), /include its target/);
  assert.equal(selfReference.runtime.getHistoryDepth().undo, 0);
});

test('AutoSum fails closed for protected, merged, and spill-child targets', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 10 });
  sheet.protectionRules.push({ id: 'locked-target', scope: 'range', sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }, locked: true, allow: {} });
  assert.throws(() => runtime.execute('formula.autosum', {
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 0 },
  }), /protected/);

  const merged = setup();
  const mergedSheet = merged.workbook.getSheet(merged.workbook.primarySheetId);
  mergedSheet.cells.set(0, 0, { value: 10 });
  mergedSheet.merges.push({ range: { sheetId: mergedSheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 }, anchor: { row: 1, column: 0 } });
  assert.throws(() => merged.runtime.execute('formula.autosum', {
    sheetId: mergedSheet.id,
    range: { sheetId: mergedSheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 1 },
  }), /merged range/);

  const spill = setup();
  const spillSheet = spill.workbook.getSheet(spill.workbook.primarySheetId);
  spillSheet.cells.set(0, 0, { value: 10 });
  spillSheet.spillRanges.push({
    sheetId: spillSheet.id,
    anchor: { row: 1, column: 0 },
    range: { sheetId: spillSheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 },
    values: [[1, 2]],
    state: 'ok',
  });
  assert.throws(() => spill.runtime.execute('formula.autosum', {
    sheetId: spillSheet.id,
    range: { sheetId: spillSheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 1 },
  }), /spill child/);

  const styledBlank = setup();
  const styledSheet = styledBlank.workbook.getSheet(styledBlank.workbook.primarySheetId);
  styledSheet.cells.set(0, 0, { value: 10 });
  styledSheet.cells.set(1, 0, { value: undefined as never, style: { bold: true } });
  styledBlank.runtime.execute('formula.autosum', {
    sheetId: styledSheet.id,
    range: { sheetId: styledSheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    target: { row: 1, column: 0 },
  });
  assert.equal(styledSheet.cells.get(1, 0)?.formula, '=SUM(A1:A1)');
  assert.equal(styledSheet.cells.get(1, 0)?.style?.bold, true);
});

test('formula.autosum emits a formula and has one undo entry', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 10 });
  sheet.cells.set(1, 0, { value: 20 });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };
  const result = runtime.execute('formula.autosum', { sheetId: sheet.id, range });
  assert.equal(result.affectedRanges[0]?.startRow, 2);
  assert.equal(sheet.cells.get(2, 0)?.formula, '=SUM(A1:A2)');
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.undo();
  assert.equal(sheet.cells.get(2, 0), undefined);
});

test('format painter changes only presentation and supports undo/redo', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'source', style: { bold: true, background: '#ff0' }, numberFormat: '0.00' });
  sheet.cells.set(0, 1, { value: 'target', style: { italic: true } });
  const range = (column: number) => ({ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: column, endColumn: column });
  const pattern = {
    rowCount: 1,
    columnCount: 1,
    cells: [[{ style: { bold: true, background: '#ff0' }, numberFormat: '0.00' }]],
  };
  sheet.cells.set(0, 0, { value: 'source', style: { underline: true }, numberFormat: '@' });
  runtime.execute('format.painter.apply', { sheetId: sheet.id, targetRange: range(1), pattern });
  assert.equal(sheet.cells.get(0, 1)?.value, 'target');
  assert.equal(sheet.cells.get(0, 1)?.style?.bold, true);
  assert.equal(sheet.cells.get(0, 1)?.style?.italic, undefined);
  assert.equal(sheet.cells.get(0, 1)?.numberFormat, '0.00');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 1)?.style?.italic, true);
  runtime.redo();
  assert.equal(sheet.cells.get(0, 1)?.style?.bold, true);
});

test('format painter applies a frozen multi-cell pattern across sheets and replays only style.set mutations', () => {
  const { workbook, runtime } = setup();
  const source = workbook.getSheet(workbook.primarySheetId);
  const target = workbook.addSheet('target', 'Target');
  source.cells.set(0, 0, { value: 'A', style: { bold: true }, numberFormat: '0.00' });
  source.cells.set(0, 1, { value: 'B', style: { italic: true }, numberFormat: '@' });
  const before = workbook.snapshot();
  const result = runtime.execute('format.painter.apply', {
    sheetId: target.id,
    targetRange: { sheetId: target.id, startRow: 2, endRow: 3, startColumn: 2, endColumn: 5 },
    pattern: {
      rowCount: 1,
      columnCount: 2,
      cells: [[
        { style: { bold: true }, numberFormat: '0.00' },
        { style: { italic: true }, numberFormat: '@' },
      ]],
    },
  });
  assert.equal(result.mutationCount, 8);
  assert.equal(target.cells.get(2, 2)?.style?.bold, true);
  assert.equal(target.cells.get(2, 3)?.style?.italic, true);
  assert.equal(target.cells.get(2, 4)?.numberFormat, '0.00');
  assert.equal(runtime.getUndoEntries().at(-1)?.redo.every((mutation) => mutation.id === 'style.set'), true);

  const remoteWorkbook = WorkbookModel.fromSnapshot(before);
  const remoteRuntime = new CommandRuntime(remoteWorkbook);
  registerSheetCommands(remoteRuntime);
  remoteRuntime.applyRemoteMutations(runtime.getUndoEntries().at(-1)!.redo);
  assert.deepEqual(remoteWorkbook.snapshot().sheets, workbook.snapshot().sheets);
});

test('format painter removes target presentation that is absent from the captured pattern', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'target', style: { italic: true }, numberFormat: '0.00' });
  runtime.execute('format.painter.apply', {
    sheetId: sheet.id,
    targetRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    pattern: { rowCount: 1, columnCount: 1, cells: [[{}]] },
  });
  assert.equal(sheet.cells.get(0, 0)?.value, 'target');
  assert.equal(sheet.cells.get(0, 0)?.style, undefined);
  assert.equal(sheet.cells.get(0, 0)?.numberFormat, undefined);
});

test('filter toggle, clearCriteria and reapply preserve the filter contract', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const range = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
  runtime.execute('sheet.autoFilter.toggle', { sheetId: sheet.id, range });
  assert.ok(sheet.autoFilter);
  sheet.autoFilter!.columns[0] = { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['A'], includeBlank: false } };
  runtime.execute('sheet.autoFilter.clearCriteria', { sheetId: sheet.id });
  assert.equal(Object.values(sheet.autoFilter?.columns ?? {}).some((column) => Boolean(column.criterion)), false);
  assert.deepEqual(runtime.execute('sheet.autoFilter.reapply', { sheetId: sheet.id }).affectedRanges, [range]);
  runtime.execute('sheet.autoFilter.toggle', { sheetId: sheet.id, range });
  assert.equal(sheet.autoFilter, undefined);
});

test('replace all is one history action and restores every cell', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'old value' });
  sheet.cells.set(0, 1, { value: 'old' });
  runtime.execute('sheet.range.replace', { sheetId: sheet.id, find: 'old', replace: 'new', inputContext: TEST_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 0)?.value, 'new value');
  assert.equal(sheet.cells.get(0, 1)?.value, 'new');
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'old value');
  assert.equal(sheet.cells.get(0, 1)?.value, 'old');
});

test('Go To Special returns discontiguous formula and validation matches', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 1 });
  sheet.cells.set(1, 0, { value: null, formula: '=A1+1' });
  sheet.cells.set(2, 0, { value: 3 });
  sheet.dataValidations.push({
    id: 'home-validation',
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }],
    type: 'whole',
    formula1: '0',
    alertStyle: 'stop',
  });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 };
  assert.deepEqual(runtime.execute('selection.gotoSpecial', { sheetId: sheet.id, range, kind: 'formulas' }).affectedRanges, [
    { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
  ]);
  assert.equal(runtime.execute('selection.gotoSpecial', { sheetId: sheet.id, range, kind: 'data-validation' }).affectedRanges.length, 3);
});

test('range move and style preset are atomic and reversible', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: null, formula: '=B1', style: { bold: true } });
  const sourceRange = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  runtime.execute('sheet.range.move', { sheetId: sheet.id, sourceRange, targetOrigin: { row: 1, column: 1 } });
  assert.equal(sheet.cells.get(0, 0), undefined);
  assert.equal(sheet.cells.get(1, 1)?.formula, '=B1');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.formula, '=B1');
  runtime.execute('sheet.style.preset.apply', { sheetId: sheet.id, ranges: [sourceRange], preset: 'good' });
  assert.equal(sheet.cells.get(0, 0)?.styleId, 'good');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.styleId, undefined);
});

test('workbook cell templates apply style, editor and validation through one command transaction', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const range = { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 };
  runtime.execute('workbook.cellTemplate.set', {
    sheetId: sheet.id,
    template: {
      id: 'status',
      name: 'Status',
      style: { background: '#e2f0d9', indent: 1 },
      editor: { kind: 'list', values: ['Open', 'Closed'] },
    },
  });
  runtime.execute('sheet.cellTemplate.apply', { sheetId: sheet.id, ranges: [range], templateId: 'status' });
  assert.equal(sheet.cells.get(1, 1)?.style?.indent, 1);
  assert.equal(sheet.cells.get(1, 1)?.editor?.kind, 'list');
  assert.deepEqual(sheet.dataValidations[0]?.listSource, { kind: 'values', values: ['Open', 'Closed'] });
  runtime.undo();
  assert.equal(sheet.cells.get(1, 1), undefined);
  runtime.redo();
  assert.equal(sheet.cells.get(1, 1)?.editor?.kind, 'list');
});

test('Home cell commands fail closed on block-backed data regions', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.addDataRegion({
    id: 'block-home',
    sourceId: 'source-home',
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    headerRow: 0,
    revision: 1,
  });
  const range = { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 };
  assert.throws(() => runtime.execute('sheet.range.fill', {
    sheetId: sheet.id,
    sourceRange: range,
    targetRange: range,
    direction: 'down',
    mode: 'copy',
  }), /data-region/);
  assert.throws(() => runtime.execute('formula.autosum', { sheetId: sheet.id, range }), /data-region/);
  assert.throws(() => runtime.execute('sheet.range.replace', { sheetId: sheet.id, find: 'a', replace: 'b', range, inputContext: TEST_INPUT_CONTEXT }), /data-region/);
});

test('canonical fill is one mutation, preserves seeds, and replays/undoes atomically', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const sourceRange = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };
  const targetRange = { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 };
  sheet.cells.set(0, 0, { value: 1 });
  sheet.cells.set(1, 0, { value: 3 });
  const before = workbook.snapshot();
  const seen: string[] = [];
  runtime.onMutation((mutation, source) => { if (source === 'command') seen.push(mutation.id); });

  const result = runtime.execute('sheet.range.fill', {
    sheetId: sheet.id,
    sourceRange,
    targetRange,
    direction: 'down',
    mode: 'series',
  });
  assert.equal(result.mutationCount, 1);
  assert.deepEqual(seen, ['fill.applied']);
  assert.deepEqual([2, 3, 4].map((row) => sheet.cells.get(row, 0)?.value), [5, 7, 9]);
  assert.deepEqual([sheet.cells.get(0, 0)?.value, sheet.cells.get(1, 0)?.value], [1, 3]);
  assert.equal(runtime.getHistoryDepth().undo, 1);

  const entry = runtime.getUndoEntries()[0];
  assert.ok(entry);
  assert.equal(entry.redo.length, 1);
  const remoteWorkbook = WorkbookModel.fromSnapshot(before);
  const remoteRuntime = new CommandRuntime(remoteWorkbook);
  registerSheetCommands(remoteRuntime);
  remoteRuntime.applyRemoteMutations(entry.redo);
  const remoteSheet = remoteWorkbook.getSheet(sheet.id);
  assert.deepEqual([2, 3, 4].map((row) => remoteSheet.cells.get(row, 0)?.value), [5, 7, 9]);

  assert.equal(runtime.undo(), true);
  assert.equal(sheet.cells.get(2, 0), undefined);
  assert.deepEqual([sheet.cells.get(0, 0)?.value, sheet.cells.get(1, 0)?.value], [1, 3]);
  assert.equal(runtime.redo(), true);
  assert.deepEqual([2, 3, 4].map((row) => sheet.cells.get(row, 0)?.value), [5, 7, 9]);
});

test('fill fails closed for protected or block-backed targets without history or partial writes', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const sourceRange = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  const targetRange = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 };
  sheet.cells.set(0, 0, { value: 10 });
  sheet.protectionRules.push({
    id: 'fill-protected',
    scope: 'range',
    range: targetRange,
    locked: true,
    allow: {},
  });
  const before = workbook.snapshot();
  assert.throws(() => runtime.execute('sheet.range.fill', {
    sheetId: sheet.id,
    sourceRange,
    targetRange,
    direction: 'down',
    mode: 'series',
  }), /protected/);
  assert.equal(runtime.getHistoryDepth().undo, 0);
  assert.deepEqual(workbook.snapshot(), before);

  sheet.protectionRules.length = 0;
  sheet.addDataRegion({
    id: 'fill-block',
    sourceId: 'fill-source',
    range: targetRange,
    headerRow: 0,
    revision: 1,
  });
  assert.throws(() => runtime.execute('sheet.range.fill', {
    sheetId: sheet.id,
    sourceRange,
    targetRange,
    direction: 'down',
    mode: 'series',
  }), /data-region/);
  assert.equal(runtime.getHistoryDepth().undo, 0);
  assert.equal(sheet.cells.get(1, 0), undefined);
});

test('merge center keeps the anchor, clears non-anchor contents, and undoes atomically', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'anchor' });
  sheet.cells.set(0, 1, { value: 'discarded' });
  sheet.cells.set(1, 0, { value: 'discarded-too' });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
  runtime.execute('sheet.merge.center', { sheetId: sheet.id, range });
  assert.equal(sheet.cells.get(0, 0)?.value, 'anchor');
  assert.equal(sheet.cells.get(0, 1)?.value, null);
  assert.equal(sheet.cells.get(1, 0)?.value, null);
  assert.equal(sheet.merges.length, 1);
  runtime.undo();
  assert.equal(sheet.cells.get(0, 1)?.value, 'discarded');
  assert.equal(sheet.cells.get(1, 0)?.value, 'discarded-too');
  assert.equal(sheet.merges.length, 0);
});

test('merge cells preserves alignment while merge across merges each row without centering', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'r1' });
  sheet.cells.set(0, 1, { value: 'discard-r1' });
  sheet.cells.set(1, 0, { value: 'r2', style: { horizontalAlignment: 'left' } });
  sheet.cells.set(1, 1, { value: 'discard-r2' });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
  runtime.execute('sheet.merge.cells', { sheetId: sheet.id, range });
  assert.equal(sheet.cells.get(0, 0)?.style?.horizontalAlignment, undefined);
  assert.equal(sheet.cells.get(0, 1)?.value, null);
  runtime.undo();
  runtime.execute('sheet.merge.across', { sheetId: sheet.id, range });
  assert.equal(sheet.merges.length, 2);
  assert.equal(sheet.cells.get(1, 0)?.style?.horizontalAlignment, 'left');
  assert.equal(sheet.cells.get(0, 0)?.style?.horizontalAlignment, undefined);
  assert.equal(sheet.cells.get(1, 1)?.value, null);
  runtime.undo();
  assert.equal(sheet.merges.length, 0);
  assert.equal(sheet.cells.get(1, 1)?.value, 'discard-r2');
});

test('unmerge removes spans without resurrecting discarded content', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'anchor' });
  sheet.cells.set(0, 1, { value: 'discarded' });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 };
  runtime.execute('sheet.merge.cells', { sheetId: sheet.id, range });
  runtime.execute('sheet.merge.unmerge', { sheetId: sheet.id, range });
  assert.equal(sheet.merges.length, 0);
  assert.equal(sheet.cells.get(0, 0)?.value, 'anchor');
  assert.equal(sheet.cells.get(0, 1)?.value, null);
  runtime.undo();
  assert.equal(sheet.merges.length, 1);
  assert.equal(sheet.cells.get(0, 1)?.value, null);
});

test('dataRegion.materialize commits the prepared payload and restores it on undo', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const range = { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };
  const region = { id: 'home-region', sourceId: 'home-source', range, headerRow: 0, revision: 1 };
  const manifest = {
    schema: 'DataSourceManifest' as const,
    version: 1 as const,
    id: 'home-source',
    name: 'Home source',
    kind: 'worksheet-range' as const,
    rowCount: 1,
    fields: [{ id: 'value', name: 'Value', ordinal: 0, type: 'number' as const }],
    blockRowCount: 64,
    blocks: [],
    revision: 1,
  };
  workbook.dataModel.sources.set(manifest.id, manifest);
  sheet.addDataRegion(region);
  sheet.cells.set(0, 0, { value: 'Value' });
  sheet.cells.set(1, 0, { value: 10 });
  const payload = {
    sheetId: sheet.id,
    region,
    regionIndex: 0,
    manifest,
    willRemoveSource: true,
    range,
    previousCells: [{ row: 0, column: 0, cell: { value: 'Value' } }, { row: 1, column: 0, cell: { value: 10 } }],
    materializedCells: [{ row: 0, column: 0, cell: { value: 'Value' } }, { row: 1, column: 0, cell: { value: 10 } }],
    materializedCellCount: 2,
  };
  runtime.execute('dataRegion.materialize.commit', payload);
  assert.equal(sheet.dataRegions.length, 0);
  assert.equal(workbook.dataModel.sources.has(manifest.id), false);
  runtime.undo();
  assert.equal(sheet.dataRegions.length, 1);
  assert.equal(workbook.dataModel.sources.has(manifest.id), true);
  assert.equal(sheet.cells.get(1, 0)?.value, 10);
});

test('table style and drawing pane commands are canonical, persisted, and undoable', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.sheetTables.push({
    id: 'home-table',
    sheetId: sheet.id,
    name: 'HomeTable',
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    hasHeaderRow: true,
    hasTotalRow: false,
    showBandedRows: false,
    showBandedColumns: false,
    showFirstColumn: false,
    showLastColumn: false,
    showFilterButton: true,
    autoExpand: 'both',
    columns: [{ id: 'a', name: 'A' }, { id: 'b', name: 'B' }],
  });
  runtime.execute('sheetTable.style.set', { sheetId: sheet.id, tableId: 'home-table', styleName: 'TableStyleMedium2', showBandedRows: true });
  assert.equal(sheet.sheetTables[0]?.styleName, 'TableStyleMedium2');
  assert.equal(sheet.sheetTables[0]?.showBandedRows, true);
  runtime.undo();
  assert.equal(sheet.sheetTables[0]?.styleName, undefined);
  assert.equal(sheet.sheetTables[0]?.showBandedRows, false);

  const drawing = {
    id: 'home-drawing',
    sheetId: sheet.id,
    kind: 'shape' as const,
    anchor: { kind: 'absolute' as const },
    transform: { x: 0, y: 0, width: 10, height: 10 },
    zIndex: 0,
    payloadId: 'home-payload',
  };
  sheet.drawings.push(drawing);
  sheet.drawingPayloads.set(drawing.payloadId, { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' });
  runtime.execute('drawing.visibility.set', { sheetId: sheet.id, drawingId: drawing.id, visible: false });
  runtime.execute('drawing.rename', { sheetId: sheet.id, drawingId: drawing.id, name: 'Named object' });
  assert.equal(sheet.drawings[0]?.visible, false);
  assert.equal(sheet.drawings[0]?.name, 'Named object');
  runtime.undo();
  runtime.undo();
  assert.equal(sheet.drawings[0]?.visible, true);
  assert.equal(sheet.drawings[0]?.name, undefined);
});

test('replacement parser preserves every supported tagged value', () => {
  assert.deepEqual(parseReplacementValue('0'), { kind: 'number', value: 0 });
  assert.equal(Object.is(parseReplacementValue('-0').value, -0), true);
  assert.deepEqual(parseReplacementValue('-12.5'), { kind: 'number', value: -12.5 });
  assert.deepEqual(parseReplacementValue('TRUE'), { kind: 'boolean', value: true });
  assert.deepEqual(parseReplacementValue('false'), { kind: 'boolean', value: false });
  assert.deepEqual(parseReplacementValue("'0"), { kind: 'text', value: "'0" });
  assert.deepEqual(parseReplacementValue('=A1+1'), { kind: 'formula', value: null, formula: '=A1+1' });
  assert.deepEqual(parseReplacementValue('#DIV/0!'), { kind: 'error', value: null, code: '#DIV/0!' });
  assert.deepEqual(parseReplacementValue(''), { kind: 'empty', value: null });
});

test('replacement parser rejects invalid formulas and overflowing numeric literals', () => {
  assert.throws(() => parseReplacementValue('='), /Invalid replacement formula/);
  assert.throws(() => parseReplacementValue('1e999'), /not finite/);
});

test('replace writes numeric zero as a number and is one undo/redo/replay operation', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 10 });
  sheet.cells.set(1, 0, { value: '10' });
  const initial = workbook.snapshot();

  runtime.execute('sheet.range.replace', {
    sheetId: sheet.id,
    find: '10',
    replace: '0',
    inputContext: TEST_INPUT_CONTEXT,
    entireCell: true,
  });

  assert.equal(sheet.cells.get(0, 0)?.value, 0);
  assert.equal(typeof sheet.cells.get(0, 0)?.value, 'number');
  assert.equal(sheet.cells.get(1, 0)?.value, 0);
  assert.equal(typeof sheet.cells.get(1, 0)?.value, 'number');
  assert.equal(runtime.getHistoryDepth().undo, 1);

  const entry = runtime.getUndoEntries()[0];
  assert.ok(entry);
  const remoteWorkbook = WorkbookModel.fromSnapshot(initial);
  const remoteRuntime = new CommandRuntime(remoteWorkbook);
  registerSheetCommands(remoteRuntime);
  remoteRuntime.applyRemoteMutations(entry.redo);
  const remoteSheet = remoteWorkbook.getSheet(sheet.id);
  assert.equal(remoteSheet.cells.get(0, 0)?.value, 0);
  assert.equal(typeof remoteSheet.cells.get(0, 0)?.value, 'number');
  assert.equal(remoteSheet.cells.get(1, 0)?.value, 0);

  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
  assert.equal(sheet.cells.get(1, 0)?.value, '10');
  runtime.redo();
  assert.equal(sheet.cells.get(0, 0)?.value, 0);
  assert.equal(sheet.cells.get(1, 0)?.value, 0);
});

test('replace parses negative and fractional replacements without previous-value coercion', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 42 });
  runtime.execute('sheet.range.replace', {
    sheetId: sheet.id,
    find: '42',
    replace: '-3.25',
    inputContext: TEST_INPUT_CONTEXT,
    entireCell: true,
  });
  assert.equal(sheet.cells.get(0, 0)?.value, -3.25);
  assert.equal(typeof sheet.cells.get(0, 0)?.value, 'number');
});

test('replace handles explicit formula and formula error values through the typed contract', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: null, formula: '=A1+1' });
  runtime.execute('sheet.range.replace', {
    sheetId: sheet.id,
    find: '=A1+1',
    replace: '=B1+1',
    inputContext: TEST_INPUT_CONTEXT,
    entireCell: true,
    searchIn: 'formulas',
  });
  assert.equal(sheet.cells.get(0, 0)?.formula, '=B1+1');
  assert.equal(sheet.cells.get(0, 0)?.value, null);

  sheet.cells.set(1, 0, { value: null, formulaValue: { kind: 'error', code: '#DIV/0!', message: '' } });
  runtime.execute('sheet.range.replace', {
    sheetId: sheet.id,
    find: '#DIV/0!',
    replace: '0',
    inputContext: TEST_INPUT_CONTEXT,
    entireCell: true,
    searchIn: 'values',
  });
  assert.equal(sheet.cells.get(1, 0)?.value, 0);
  assert.equal(sheet.cells.get(1, 0)?.formulaValue, undefined);
  assert.equal(sheet.cells.get(1, 0)?.formula, undefined);
});

test('empty, overflowing, and invalid formula replacements fail before any cell or history mutation', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'old' });
  sheet.cells.set(0, 1, { value: 'old' });
  const before = workbook.snapshot();

  for (const replace of ['', '1e999']) {
    assert.throws(() => runtime.execute('sheet.range.replace', {
      sheetId: sheet.id,
      find: 'old',
      replace,
      entireCell: true,
    }));
    assert.deepEqual(workbook.snapshot(), before);
    assert.equal(runtime.getHistoryDepth().undo, 0);
  }

  sheet.cells.set(0, 0, { value: null, formula: '=A1' });
  const beforeFormulaFailure = workbook.snapshot();
  assert.throws(() => runtime.execute('sheet.range.replace', {
    sheetId: sheet.id,
    find: '=A1',
    replace: 'not-a-formula',
    inputContext: TEST_INPUT_CONTEXT,
    entireCell: true,
    searchIn: 'formulas',
  }), /formula/);
  assert.deepEqual(workbook.snapshot(), beforeFormulaFailure);
  assert.equal(runtime.getHistoryDepth().undo, 0);
});

test('replace all does not re-match its own zero result', () => {
  const { workbook, runtime } = setup();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: '111' });
  runtime.execute('sheet.range.replace', {
    sheetId: sheet.id,
    find: '1',
    replace: '0',
    inputContext: TEST_INPUT_CONTEXT,
  });
  assert.equal(sheet.cells.get(0, 0)?.value, 0);
  assert.equal(typeof sheet.cells.get(0, 0)?.value, 'number');
  assert.equal(runtime.getHistoryDepth().undo, 1);
});
