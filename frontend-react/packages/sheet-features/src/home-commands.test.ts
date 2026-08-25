import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from './index';

function setup(): { workbook: WorkbookModel; runtime: CommandRuntime } {
  const workbook = new WorkbookModel('home-commands', 'Home Commands');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  return { workbook, runtime };
}

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
  runtime.execute('format.painter.apply', { sheetId: sheet.id, sourceRange: range(0), targetRange: range(1) });
  assert.equal(sheet.cells.get(0, 1)?.value, 'target');
  assert.equal(sheet.cells.get(0, 1)?.style?.bold, true);
  assert.equal(sheet.cells.get(0, 1)?.numberFormat, '0.00');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 1)?.style?.italic, true);
  runtime.redo();
  assert.equal(sheet.cells.get(0, 1)?.style?.bold, true);
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
  runtime.execute('sheet.range.replace', { sheetId: sheet.id, find: 'old', replace: 'new' });
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
  sheet.dataRegions.push({
    id: 'block-home',
    sourceId: 'source-home',
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    headerRow: 0,
    revision: 1,
  });
  const range = { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 };
  assert.throws(() => runtime.execute('sheet.range.fill', { sheetId: sheet.id, sourceRange: range, targetRange: range }), /data-region/);
  assert.throws(() => runtime.execute('formula.autosum', { sheetId: sheet.id, range }), /data-region/);
  assert.throws(() => runtime.execute('sheet.range.replace', { sheetId: sheet.id, find: 'a', replace: 'b', range }), /data-region/);
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
  sheet.dataRegions.push(region);
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
