import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import {
  registerSheetCommands,
  formatTsv,
  parseTsv as parseTsvWithContext,
  parseClipboardPayload as parseClipboardPayloadWithContext,
  shiftFormula,
  FormulaRelocationError,
  copyRangeToClipboardData,
  clipboardRepresentations,
  createPasteSpecialSpec,
  createCellSetMutationParams,
  isCellSetMutationParams,
} from './index';
import type { CellInputInterpretationContext } from './text-input';

const TEST_INPUT_CONTEXT: CellInputInterpretationContext = {
  sourceKind: 'clipboard-text', cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900',
  referenceDate: { year: 2026, month: 8, day: 27, hour: 0, minute: 0, second: 0, millisecond: 0 },
};
const DIRECT_INPUT_CONTEXT: CellInputInterpretationContext = { ...TEST_INPUT_CONTEXT, sourceKind: 'direct-entry' };
const parseTsv = (text: string) => parseTsvWithContext(text, TEST_INPUT_CONTEXT);
const parseClipboardPayload = (payload: Parameters<typeof parseClipboardPayloadWithContext>[0]) => parseClipboardPayloadWithContext(payload, TEST_INPUT_CONTEXT);

test('sheet commands: cell.set, range.set, and undo/redo', () => {
  const workbook = new WorkbookModel('unit-sheet-cmd', 'Commands');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);

  // 1. Single cell set
  runtime.execute('sheet.cell.set', {
    sheetId: 'sheet-1',
    row: 0,
    column: 0,
    value: { value: 'Title', style: { bold: true } },
  });
  const sheet = workbook.getSheet('sheet-1');
  assert.equal(sheet.cells.get(0, 0)?.value, 'Title');
  assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);

  // 2. Undo
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0), undefined);

  // 3. Redo
  runtime.redo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'Title');

  // 4. Batch range set
  runtime.execute('sheet.range.set', {
    sheetId: 'sheet-1',
    startRow: 1,
    startColumn: 0,
    values: [
      [{ value: 10 }, { value: 20 }],
      [{ value: 30 }, { value: 40 }],
    ],
  });
  assert.equal(sheet.cells.get(1, 0)?.value, 10);
  assert.equal(sheet.cells.get(1, 1)?.value, 20);
  assert.equal(sheet.cells.get(2, 0)?.value, 30);
  assert.equal(sheet.cells.get(2, 1)?.value, 40);

  runtime.undo();
  assert.equal(sheet.cells.get(1, 0), undefined);
  assert.equal(sheet.cells.get(2, 1), undefined);
  assert.equal(sheet.cells.get(0, 0)?.value, 'Title'); // untouched
});

test('cell.set authority matches canonical values independently of JSON object key order', () => {
  const workbook = new WorkbookModel('unit-cell-write-authority', 'Cell authority');
  const sheet = workbook.getSheet('sheet-1');
  const params = createCellSetMutationParams(sheet, {
    sheetId: sheet.id,
    row: 0,
    column: 0,
    value: { value: 'ordered', style: { bold: true, italic: false } },
  }, 'external-sync');
  const reordered = {
    column: params.column,
    row: params.row,
    sheetId: params.sheetId,
    value: { style: { italic: false, bold: true }, value: 'ordered' },
    writeAuthority: {
      validationDecision: params.writeAuthority.validationDecision,
      candidate: { style: { italic: false, bold: true }, value: 'ordered' },
      target: params.writeAuthority.target,
      kind: params.writeAuthority.kind,
    },
  };
  assert.equal(isCellSetMutationParams(reordered), true);
});

test('checkbox editor normalizes supported values atomically and restores exact cells', () => {
  const workbook = new WorkbookModel('unit-checkbox-editor', 'Checkbox Editor');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: null, style: { bold: true } });
  sheet.cells.set(0, 1, { value: true });
  sheet.cells.set(0, 2, { value: false });
  sheet.cells.set(0, 3, { value: 0 });
  sheet.cells.set(0, 4, { value: 1 });
  sheet.cells.set(0, 5, { value: ' TRUE ' });
  sheet.cells.set(0, 6, { value: 'FALSE' });
  const before = structuredClone([...Array.from({ length: 7 }, (_, column) => sheet.cells.get(0, column))]);

  runtime.execute('sheet.cellEditor.set', {
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 6 }],
    editor: { kind: 'checkbox' },
  });

  assert.deepEqual(
    Array.from({ length: 7 }, (_, column) => sheet.cells.get(0, column)?.value),
    [false, true, false, false, true, true, false],
  );
  assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);
  assert.equal(runtime.getHistoryDepth().undo, 1);

  runtime.undo();
  assert.deepEqual(Array.from({ length: 7 }, (_, column) => sheet.cells.get(0, column)), before);
  runtime.redo();
  assert.equal(sheet.cells.get(0, 4)?.value, true);

  runtime.execute('sheet.cellEditor.set', {
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 6 }],
    editor: undefined,
  });
  assert.equal(sheet.cells.get(0, 0)?.editor, undefined);
  assert.equal(sheet.cells.get(0, 0)?.value, false);
});

test('checkbox editor rejects unsupported and formula values without partial mutation', () => {
  const workbook = new WorkbookModel('unit-checkbox-reject', 'Checkbox Reject');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(1, 0, { value: 'not-a-boolean' });
  sheet.cells.set(1, 1, { value: 2 });
  sheet.cells.set(1, 2, { formula: '=1', value: null });
  const before = structuredClone([
    sheet.cells.get(1, 0),
    sheet.cells.get(1, 1),
    sheet.cells.get(1, 2),
  ]);

  assert.throws(() => runtime.execute('sheet.cellEditor.set', {
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 2 }],
    editor: { kind: 'checkbox' },
  }), /Checkbox source value|formula cell/);
  assert.deepEqual([
    sheet.cells.get(1, 0),
    sheet.cells.get(1, 1),
    sheet.cells.get(1, 2),
  ], before);
  assert.equal(runtime.getHistoryDepth().undo, 0);
});

test('checkbox.toggle flips canonical ranges as one undoable operation and rejects mixed selections', () => {
  const workbook = new WorkbookModel('unit-checkbox-toggle', 'Checkbox Toggle');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(2, 0, { value: false, editor: { kind: 'checkbox' } });
  sheet.cells.set(2, 1, { value: true, editor: { kind: 'checkbox' } });
  sheet.cells.set(3, 0, { value: true, editor: { kind: 'checkbox' } });
  sheet.cells.set(3, 1, { value: false, editor: { kind: 'checkbox' } });

  const result = runtime.execute('checkbox.toggle', {
    sheetId: sheet.id,
    ranges: [
      { sheetId: sheet.id, startRow: 2, endRow: 3, startColumn: 0, endColumn: 1 },
      { sheetId: sheet.id, startRow: 3, endRow: 3, startColumn: 1, endColumn: 1 },
    ],
  });
  assert.equal(result.mutationCount, 4);
  assert.deepEqual(
    Array.from({ length: 4 }, (_, index) => sheet.cells.get(2 + Math.floor(index / 2), index % 2)?.value),
    [true, false, false, true],
  );
  assert.equal(runtime.getHistoryDepth().undo, 1);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, false);
  assert.equal(sheet.cells.get(3, 1)?.value, false);
  runtime.redo();
  assert.equal(sheet.cells.get(2, 0)?.value, true);

  const before = structuredClone(sheet.cells.get(2, 0));
  sheet.cells.set(2, 2, { value: 'text' });
  assert.throws(() => runtime.execute('checkbox.toggle', {
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 2 }],
  }), /canonical Boolean checkbox/);
  assert.deepEqual(sheet.cells.get(2, 0), before);
});

test('checkbox cell text commits stay Boolean and reject unsupported input', () => {
  const workbook = new WorkbookModel('unit-checkbox-commit', 'Checkbox Commit');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: false, editor: { kind: 'checkbox' } });

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: 'TRUE', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 0)?.value, true);
  assert.throws(() => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: 'maybe', inputContext: DIRECT_INPUT_CONTEXT }), /Checkbox source value/);
  assert.equal(sheet.cells.get(0, 0)?.value, true);
});

test('TableSheet designer updates the canonical definition and rejects unknown fields', () => {
  const workbook = new WorkbookModel('unit-table-sheet-designer', 'TableSheet Designer');
  const table = {
    id: 'table-1',
    name: 'Orders',
    sourceSheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 },
    rowCount: 4,
    fields: [
      { id: 'order', name: 'Order', ordinal: 0, type: 'text' as const },
      { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' as const },
      { id: 'region', name: 'Region', ordinal: 2, type: 'text' as const },
    ],
    blockSize: 4096,
    blocks: [],
    revision: 0,
  };
  workbook.addTable(table);
  const sheet = workbook.addAdvancedSheet({ id: 'table-sheet-1', name: 'Orders view', kind: 'table-sheet', tableSheet: { viewId: table.id, columns: table.fields.map((field) => ({ fieldId: field.id, caption: field.name, type: field.type })), grouping: [] } });
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const next = { viewId: table.id, columns: [
    { fieldId: 'region', caption: 'Region', type: 'text' as const, widthPx: 180 },
    { fieldId: 'amount', caption: 'Amount', type: 'number' as const },
  ], grouping: [{ fieldId: 'region', collapsed: false }], sortState: [{ fieldId: 'amount', direction: 'desc' as const }] };
  runtime.execute('tableSheet.update', { sheetId: sheet.id, definition: next });
  assert.deepEqual(sheet.tableSheet, next);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.undo();
  assert.equal(sheet.tableSheet?.columns.length, 3);
  runtime.redo();
  assert.equal(sheet.tableSheet?.columns[0]?.fieldId, 'region');
  const before = structuredClone(sheet.tableSheet);
  assert.throws(() => runtime.execute('tableSheet.update', { sheetId: sheet.id, definition: { ...next, columns: [{ fieldId: 'missing', caption: 'Missing' }], grouping: [] } }), /definition is invalid|binding table|fields/i);
  assert.deepEqual(sheet.tableSheet, before);
});

test('GanttSheet designer updates canonically, supports undo/redo, and rejects unmapped fields', () => {
  const workbook = new WorkbookModel('unit-gantt-sheet-designer', 'GanttSheet Designer');
  const table = {
    id: 'gantt-table-1', name: 'Tasks', sourceSheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 6 }, rowCount: 3,
    fields: [
      { id: 'id', name: 'ID', ordinal: 0, type: 'text' as const },
      { id: 'title', name: 'Title', ordinal: 1, type: 'text' as const },
      { id: 'start', name: 'Start', ordinal: 2, type: 'date' as const },
      { id: 'end', name: 'End', ordinal: 3, type: 'date' as const },
      { id: 'progress', name: 'Progress', ordinal: 4, type: 'number' as const },
      { id: 'parent', name: 'Parent', ordinal: 5, type: 'text' as const },
      { id: 'deps', name: 'Dependencies', ordinal: 6, type: 'text' as const },
    ], blockSize: 4096, blocks: [], revision: 0,
  };
  workbook.addTable(table);
  const definition = { viewId: table.id, fieldMap: { id: 'id', title: 'title', start: 'start', end: 'end', progress: 'progress', parentId: 'parent', dependencies: 'deps' }, calendar: { workingDays: [1, 2, 3, 4, 5], dayStartHour: 9, dayEndHour: 18 }, timeline: { unit: 'week' as const }, dependencyStyle: { color: '#64748b', width: 1 } };
  const sheet = workbook.addAdvancedSheet({ id: 'gantt-sheet-1', name: 'Tasks view', kind: 'gantt-sheet', ganttSheet: definition });
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  runtime.execute('ganttSheet.update', { sheetId: sheet.id, definition: { ...definition, timeline: { unit: 'day' as const } } });
  assert.equal(sheet.ganttSheet?.timeline.unit, 'day');
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.undo();
  assert.equal(sheet.ganttSheet?.timeline.unit, 'week');
  runtime.redo();
  assert.equal(sheet.ganttSheet?.timeline.unit, 'day');
  const before = structuredClone(sheet.ganttSheet);
  assert.throws(() => runtime.execute('ganttSheet.update', { sheetId: sheet.id, definition: { ...definition, fieldMap: { ...definition.fieldMap, start: 'missing' } } }), /field mapping|unavailable|invalid/i);
  assert.deepEqual(sheet.ganttSheet, before);
});

test('ReportSheet designer updates canonical bindings and render settings', () => {
  const workbook = new WorkbookModel('unit-report-sheet-designer', 'ReportSheet Designer');
  const table = { id: 'report-table-1', name: 'Rows', sourceSheetId: 'sheet-1', sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }, rowCount: 2, fields: [{ id: 'title', name: 'Title', ordinal: 0, type: 'text' as const }, { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' as const }], blockSize: 4096, blocks: [], revision: 0 };
  workbook.addTable(table);
  const definition = { templateSheetId: 'sheet-1', tableId: table.id, bindings: [], pagination: { enabled: true, rowsPerPage: 50, repeatHeaderRows: [0] }, renderMode: 'design' as const, layout: { orientation: 'portrait' as const, marginTopPx: 24, marginRightPx: 24, marginBottomPx: 24, marginLeftPx: 24 }, dataEntry: [] };
  const sheet = workbook.addAdvancedSheet({ id: 'report-sheet-1', name: 'Report view', kind: 'report-sheet', reportSheet: definition });
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const next = { ...definition, bindings: [{ cell: { row: 1, column: 0 }, expression: 'title', kind: 'field' as const, direction: 'vertical' as const, fill: 'down' as const }], renderMode: 'preview' as const, dataEntry: [{ fieldId: 'title', writable: true }] };
  runtime.execute('reportSheet.update', { sheetId: sheet.id, definition: next });
  assert.equal(sheet.reportSheet?.renderMode, 'preview');
  assert.equal(sheet.reportSheet?.bindings[0]?.expression, 'title');
  runtime.undo();
  assert.equal(sheet.reportSheet?.renderMode, 'design');
  const before = structuredClone(sheet.reportSheet);
  assert.throws(() => runtime.execute('reportSheet.update', { sheetId: sheet.id, definition: { ...next, bindings: [{ ...next.bindings[0]!, expression: 'missing' }] } }), /binding field|unavailable|invalid/i);
  assert.deepEqual(sheet.reportSheet, before);
});

test('sheet commands: range.clear, style.set, and merges', () => {
  const workbook = new WorkbookModel('unit-sheet-cmd2', 'Commands2');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  // Setup cells
  sheet.cells.set(0, 0, { value: 'Hello', style: { bold: true, background: '#ff0000' } });
  sheet.cells.set(0, 1, { value: 'World', style: { bold: true } });

  // Style batch
  runtime.execute('sheet.style.set', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    style: { background: '#00ff00' },
  });
  assert.equal(sheet.cells.get(0, 0)?.style?.background, '#00ff00');
  assert.equal(sheet.cells.get(0, 1)?.style?.background, '#00ff00');

  // Merges
  runtime.execute('sheet.merge.set', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
  });
  assert.equal(sheet.merges.length, 1);
  assert.equal(sheet.isMerged(0, 0)?.anchor.column, 0);

  runtime.execute('sheet.merge.remove', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
  });
  assert.equal(sheet.merges.length, 0);

  // Range clear
  runtime.execute('sheet.range.clear', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    family: 'all',
  });
  assert.equal(sheet.cells.get(0, 0), undefined);
  assert.equal(sheet.cells.get(0, 1), undefined);
});

test('canonical border command preserves non-border style and is atomic across undo/redo', () => {
  const workbook = new WorkbookModel('unit-border-topology', 'Borders');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  const range = { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 };
  sheet.cells.set(1, 1, { value: 'keep', style: { bold: true, background: '#fef3c7' } });

  runtime.execute('sheet.borders.set', { sheetId: sheet.id, range, placement: 'outside', line: { style: 'thin', color: '#334155' } });
  assert.equal(runtime.getHistoryDepth().undo, 1);
  assert.deepEqual(sheet.cells.get(1, 1)?.style?.borders, {
    top: { style: 'thin', color: '#334155' },
    left: { style: 'thin', color: '#334155' },
  });
  assert.deepEqual(sheet.cells.get(1, 2)?.style?.borders, { top: { style: 'thin', color: '#334155' }, right: { style: 'thin', color: '#334155' } });
  assert.equal(sheet.cells.get(1, 1)?.style?.bold, true);
  assert.equal(sheet.cells.get(1, 1)?.value, 'keep');
  assert.equal(sheet.cells.get(2, 1)?.style?.borders?.bottom?.style, 'thin');
  assert.equal(sheet.cells.get(2, 2)?.style?.borders?.right?.style, 'thin');

  runtime.undo();
  assert.equal(sheet.cells.get(1, 1)?.style?.borders, undefined);
  assert.equal(sheet.cells.get(1, 1)?.style?.bold, true);
  assert.equal(sheet.cells.get(1, 1)?.value, 'keep');
  runtime.redo();
  assert.equal(sheet.cells.get(1, 1)?.style?.borders?.top?.style, 'thin');

  runtime.execute('sheet.borders.set', { sheetId: sheet.id, range, placement: 'none' });
  assert.deepEqual(sheet.cells.get(1, 1)?.style?.borders, {});
  assert.equal(sheet.cells.get(1, 1)?.style?.background, '#fef3c7');
});

test('alignment styles use one canonical mutation across local undo/redo and remote replay', () => {
  const workbook = new WorkbookModel('unit-alignment-style', 'Alignment');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 'Across' });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  const style = {
    horizontalAlignment: 'centerContinuous' as const,
    verticalAlignment: 'distributed' as const,
    shrinkToFit: true,
    indent: 2,
    readingOrder: 'rtl' as const,
    textOrientation: 'stacked' as const,
  };

  runtime.execute('sheet.style.set', { sheetId: sheet.id, range, style });
  assert.deepEqual(sheet.cells.get(0, 0)?.style, style);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.style, undefined);
  runtime.redo();
  assert.deepEqual(sheet.cells.get(0, 0)?.style, style);

  const beforeRejected = structuredClone(sheet.cells.get(0, 0));
  assert.throws(() => runtime.applyRemoteMutations([{
    id: 'style.set',
    unitId: workbook.unitId,
    sheetId: sheet.id,
    params: { sheetId: sheet.id, range, style: { horizontalAlignment: 'unsafe-native-value' } },
    affectedRanges: [range],
  }]), /Invalid mutation history/);
  assert.deepEqual(sheet.cells.get(0, 0), beforeRejected);

  const replayWorkbook = new WorkbookModel('unit-alignment-style', 'Alignment');
  const replayRuntime = new CommandRuntime(replayWorkbook);
  registerSheetCommands(replayRuntime);
  replayWorkbook.getSheet('sheet-1').cells.set(0, 0, { value: 'Across' });
  replayRuntime.applyRemoteMutations([{
    id: 'style.set',
    unitId: replayWorkbook.unitId,
    sheetId: 'sheet-1',
    params: { sheetId: 'sheet-1', range, style },
    affectedRanges: [range],
  }]);
  assert.deepEqual(replayWorkbook.getSheet('sheet-1').cells.get(0, 0)?.style, style);
});

test('sheet commands: sort and canonical fill', () => {
  const workbook = new WorkbookModel('unit-sort', 'SortAutofill');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  // Populate data for sorting
  sheet.cells.set(0, 0, { value: 'Name' });
  sheet.cells.set(0, 1, { value: 'Score' });
  sheet.cells.set(1, 0, { value: 'Charlie' });
  sheet.cells.set(1, 1, { value: 70 });
  sheet.cells.set(2, 0, { value: 'Alice' });
  sheet.cells.set(2, 1, { value: 95 });
  sheet.cells.set(3, 0, { value: 'Bob' });
  sheet.cells.set(3, 1, { value: 85 });

  // Sort ascending by Score (col 1), with header
  runtime.execute('sheet.sort', {
    sheetId: 'sheet-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    sortColumn: 1,
    ascending: true,
    hasHeader: true,
  });

  assert.equal(sheet.cells.get(0, 0)?.value, 'Name'); // header intact
  assert.equal(sheet.cells.get(1, 0)?.value, 'Charlie'); // 70
  assert.equal(sheet.cells.get(2, 0)?.value, 'Bob');     // 85
  assert.equal(sheet.cells.get(3, 0)?.value, 'Alice');   // 95

  // Canonical fill formula copy
  sheet.cells.set(5, 0, { value: null, formula: '=A1+B1' });
  runtime.execute('sheet.range.fill', {
    sheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0 },
    targetRange: { sheetId: 'sheet-1', startRow: 5, endRow: 7, startColumn: 0, endColumn: 0 },
    direction: 'down',
    mode: 'copy',
  });
  assert.equal(sheet.cells.get(6, 0)?.formula, '=A2+B2');
  assert.equal(sheet.cells.get(7, 0)?.formula, '=A3+B3');
});

test('clipboard: TSV format, parse, and formula shifting', () => {
  const tsv = formatTsv([
    [{ value: 'Product' }, { value: 'Price' }],
    [{ value: 'Keyboard' }, { value: 120 }],
  ]);
  assert.equal(tsv, 'Product\tPrice\nKeyboard\t120');

  const parsed = parseTsv('Product\tPrice\nKeyboard\t120\nTotal\t=SUM(B2:B2)');
  assert.equal(parsed.length, 3);
  assert.equal(parsed[0]?.[0]?.value, 'Product');
  assert.equal(parsed[1]?.[1]?.value, 120);
  assert.equal(parsed[2]?.[1]?.formula, '=SUM(B2:B2)');

  assert.equal(shiftFormula('=A1+$B$1+C$1+$D1', 2, 3), '=D3+$B$1+F$1+$D3');
  assert.equal(shiftFormula('=SUM(A2#)', 1, 1), '=SUM(B3#)');
  assert.throws(() => shiftFormula('=A1+', 1, 1), FormulaRelocationError);
});

test('clipboard payload carries provenance and PasteSpecialSpec preserves its contracts', () => {
  const workbook = new WorkbookModel('unit-paste-modes', 'Paste Modes');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));
  sheet.cells.set(0, 0, { value: 12, formula: '=A2', style: { bold: true }, numberFormat: '0.00' });
  sheet.cells.set(0, 1, { value: 'keep', style: { italic: true }, numberFormat: '@' });

  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0,
  });
  assert.equal(payload.source, 'internal');
  assert.equal(payload.schema, 'SparseClipboardPayload');
  assert.equal(payload.occupiedCells.length, 1);

  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 0, column: 1 },
    clipboard: payload,
    transfer: 'copy',
    spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }),
  });
  assert.deepEqual(sheet.cells.get(0, 1), { value: 12 });

  sheet.cells.set(0, 1, { value: 'keep', formula: '=A1', style: { italic: true }, numberFormat: '@' });
  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 0, column: 1 },
    clipboard: payload,
    transfer: 'copy',
    spec: createPasteSpecialSpec({ content: 'none', formatting: 'source-formatting', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }),
  });
  assert.equal(sheet.cells.get(0, 1)?.value, 'keep');
  assert.equal(sheet.cells.get(0, 1)?.formula, '=A1');
  assert.equal(sheet.cells.get(0, 1)?.style?.bold, true);
  assert.equal(sheet.cells.get(0, 1)?.numberFormat, '0.00');
});

test('clipboard uses quoted TSV and host-neutral HTML representations', () => {
  const text = formatTsv([[{ value: 'a\tb' }, { value: 'line\nnext' }, { value: '"quoted"' }]]);
  assert.deepEqual(parseTsv(text).map((row) => row.map((cell) => cell.value)), [['a\tb', 'line\nnext', '"quoted"']]);
  const payload = {
    schema: 'SparseClipboardPayload' as const,
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    sourceExtent: { rows: 1, columns: 1 },
    occupiedCells: [],
    transfer: 'copy' as const,
    rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
    representations: [{ mime: 'text/html', data: '<table><tr><td>42</td><td data-formula="=A1+1">43</td></tr></table>' }],
  };
  const parsed = parseClipboardPayload(payload);
  assert.deepEqual(parsed.occupiedCells.map((cell) => cell.value.value), [42, null]);
  assert.equal(parsed.occupiedCells[1]?.value.formula, '=A1+1');
  assert.equal(clipboardRepresentations(parsed).some((entry) => entry.mime === 'text/html'), true);
  assert.deepEqual(parseTsv(' 42 \t true ').map((row) => row.map((cell) => cell.value)), [[' 42 ', ' true ']]);
});

test('paste special copies range-owned metadata atomically and restores it once', () => {
  const workbook = new WorkbookModel('unit-paste-metadata', 'Paste Metadata');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const source = workbook.getSheet('sheet-1');
  source.cells.set(0, 0, { value: 'source', style: { bold: true } });
  source.columnWidthsPx[0] = 144;
  source.review.setNote(0, 0, { id: 'n1', author: 'u', text: 'note', createdAt: '2026-01-01', visible: true });
  source.hyperlinks.set('0:0', { id: 'h1', target: { kind: 'url', url: 'https://example.com' } });
  source.review.addThread({ id: 'c1', sheetId: source.id, row: 0, column: 0, author: 'u', text: 'comment', createdAt: '2026-01-01', replies: [] });
  source.dataValidations.push({ id: 'dv1', sheetId: source.id, ranges: [{ sheetId: source.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], type: 'whole', formula1: '1' });
  source.conditionalFormats.push({ id: 'cf1', sheetId: source.id, ranges: [{ sheetId: source.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], type: 'highlight', operator: 'equal', value1: 'source' });
  const payload = copyRangeToClipboardData(workbook, { sheetId: source.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  const target = { sheetId: source.id, row: 2, column: 2 };
  const result = runtime.execute('sheet.range.paste', {
    sheetId: source.id,
    targetOrigin: { row: target.row, column: target.column },
    clipboard: payload,
    transfer: 'copy',
    spec: createPasteSpecialSpec({ metadata: { commentsNotes: true, validation: true, columnWidths: true, conditionalFormats: true, hyperlinks: true } }),
  });
  assert.equal(result.mutationCount, 1);
  assert.equal(source.cells.get(2, 2)?.value, 'source');
  assert.equal(source.review.getNoteAt(2, 2)?.text, 'note');
  assert.equal(source.hyperlinks.get('2:2')?.target.kind, 'url');
  assert.equal(source.review.getThreadsAt(2, 2).length, 1);
  assert.equal(source.dataValidations.some((rule) => rule.ranges.some((range) => range.startRow === 2 && range.startColumn === 2)), true);
  assert.equal(source.conditionalFormats.some((rule) => rule.ranges.some((range) => range.startRow === 2 && range.startColumn === 2)), true);
  assert.equal(source.columnWidthsPx[2], 144);
  runtime.undo();
  assert.equal(source.cells.get(2, 2), undefined);
  assert.equal(source.review.hasNoteAt(2, 2), false);
  assert.equal(source.hyperlinks.has('2:2'), false);
  assert.equal(source.review.getThreadsAt(2, 2).length, 0);
  assert.equal(source.dataValidations.some((rule) => rule.ranges.some((range) => range.startRow === 2 && range.startColumn === 2)), false);
  assert.equal(source.columnWidthsPx[2], undefined);
  runtime.redo();
  assert.equal(source.cells.get(2, 2)?.value, 'source');
  assert.equal(runtime.getHistoryDepth().undo, 1);
});

test('paste special arithmetic, skip blanks and protected rejection are fail-closed', () => {
  const workbook = new WorkbookModel('unit-paste-arithmetic', 'Paste Arithmetic');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 2 });
  sheet.cells.set(1, 0, { value: 3 });
  const payload = copyRangeToClipboardData(workbook, { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 1, column: 0 },
    clipboard: payload,
    transfer: 'copy',
    spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', operation: 'add', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }),
  });
  assert.equal(sheet.cells.get(1, 0)?.value, 5);
  sheet.cells.set(2, 0, { value: 'keep' });
  const blankPayload = copyRangeToClipboardData(workbook, { sheetId: sheet.id, startRow: 3, endRow: 3, startColumn: 0, endColumn: 0 });
  runtime.execute('sheet.range.paste', { sheetId: sheet.id, targetOrigin: { row: 2, column: 0 }, clipboard: blankPayload, transfer: 'copy', spec: createPasteSpecialSpec({ skipBlanks: true }) });
  assert.equal(sheet.cells.get(2, 0)?.value, 'keep');
  sheet.protectionRules.push({ id: 'locked', scope: 'range', sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 4, endRow: 4, startColumn: 0, endColumn: 0 }, locked: true, allow: {} });
  assert.throws(() => runtime.execute('sheet.range.paste', { sheetId: sheet.id, targetOrigin: { row: 4, column: 0 }, clipboard: payload, transfer: 'copy', spec: createPasteSpecialSpec() }), /protected/);
  assert.equal(sheet.cells.get(4, 0), undefined);
  runtime.execute('sheet.range.paste', { sheetId: sheet.id, targetOrigin: { row: 5, column: 0 }, clipboard: payload, transfer: 'copy', spec: createPasteSpecialSpec({ content: 'none', formatting: 'none', link: true, metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) });
  assert.equal(sheet.cells.get(5, 0)?.formula, "='Sheet1'!A1");
});

test('sheet.cell.commitText is the single raw-text input path', () => {
  const workbook = new WorkbookModel('unit-commit-text', 'Commit Text');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));
  sheet.cells.set(0, 0, {
    value: 'old',
    style: { bold: true, background: '#fff' },
    styleId: 'style-1',
    numberFormat: '@',
    displayValue: 'old',
  });

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: ' 42 ', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 0)?.value, ' 42 ');
  assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);
  assert.equal(sheet.cells.get(0, 0)?.styleId, 'style-1');
  assert.equal(sheet.cells.get(0, 0)?.displayValue, undefined);

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: '=A1 + 1', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1 + 1');
  assert.equal(sheet.cells.get(0, 0)?.value, null);
  assert.equal(sheet.cells.get(0, 0)?.style?.background, '#fff');

  assert.throws(
    () => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: '=A1 +', inputContext: DIRECT_INPUT_CONTEXT }),
    /formula|expression|unexpected/i,
  );
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1 + 1');

  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, ' 42 ');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'old');
  assert.equal(sheet.cells.get(0, 0)?.formula, undefined);
});

test('sheet.cell.commitText parses scalars, validates input and protects spill children', () => {
  const workbook = new WorkbookModel('unit-commit-text-guards', 'Commit Guards');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));

  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 0, text: '42', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 0)?.value, 42);
  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 1, text: 'TRUE', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 1)?.value, true);
  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 0, column: 2, text: ' true ', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(0, 2)?.value, ' true ');

  sheet.dataValidations.push({
    id: 'list-commit',
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }],
    type: 'list',
    listSource: { kind: 'values', values: ['Allowed'] },
    alertStyle: 'stop',
  });
  assert.throws(
    () => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 1, column: 0, text: 'Rejected', inputContext: DIRECT_INPUT_CONTEXT }),
    /允许的列表|validation/i,
  );
  assert.equal(sheet.cells.get(1, 0), undefined);

  sheet.spillRanges.push({
    sheetId: sheet.id,
    anchor: { row: 2, column: 0 },
    range: { sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 1 },
    values: [[1, 2]],
    state: 'ok',
  });
  assert.throws(
    () => runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 2, column: 1, text: 'blocked', inputContext: DIRECT_INPUT_CONTEXT }),
    /spill child/i,
  );
  runtime.execute('sheet.cell.commitText', { sheetId: sheet.id, row: 2, column: 0, text: 'anchor', inputContext: DIRECT_INPUT_CONTEXT });
  assert.equal(sheet.cells.get(2, 0)?.value, 'anchor');
});

test('cut paste is one cross-sheet transaction and preserves formula references', () => {
  const workbook = new WorkbookModel('unit-cut', 'Cut');
  const target = workbook.addSheet('sheet-2', 'Target');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const source = workbook.getSheet('sheet-1');
  source.cells.set(1, 1, { value: null, formula: '=A1+$B$1' });
  target.cells.set(4, 4, { value: 'old' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: source.id,
    startRow: 1,
    endRow: 1,
    startColumn: 1,
    endColumn: 1,
  });
  payload.transfer = 'move';
  const result = runtime.execute('sheet.range.paste', {
    sheetId: target.id,
    targetOrigin: { row: 4, column: 4 },
    clipboard: payload,
    transfer: 'move',
    spec: createPasteSpecialSpec(),
  });
  assert.equal(result.mutationCount, 1);
  assert.equal(source.cells.get(1, 1), undefined);
  assert.equal(target.cells.get(4, 4)?.formula, '=A1+$B$1');
  assert.equal(target.cells.get(4, 4)?.value, null);
  runtime.undo();
  assert.equal(source.cells.get(1, 1)?.formula, '=A1+$B$1');
  assert.equal(target.cells.get(4, 4)?.value, 'old');
  runtime.redo();
  assert.equal(source.cells.get(1, 1), undefined);
  assert.equal(target.cells.get(4, 4)?.formula, '=A1+$B$1');
});

test('copy paste shifts relative references while preserving mixed and absolute references', () => {
  const workbook = new WorkbookModel('unit-copy', 'Copy');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(1, 1, { value: null, formula: '=A1+$B$1+C$1+$D1' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 1,
    endRow: 1,
    startColumn: 1,
    endColumn: 1,
  });
  runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 4, column: 4 },
    clipboard: payload,
    transfer: 'copy',
    spec: createPasteSpecialSpec(),
  });
  assert.equal(sheet.cells.get(4, 4)?.formula, '=D4+$B$1+F$1+$D4');
  assert.equal(sheet.cells.get(1, 1)?.formula, '=A1+$B$1+C$1+$D1');
});

test('paste replay rejects a transfer mismatch before touching the workbook', () => {
  const workbook = new WorkbookModel('unit-paste-contract', 'Paste Contract');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: null, formula: '=A1' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0,
  });
  payload.transfer = 'move';
  assert.throws(() => runtime.applyRemoteMutations([{
    id: 'range.paste',
    unitId: workbook.unitId,
    sheetId: sheet.id,
    params: {
      sheetId: sheet.id,
      targetOrigin: { row: 1, column: 1 },
      clipboard: payload,
      transfer: 'copy',
      spec: createPasteSpecialSpec(),
      startRow: 1,
      startColumn: 1,
      sourceRange: structuredClone(payload.range),
      clearSource: true,
    },
    affectedRanges: [],
  }]), /Invalid mutation history/);
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1');
  assert.equal(sheet.cells.get(1, 1), undefined);
});

test('copy paste rejects a malformed formula before creating a mutation', () => {
  const workbook = new WorkbookModel('unit-paste-fail-close', 'Paste Fail Close');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: null, formula: '=A1+' });
  sheet.cells.set(1, 1, { value: 'unchanged' });
  const payload = copyRangeToClipboardData(workbook, {
    sheetId: sheet.id,
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0,
  });
  assert.throws(() => runtime.execute('sheet.range.paste', {
    sheetId: sheet.id,
    targetOrigin: { row: 1, column: 1 },
    clipboard: payload,
    transfer: 'copy',
    spec: createPasteSpecialSpec(),
  }), FormulaRelocationError);
  assert.equal(sheet.cells.get(1, 1)?.value, 'unchanged');
  assert.equal(runtime.getHistoryDepth().undo, 0);
});

test('range clear modes are independent and restore auxiliary metadata', () => {
  const workbook = new WorkbookModel('unit-clear-modes', 'Clear Modes');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, {
    value: 10,
    formula: '=A1',
    style: { bold: true },
    hyperlink: 'https://example.com',
  });
  sheet.review.setNote(0, 0, { id: 'note', author: 'u', text: 'standalone', createdAt: '2026-01-01', visible: true });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, family: 'formats' });
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
  assert.equal(sheet.cells.get(0, 0)?.formula, '=A1');
  assert.equal(sheet.cells.get(0, 0)?.style, undefined);
  assert.equal(sheet.cells.get(0, 0)?.hyperlink, 'https://example.com');
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, family: 'comments-and-notes' });
  assert.equal(sheet.review.hasNoteAt(0, 0), false);
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, family: 'hyperlinks' });
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
  assert.equal(sheet.cells.get(0, 0)?.hyperlink, undefined);
  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range, family: 'all' });
  assert.equal(sheet.cells.get(0, 0), undefined);
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 10);
});

test('clear formats/all crop conditional-format intersections and restore atomically', () => {
  const workbook = new WorkbookModel('unit-clear-conditional-format', 'Clear Conditional Format');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  const rule = {
    id: 'cf-clear-1',
    sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 }],
    type: 'highlight' as const,
    style: { background: '#ffeeaa' },
  };
  sheet.conditionalFormats.push(rule);
  sheet.cells.set(2, 2, { value: 'preserve', style: { bold: true } });
  const clearRange = { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 };

  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range: clearRange, family: 'formats' });
  assert.equal(sheet.cells.get(2, 2)?.value, 'preserve');
  assert.equal(sheet.cells.get(2, 2)?.style, undefined);
  assert.deepEqual(sheet.conditionalFormats[0]?.ranges, [
    { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 },
    { sheetId: sheet.id, startRow: 4, endRow: 4, startColumn: 0, endColumn: 4 },
    { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 },
    { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 4, endColumn: 4 },
  ]);

  runtime.undo();
  assert.deepEqual(sheet.conditionalFormats, [rule]);
  assert.deepEqual(sheet.cells.get(2, 2)?.style, { bold: true });

  runtime.execute('sheet.range.clear', { sheetId: sheet.id, range: clearRange, family: 'all' });
  assert.equal(sheet.cells.get(2, 2), undefined);
  assert.equal(sheet.conditionalFormats.length, 1);
  runtime.undo();
  assert.equal(sheet.cells.get(2, 2)?.value, 'preserve');
  assert.deepEqual(sheet.conditionalFormats, [rule]);
});

function sheetId(workbook: WorkbookModel): string {
  return workbook.primarySheetId;
}

test('sheet commands: hide and unhide rows and columns are undoable commands', () => {
  const workbook = new WorkbookModel('unit-hidden', 'Hidden');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  runtime.execute('sheet.row.hide', { sheetId: sheet.id, index: 2 });
  runtime.execute('sheet.column.hide', { sheetId: sheet.id, index: 3 });
  assert.equal(sheet.hiddenRows.has(2), true);
  assert.equal(sheet.hiddenColumns.has(3), true);

  runtime.undo();
  assert.equal(sheet.hiddenColumns.has(3), false);
  runtime.undo();
  assert.equal(sheet.hiddenRows.has(2), false);

  runtime.redo();
  runtime.redo();
  assert.equal(sheet.hiddenRows.has(2), true);
  assert.equal(sheet.hiddenColumns.has(3), true);

  runtime.execute('sheet.rows.unhide.all', { sheetId: sheet.id });
  runtime.execute('sheet.columns.unhide.all', { sheetId: sheet.id });
  assert.equal(sheet.hiddenRows.size, 0);
  assert.equal(sheet.hiddenColumns.size, 0);
  runtime.undo();
  assert.equal(sheet.hiddenColumns.has(3), true);
  runtime.undo();
  assert.equal(sheet.hiddenRows.has(2), true);
});

test('sheet rows visibility is one atomic selection-aware mutation and replays remotely', () => {
  const workbook = new WorkbookModel('unit-rows-visibility', 'Rows visibility');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  const initial = workbook.snapshot();

  runtime.execute('sheet.rows.visibility.set', { sheetId: sheet.id, rows: [1, 2, 2, 99999], hidden: true });
  assert.deepEqual([...sheet.hiddenRows].sort((left, right) => left - right), [1, 2]);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.undo();
  assert.deepEqual([...sheet.hiddenRows], []);
  runtime.redo();
  assert.deepEqual([...sheet.hiddenRows].sort((left, right) => left - right), [1, 2]);

  const remote = WorkbookModel.fromSnapshot(initial);
  const remoteRuntime = new CommandRuntime(remote);
  registerSheetCommands(remoteRuntime);
  remoteRuntime.applyRemoteMutations(runtime.getUndoEntries()[0]!.redo);
  assert.deepEqual([...remote.getSheet('sheet-1').hiddenRows].sort((left, right) => left - right), [1, 2]);
});

test('sheet commands: row insert/delete use StructuralTransform and preserve undo', () => {
  const workbook = new WorkbookModel('unit-structural', 'Structural');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');

  sheet.cells.set(2, 0, { value: 42, formula: '=A1+1' });
  sheet.merges.push({
    range: { sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 1 },
    anchor: { row: 2, column: 0 },
  });
  sheet.pane = { kind: 'frozen', xSplit: 0, ySplit: 1, startRow: 0, startColumn: 0, state: 'frozen' };

  runtime.execute('sheet.rows.insert', { sheetId: sheet.id, at: 1, count: 2 });
  assert.equal(sheet.rowCount, 1002);
  assert.equal(sheet.cells.get(4, 0)?.value, 42);
  assert.equal(sheet.cells.get(4, 0)?.formula, '=A1+1');
  assert.equal(sheet.merges[0]?.range.startRow, 4);
  assert.equal(sheet.pane.kind === 'frozen' ? sheet.pane.ySplit : 0, 3);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, 42);
  assert.equal(sheet.merges[0]?.range.startRow, 2);

  runtime.execute('sheet.rows.delete', { sheetId: sheet.id, at: 2, count: 1 });
  assert.equal(sheet.cells.get(2, 0), undefined);

  runtime.undo();
  assert.equal(sheet.cells.get(2, 0)?.value, 42);
});

test('sheet.extent.grow expands sparse bounds without polluting edit history and rejects unsupported limits', () => {
  const workbook = new WorkbookModel('unit-extent-grow', 'Extent growth');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  const mutations: string[] = [];
  runtime.onMutation((mutation) => mutations.push(mutation.id));

  runtime.execute('sheet.extent.grow', { sheetId: sheet.id, rowCount: 2_000, columnCount: 52 });

  assert.equal(sheet.rowCount, 2_000);
  assert.equal(sheet.columnCount, 52);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
  assert.deepEqual(mutations, ['sheet.extent.grow']);

  assert.throws(
    () => runtime.execute('sheet.extent.grow', { sheetId: sheet.id, rowCount: 1_048_577, columnCount: 52 }),
    /Invalid sheet\.extent\.grow command payload/,
  );
  assert.equal(sheet.rowCount, 2_000);
  assert.equal(sheet.columnCount, 52);
});

test('sheet.cells.insert undo restores the complete affected band', () => {
  const workbook = new WorkbookModel('unit-shift-undo', 'Shift Undo');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet(sheetId(workbook));
  sheet.cells.set(0, 0, { value: 'top' });
  sheet.cells.set(1, 0, { value: 'bottom' });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };

  runtime.execute('sheet.cells.insert', { sheetId: sheet.id, range, operation: 'insert', axis: 'row' });
  assert.equal(sheet.cells.get(0, 0), undefined);
  assert.equal(sheet.cells.get(1, 0)?.value, 'top');
  assert.equal(sheet.cells.get(2, 0)?.value, 'bottom');
  runtime.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'top');
  assert.equal(sheet.cells.get(1, 0)?.value, 'bottom');
});
