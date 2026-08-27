import assert from 'node:assert/strict';
import test from 'node:test';
import { ColumnDimensionController } from './column-dimension-controller';
import { createAutoFitBlock } from './column-autofit-protocol';
import type { CanvasSheetSnapshot, SelectionState, WorkbookSession } from '@react-sheets/spreadsheet-app';

type TestCell = { value: string; displayValue?: string; formula?: string; style?: { padding?: number } };

function measurementContext(): CanvasRenderingContext2D {
  return {
    save() {},
    restore() {},
    measureText(text: string) { return { width: text.length * 10 } as TextMetrics; },
    font: '',
  } as unknown as CanvasRenderingContext2D;
}

function makeSheet(cells: Record<string, TestCell>, rows = 2, columns = 4): CanvasSheetSnapshot {
  const byAddress = new Map(Object.entries(cells));
  return {
    id: 'sheet-1',
    name: 'Sheet1',
    columns: Array.from({ length: columns }, (_, column) => String.fromCharCode(65 + column)),
    columnCount: columns,
    rowCount: rows,
    occupiedCellCount: byAddress.size,
    getCell: (row: number, column: number) => byAddress.get(`${row}:${column}`),
    usedRange: { sheetId: 'sheet-1', startRow: 0, endRow: rows - 1, startColumn: 0, endColumn: columns - 1 },
    drawings: [],
    drawingPayloads: new Map(),
    pivots: [],
    pivotResults: {},
    pivotProjections: {},
    sparklines: [],
    conditionalFormats: [],
    dataValidations: [],
    merges: [],
    pane: { kind: 'none' },
    getFilterOwner: () => undefined,
    getActiveAutoFilter: () => undefined,
    defaultRowHeightPx: 20,
    defaultColumnWidthPx: 64,
    maximumDigitWidthPx: 7,
    rowHeightsPx: {},
    columnWidthsPx: {},
    hiddenRows: [],
    hiddenColumns: [],
    outlineGroups: [],
    outlineControls: [],
    filterRangeColumns: [],
    activeFilterColumns: [],
    filterButtons: [],
    filterButtonStates: [],
    getFilterValueDomain: () => [],
    getFilterDomainDescriptor: () => ({ families: [] }) as never,
    getFilterCriterion: () => undefined,
    getFilterColorDomain: () => [],
    getFilterIconDomain: () => [],
    sheetTables: [],
    forEachOccupiedCell: (visitor: (row: number, column: number) => void) => Object.keys(cells).forEach((key) => {
      const [row, column] = key.split(':').map(Number);
      visitor(row!, column!);
    }),
  } as unknown as CanvasSheetSnapshot;
}

function installCanvas(): { restore: () => void } {
  const previousDocument = (globalThis as { document?: unknown }).document;
  const context = measurementContext();
  (globalThis as { document?: unknown }).document = {
    createElement: () => ({ getContext: () => context }),
  };
  return { restore: () => { (globalThis as { document?: unknown }).document = previousDocument; } };
}

function controllerFor(sheet: CanvasSheetSnapshot, calls: { columns?: unknown[]; rows?: unknown[] }, selection: SelectionState = { ranges: [], activeCell: { row: 0, column: 0 }, anchorCell: { row: 0, column: 0 }, primaryRangeIndex: 0 }): ColumnDimensionController {
  return new ColumnDimensionController(
    {
      applyColumnWidths: (entries: readonly unknown[]) => { calls.columns = [...entries]; },
      applyRowHeights: (entries: readonly unknown[]) => { calls.rows = [...entries]; },
    } as unknown as WorkbookSession,
    () => sheet,
    () => selection as never,
  );
}

test('selectedRows uses full-row and ordinary cell selections through one deterministic rule', () => {
  const sheet = makeSheet({}, 10, 4);
  const controller = controllerFor(sheet, {}, {
    ranges: [
      { sheetId: sheet.id, startRow: 2, endRow: 3, startColumn: 0, endColumn: 3 },
      { sheetId: sheet.id, startRow: 7, endRow: 8, startColumn: 1, endColumn: 2 },
    ],
    primaryRangeIndex: 0,
    activeCell: { row: 2, column: 0 },
    anchorCell: { row: 2, column: 0 },
  });
  assert.deepEqual(controller.selectedRows(), [2, 3, 7, 8]);
  assert.deepEqual(controller.selectedRows(false), [2, 3]);
  assert.deepEqual(controller.rowsForBoundary(3), [2, 3]);
  assert.deepEqual(controller.rowsForBoundary(5), [5]);
});

test('AutoFit rows and main-thread columns measure canonical 0, FALSE, formatted text and errors', async () => {
  const canvas = installCanvas();
  try {
    const sheet = makeSheet({
      '0:0': { value: '0', displayValue: '0', formula: '=0' },
      '0:1': { value: 'FALSE', displayValue: 'FALSE', formula: '=FALSE()' },
      '0:2': { value: '#DIV/0!', displayValue: '#DIV/0!', formula: '=1/0' },
      '1:0': { value: '' },
    });
    const calls: { columns?: unknown[]; rows?: unknown[] } = {};
    const controller = controllerFor(sheet, calls);

    await controller.autoFitRows([0, 1]);
    const heights = calls.rows as Array<{ row: number; heightPx: number }>;
    assert.ok((heights.find((entry) => entry.row === 0)?.heightPx ?? 0) > 8);
    assert.equal(heights.find((entry) => entry.row === 1)?.heightPx, 8);

    await controller.autoFit([0, 1, 2, 3]);
    const widths = calls.columns as Array<{ column: number; widthPx: number }>;
    assert.ok((widths.find((entry) => entry.column === 0)?.widthPx ?? 0) > 8);
    assert.ok((widths.find((entry) => entry.column === 1)?.widthPx ?? 0) > 8);
    assert.ok((widths.find((entry) => entry.column === 2)?.widthPx ?? 0) > 8);
    assert.equal(widths.find((entry) => entry.column === 3)?.widthPx, 8);
  } finally {
    canvas.restore();
  }
});

test('AutoFit compact blocks intern matching styles and keep per-cell flags positional', () => {
  const style = { fontSizePx: 14, bold: true };
  const block = createAutoFitBlock([
    { column: 2, value: 'A', style, filterButton: true },
    { column: 4, value: 'B', style },
    { column: 8, value: 'C' },
  ]);
  assert.deepEqual([...block.columns], [2, 4, 8]);
  assert.deepEqual(block.values, ['A', 'B', 'C']);
  assert.deepEqual([...block.styleIndexes], [1, 1, 0]);
  assert.deepEqual([...block.filterButtons], [1, 0, 0]);
  assert.deepEqual(block.styles, [style]);
});

test('AutoFit worker uses the compact block protocol with the same typed-content gate and measurement', async () => {
  const previousSelf = (globalThis as { self?: unknown }).self;
  const previousOffscreenCanvas = (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas;
  const messages: unknown[] = [];
  const context = measurementContext();
  const worker = {
    onmessage: undefined as ((event: MessageEvent) => void) | undefined,
    postMessage: (message: unknown) => messages.push(message),
  };
  class FakeOffscreenCanvas { getContext(): CanvasRenderingContext2D { return context; } }
  (globalThis as { self?: unknown }).self = worker;
  (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = FakeOffscreenCanvas;
  try {
    await import('./column-autofit-worker');
    worker.onmessage!({ data: { kind: 'start', taskId: 'test', columns: [0, 1, 2, 3] } } as MessageEvent);
    worker.onmessage!({ data: {
      kind: 'chunk',
      taskId: 'test',
      block: createAutoFitBlock([
        { column: 0, value: '0' },
        { column: 1, value: 'FALSE' },
        { column: 2, value: '#DIV/0!' },
        { column: 3, value: '' },
      ]),
    } } as MessageEvent);
    worker.onmessage!({ data: { kind: 'finish', taskId: 'test' } } as MessageEvent);
    const complete = messages.find((message) => (message as { kind?: string }).kind === 'complete') as { widths: Array<{ column: number; widthPx: number }> };
    assert.ok(complete);
    assert.ok((complete.widths.find((entry) => entry.column === 0)?.widthPx ?? 0) > 8);
    assert.ok((complete.widths.find((entry) => entry.column === 1)?.widthPx ?? 0) > 8);
    assert.ok((complete.widths.find((entry) => entry.column === 2)?.widthPx ?? 0) > 8);
    assert.equal(complete.widths.find((entry) => entry.column === 3)?.widthPx, 8);
  } finally {
    (globalThis as { self?: unknown }).self = previousSelf;
    (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = previousOffscreenCanvas;
  }
});
