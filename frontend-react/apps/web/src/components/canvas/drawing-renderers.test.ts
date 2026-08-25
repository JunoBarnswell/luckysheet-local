import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { RangeRef } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import { createCanvasFloatingDrawables, resolveCameraSourceGeometry } from './drawing-renderers';
import type { DrawingObject, DrawingPayload } from '@react-sheets/core-model';
import type { SheetSkeleton } from '@react-sheets/render-engine';

function sourceSnapshot(): CanvasSheetSnapshot {
  return {
    id: 'sheet-1',
    name: 'Source',
    columns: [],
    columnCount: 5,
    rowCount: 6,
    occupiedCellCount: 0,
    getCell: () => undefined,
    usedRange: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
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
    defaultColumnWidthPx: 50,
    maximumDigitWidthPx: 7,
    rowHeightsPx: { 1: 40 },
    columnWidthsPx: { 2: 90 },
    hiddenRows: [3],
    hiddenColumns: [1],
    outlineGroups: [],
    outlineControls: [],
    filterRangeColumns: [],
    activeFilterColumns: [],
    filterButtons: [],
    filterButtonStates: [],
    getFilterValueDomain: () => [],
    getFilterCriterion: () => undefined,
    getFilterColorDomain: () => [],
    getFilterIconDomain: () => [],
    sheetTables: [],
    previewRows: [],
  };
}

test('Camera geometry uses source row heights, column widths, and visibility projection', () => {
  const range: RangeRef = { sheetId: 'sheet-1', startRow: 1, endRow: 4, startColumn: 0, endColumn: 3 };
  assert.deepEqual(resolveCameraSourceGeometry(sourceSnapshot(), range), {
    left: 0,
    top: 20,
    width: 190,
    height: 80,
    firstRow: 1,
    lastRow: 4,
    firstColumn: 0,
    lastColumn: 3,
  });
});

test('Camera geometry rejects invalid, cross-sheet, and fully hidden source ranges', () => {
  const source = sourceSnapshot();
  assert.equal(resolveCameraSourceGeometry(source, { sheetId: 'other', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }), null);
  assert.equal(resolveCameraSourceGeometry(source, { sheetId: 'sheet-1', startRow: -1, endRow: 1, startColumn: 0, endColumn: 1 }), null);
  assert.equal(resolveCameraSourceGeometry(source, { sheetId: 'sheet-1', startRow: 0, endRow: 6, startColumn: 0, endColumn: 1 }), null);
  assert.equal(resolveCameraSourceGeometry(source, { sheetId: 'sheet-1', startRow: 3, endRow: 3, startColumn: 0, endColumn: 1 }), null);
  assert.equal(resolveCameraSourceGeometry(source, { sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 1, endColumn: 1 }), null);
});

test('PivotChart with a missing Pivot renders a broken reference instead of source-range data', () => {
  const drawing: DrawingObject = {
    id: 'broken-pivot-chart',
    sheetId: 'sheet-1',
    kind: 'chart',
    payloadId: 'broken-pivot-chart-payload',
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 160, height: 100, rotation: 0 },
    zIndex: 0,
  };
  const payload: DrawingPayload = {
    kind: 'chart',
    chartId: drawing.payloadId,
    pivotId: 'missing-pivot',
    sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }],
    chartType: 'column',
    elements: { hiddenData: 'show' },
  };
  const sheet = { id: 'sheet-1', pivotResults: {} } as unknown as CanvasSheetSnapshot;
  const drawables = createCanvasFloatingDrawables({
    drawings: [drawing],
    drawingPayloads: new Map([[drawing.payloadId, payload]]),
    allSheets: [sheet],
    sheet,
    pivotResults: {},
    sparklines: [],
    skeleton: {} as SheetSkeleton,
    imageCache: new Map(),
    requestRender: () => undefined,
    tables: [],
  });
  assert.equal(drawables.length, 1);
  assert.equal(drawables[0]?.kind, 'shape');
});
