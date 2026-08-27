import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { RangeRef } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import { connectorEndpointHitTest, createCanvasFloatingDrawables, drawCanonicalConnectorOnCanvas, resolveCameraSourceGeometry } from './drawing-renderers';
import type { ConnectorDrawingPayload, DrawingObject, DrawingPayload } from '@react-sheets/core-model';
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
    dataRegions: [],
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
    getFilterDomainDescriptor: () => ({ column: 0, values: [], scalarTypes: [], dominantType: 'empty', hasBlank: false, dateDomain: [], dateHierarchy: [], colorDomain: [], iconDomain: [], supportedFamilies: ['values'] }),
    getFilterCriterion: () => undefined,
    getFilterColorDomain: () => [],
    getFilterIconDomain: () => [],
    sheetTables: [],
    forEachOccupiedCell: () => {},
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

test('Pivot controls expose semantic child hit zones instead of a generic shape hit', () => {
  const drawing: DrawingObject = {
    id: 'slicer-control',
    sheetId: 'sheet-1',
    kind: 'slicer',
    payloadId: 'slicer-control-payload',
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 180, height: 100, rotation: 0 },
    zIndex: 0,
  };
  const payload: DrawingPayload = {
    kind: 'slicer',
    pivotId: 'pivot-1',
    fieldId: 'category',
    filter: { mode: 'all', memberKeys: [] },
    style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' },
    settings: { showHeader: true, caption: 'Category', multiSelect: true, sort: 'ascending', showNoDataItems: true, noDataItemsLast: true, showNoDataStyle: true, columnCount: 1, itemHeight: 20 },
  };
  const source = sourceSnapshot();
  const drawables = createCanvasFloatingDrawables({
    drawings: [drawing],
    drawingPayloads: new Map([[drawing.payloadId, payload]]),
    allSheets: [source],
    sheet: source,
    pivotResults: {
      'pivot-1': {
        schema: 'PivotResultTree',
        pivotId: 'pivot-1',
        fields: { fields: [{ fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0, values: ['Alpha', 'Beta'] }] },
        columnPaths: [],
        rows: [],
        grandTotal: null,
        sourceRowPaths: [],
      },
    },
    sparklines: [],
    skeleton: {} as SheetSkeleton,
    imageCache: new Map(),
    requestRender: () => undefined,
    tables: [],
  });
  assert.equal(drawables.length, 1);
  assert.equal(drawables[0]?.kind, 'pivot-control');
  const child = drawables[0]?.hitTest?.({ x: 20, y: 35 });
  assert.equal(child?.action, 'pivot.slicer.member');
  assert.deepEqual(child?.data, { kind: 'slicer-member', memberKey: { type: 'text', value: 'Alpha' } });
  const clear = drawables[0]?.hitTest?.({ x: 170, y: 13 });
  assert.equal(clear?.action, 'pivot.slicer.clear');
});

function mockCanvasContext(): { context: CanvasRenderingContext2D; calls: string[] } {
  const calls: string[] = [];
  const context = {
    save: () => calls.push('save'),
    restore: () => calls.push('restore'),
    beginPath: () => calls.push('beginPath'),
    moveTo: (x: number, y: number) => calls.push(`moveTo:${x},${y}`),
    lineTo: (x: number, y: number) => calls.push(`lineTo:${x},${y}`),
    quadraticCurveTo: (cx: number, cy: number, x: number, y: number) => calls.push(`quadratic:${cx},${cy},${x},${y}`),
    stroke: () => calls.push('stroke'),
    fill: () => calls.push('fill'),
    closePath: () => calls.push('closePath'),
    ellipse: () => calls.push('ellipse'),
    strokeRect: () => calls.push('strokeRect'),
    fillText: (text: string) => calls.push(`fillText:${text}`),
    translate: (x: number, y: number) => calls.push(`translate:${x},${y}`),
    rotate: (angle: number) => calls.push(`rotate:${angle}`),
    setLineDash: (dash: number[]) => calls.push(`dash:${dash.join(',')}`),
  } as unknown as CanvasRenderingContext2D;
  return { context, calls };
}

function connectorPayload(): ConnectorDrawingPayload {
  return {
    kind: 'connector',
    connectorType: 'curved',
    start: { drawingId: 'shape-a', connectionPoint: 'right' },
    end: { drawingId: 'shape-b', connectionPoint: 'left' },
    stroke: '#2563eb',
    strokeWidth: 2,
    startArrowhead: 'none',
    endArrowhead: 'triangle',
    route: { points: [{ x: 20, y: 30 }, { x: 60, y: 80 }, { x: 140, y: 30 }] },
  };
}

test('connector renderer follows canonical content route and draws arrowheads', () => {
  const { context, calls } = mockCanvasContext();
  drawCanonicalConnectorOnCanvas(context, connectorPayload(), { x: 0, y: 0, width: 160, height: 100 });
  assert.ok(calls.includes('moveTo:20,30'));
  assert.ok(calls.some((call) => call.startsWith('quadratic:60,80')));
  assert.ok(calls.includes('stroke'));
  assert.ok(calls.includes('fill'));
});

test('connector endpoint hit test returns semantic bound endpoint data in PaneMap-local coordinates', () => {
  const payload = connectorPayload();
  const bounds = { x: 100, y: 200, width: 80, height: 100 };
  assert.deepEqual(connectorEndpointHitTest(payload, bounds, { x: -80, y: -170 }), {
    action: 'drawing.connector.endpoint',
    data: { kind: 'connector-endpoint', edge: 'start', endpoint: payload.start },
  });
  assert.equal(connectorEndpointHitTest(payload, bounds, { x: 0, y: 0 }), null);
});

test('malformed connector payload renders an observable failure marker instead of disappearing', () => {
  const drawing: DrawingObject = {
    id: 'malformed-connector',
    sheetId: 'sheet-1',
    kind: 'connector',
    payloadId: 'malformed-connector-payload',
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 120, height: 60, rotation: 0 },
    zIndex: 0,
  };
  const malformed = {
    kind: 'connector',
    connectorType: 'elbow',
    route: { points: [{ x: 0, y: 0 }] },
  } as unknown as DrawingPayload;
  const source = sourceSnapshot();
  const drawables = createCanvasFloatingDrawables({
    drawings: [drawing],
    drawingPayloads: new Map([[drawing.payloadId, malformed]]),
    allSheets: [source],
    sheet: source,
    pivotResults: {},
    sparklines: [],
    skeleton: {} as SheetSkeleton,
    imageCache: new Map(),
    requestRender: () => undefined,
    tables: [],
  });
  assert.equal(drawables.length, 1);
  const { context, calls } = mockCanvasContext();
  drawables[0]!.draw(context, drawing.transform);
  assert.ok(calls.includes('strokeRect'));
  assert.ok(calls.some((call) => call.startsWith('fillText:Unsupported connector payload')));
});
