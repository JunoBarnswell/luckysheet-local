import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { RangeRef, SparklineGroup, SparklineModel } from '@react-sheets/core-model';
import { buildChartLayout, resolveChartDataFromSources, type CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import { connectorEndpointHitTest, createCanvasFloatingDrawables, drawCanonicalConnectorOnCanvas, resolveCameraSourceGeometry } from './drawing-renderers';
import type { ChartDrawingPayload, ConnectorDrawingPayload, DrawingObject, DrawingPayload } from '@react-sheets/core-model';
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
  pivotTaskErrors: {},
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
    source: { kind: 'pivot', pivotId: 'missing-pivot' },
    chartType: 'column',
    subtype: 'clustered',
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
  assert.equal(drawables[0]?.kind, 'chart');
  const { context, calls } = mockCanvasContext();
  drawables[0]!.draw(context, drawing.transform);
  assert.ok(calls.some((call) => call.includes('Pivot reference unavailable: missing-pivot')));
  assert.deepEqual(drawables[0]?.hitTest?.({ x: 20, y: 20 }), {
    action: 'chart.select-element',
    data: { kind: 'chart-area' },
  });
});

test('radar chart hit testing selects the same signed vertices that its layout renders', () => {
  const values: Array<Array<string | number>> = [
    ['', 'Series'],
    ['Negative', -10],
    ['Zero', 0],
    ['Positive', 10],
  ];
  const source = {
    ...sourceSnapshot(),
    getCell: (row: number, column: number) => {
      const value = values[row]?.[column];
      return value === undefined ? undefined : { address: `${row}:${column}`, value: String(value) };
    },
  } satisfies CanvasSheetSnapshot;
  const payload: ChartDrawingPayload = {
    kind: 'chart', chartId: 'radar-hit-test', chartType: 'radar', subtype: 'radar',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: source.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
    elements: { hiddenData: 'show', legend: { visible: false, position: 'bottom' } },
  };
  const drawing: DrawingObject = {
    id: 'radar-drawing', sheetId: source.id, kind: 'chart', payloadId: payload.chartId,
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 400, height: 240, rotation: 0 }, zIndex: 0,
  };
  const data = resolveChartDataFromSources(payload, (sheetId) => sheetId === source.id ? source : undefined);
  const layout = buildChartLayout(payload, data, drawing.transform.width, drawing.transform.height);
  const [drawable] = createCanvasFloatingDrawables({
    drawings: [drawing], drawingPayloads: new Map([[payload.chartId, payload]]), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [], skeleton: {} as SheetSkeleton, imageCache: new Map(),
    requestRender: () => undefined, tables: [],
  });
  const vertices = layout.radar!.points[0]!.vertices;
  for (const pointIndex of [0, 2]) {
    const vertex = vertices[pointIndex]!;
    assert.deepEqual(drawable?.hitTest?.({ x: vertex.x, y: vertex.y }), {
      action: 'chart.select-element',
      data: { kind: 'point', seriesId: layout.series[0]!.id, pointIndex, category: layout.series[0]!.points[pointIndex]!.category },
    });
  }
  assert.notDeepEqual(vertices[0], vertices[2], 'negative and positive points occupy distinct rendered and selectable positions');
});

test('waterfall hit testing includes the visible one-pixel geometry of a zero-value bar', () => {
  const values: Array<Array<string | number>> = [
    ['', 'Change'],
    ['Start', 10],
    ['Decrease', -3],
    ['Flat', 0],
  ];
  const source = {
    ...sourceSnapshot(),
    getCell: (row: number, column: number) => {
      const value = values[row]?.[column];
      return value === undefined ? undefined : { address: `${row}:${column}`, value: String(value) };
    },
  } satisfies CanvasSheetSnapshot;
  const payload: ChartDrawingPayload = {
    kind: 'chart', chartId: 'waterfall-hit-test', chartType: 'waterfall', subtype: 'waterfall',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: source.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
    elements: { hiddenData: 'show', legend: { visible: false, position: 'bottom' } },
  };
  const drawing: DrawingObject = {
    id: 'waterfall-drawing', sheetId: source.id, kind: 'chart', payloadId: payload.chartId,
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 400, height: 240, rotation: 0 }, zIndex: 0,
  };
  const data = resolveChartDataFromSources(payload, (sheetId) => sheetId === source.id ? source : undefined);
  const layout = buildChartLayout(payload, data, drawing.transform.width, drawing.transform.height);
  const [drawable] = createCanvasFloatingDrawables({
    drawings: [drawing], drawingPayloads: new Map([[payload.chartId, payload]]), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [], skeleton: {} as SheetSkeleton, imageCache: new Map(),
    requestRender: () => undefined, tables: [],
  });
  const zeroBar = layout.waterfallBars![2]!;
  const hit = drawable?.hitTest?.({
    x: zeroBar.geometry.x + zeroBar.geometry.width / 2,
    y: zeroBar.geometry.y + zeroBar.geometry.height / 2,
  });
  assert.deepEqual(hit, {
    action: 'chart.select-element',
    data: { kind: 'point', seriesId: layout.series[0]!.id, pointIndex: 2, category: layout.series[0]!.points[2]!.category },
  });
});

test('by-category histogram hit testing returns the category represented by its rendered bin', () => {
  const values: Array<Array<string | number>> = [
    ['Category', 'Weight'],
    ['North', 2],
    ['South', 3],
    ['North', 4],
  ];
  const source = {
    ...sourceSnapshot(),
    getCell: (row: number, column: number) => {
      const value = values[row]?.[column];
      return value === undefined ? undefined : { address: `${row}:${column}`, value: String(value) };
    },
  } satisfies CanvasSheetSnapshot;
  const payload: ChartDrawingPayload = {
    kind: 'chart', chartId: 'category-histogram-hit', chartType: 'histogram', subtype: 'histogram',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: source.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
    elements: { hiddenData: 'show', legend: { visible: false, position: 'bottom' } },
    histogramOptions: { mode: 'by-category' },
  };
  const drawing: DrawingObject = {
    id: 'category-histogram-drawing', sheetId: source.id, kind: 'chart', payloadId: payload.chartId,
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 400, height: 240, rotation: 0 }, zIndex: 0,
  };
  const data = resolveChartDataFromSources(payload, (sheetId) => sheetId === source.id ? source : undefined);
  const layout = buildChartLayout(payload, data, drawing.transform.width, drawing.transform.height);
  const [drawable] = createCanvasFloatingDrawables({
    drawings: [drawing], drawingPayloads: new Map([[payload.chartId, payload]]), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [], skeleton: {} as SheetSkeleton, imageCache: new Map(),
    requestRender: () => undefined, tables: [],
  });
  const bin = layout.histogramBins![0]!;
  assert.deepEqual(drawable?.hitTest?.({ x: bin.geometry.x + bin.geometry.width / 2, y: bin.geometry.y + bin.geometry.height / 2 }), {
    action: 'chart.select-element',
    data: { kind: 'histogram-bin', seriesId: layout.series[0]!.id, binIndex: 0, category: 'North' },
  });
});

test('pie data labels render and remain selectable outside their slice geometry', () => {
  const values: Array<Array<string | number>> = [
    ['Category', 'Value'],
    ['North', 2],
    ['South', 3],
  ];
  const source = {
    ...sourceSnapshot(),
    getCell: (row: number, column: number) => {
      const value = values[row]?.[column];
      return value === undefined ? undefined : { address: `${row}:${column}`, value: String(value) };
    },
  } satisfies CanvasSheetSnapshot;
  const payload: ChartDrawingPayload = {
    kind: 'chart', chartId: 'pie-label-hit', chartType: 'pie', subtype: 'pie',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: source.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
    elements: {
      hiddenData: 'show', legend: { visible: false, position: 'bottom' },
      dataLabels: { visible: true, showCategoryName: true, showValue: true, separator: ', ', position: 'outside-end' },
    },
  };
  const drawing: DrawingObject = {
    id: 'pie-drawing', sheetId: source.id, kind: 'chart', payloadId: payload.chartId,
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 400, height: 240, rotation: 0 }, zIndex: 0,
  };
  const data = resolveChartDataFromSources(payload, (sheetId) => sheetId === source.id ? source : undefined);
  const layout = buildChartLayout(payload, data, drawing.transform.width, drawing.transform.height);
  const [drawable] = createCanvasFloatingDrawables({
    drawings: [drawing], drawingPayloads: new Map([[payload.chartId, payload]]), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [], skeleton: {} as SheetSkeleton, imageCache: new Map(),
    requestRender: () => undefined, tables: [],
  });
  const slice = layout.pieSlices![0]!;
  assert.equal(slice.dataLabelText, 'North, 2');
  const { context, calls } = mockCanvasContext();
  drawable!.draw(context, drawing.transform);
  assert.ok(calls.includes('fillText:North, 2'));
  assert.deepEqual(drawable?.hitTest?.({ x: slice.dataLabelX!, y: slice.dataLabelY! }), {
    action: 'chart.select-element',
    data: { kind: 'point', seriesId: layout.series[0]!.id, pointIndex: 0, category: layout.series[0]!.points[0]!.category },
  });
});

test('same-group sparkline bounds scan large series without argument spreading', () => {
  const pointCount = 130_000;
  const group: SparklineGroup = {
    id: 'large-sparkline-group', sheetId: 'sheet-1', type: 'line', sparklineIds: ['large-sparkline'],
    verticalAxis: { mode: 'same-group' },
  };
  const sparkline: SparklineModel = {
    id: 'large-sparkline', sheetId: 'sheet-1', anchor: { row: 0, column: 0 },
    sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: pointCount - 1, startColumn: 0, endColumn: 0 },
    type: 'line', color: '#2563eb', groupId: group.id,
  };
  const source = {
    ...sourceSnapshot(),
    rowCount: pointCount,
    sparklineGroups: [group],
    getCell: (row: number, column: number) => row >= 0 && row < pointCount
      ? { address: `${row}:${column}`, value: String(row % 31), rawValue: row % 31 }
      : undefined,
  } satisfies CanvasSheetSnapshot;
  const drawables = createCanvasFloatingDrawables({
    drawings: [], drawingPayloads: new Map(), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [sparkline], skeleton: { getCellRect: () => undefined } as unknown as SheetSkeleton,
    imageCache: new Map(), requestRender: () => undefined, tables: [],
  });

  assert.equal(drawables.length, 0, 'the geometry fixture omits a cell rectangle after shared range bounds are resolved');
});

test('connected sparkline gaps reuse the previous value without rescanning earlier cells', () => {
  const source = {
    ...sourceSnapshot(),
    getCell: (row: number, column: number) => row === 1 ? undefined : {
      address: `${row}:${column}`, value: row === 0 ? '2' : '8', rawValue: row === 0 ? 2 : 8,
    },
  } satisfies CanvasSheetSnapshot;
  const sparkline: SparklineModel = {
    id: 'connected-gap', sheetId: source.id, anchor: { row: 0, column: 0 },
    sourceRange: { sheetId: source.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    type: 'line', color: '#2563eb', emptyCells: 'connect',
  };
  const drawables = createCanvasFloatingDrawables({
    drawings: [], drawingPayloads: new Map(), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [sparkline], skeleton: { getCellRect: () => ({ x: 10, y: 20, width: 30, height: 12 }) } as unknown as SheetSkeleton,
    imageCache: new Map(), requestRender: () => undefined, tables: [],
  });
  const { context, calls } = mockCanvasContext();
  drawables[0]!.draw(context, { x: 10, y: 20, width: 30, height: 12 });

  const start = calls.find((call) => call.startsWith('moveTo:0,'));
  const connectedGap = calls.find((call) => call.startsWith('lineTo:15,'));
  const finalPoint = calls.find((call) => call.startsWith('lineTo:30,'));
  assert.ok(start && connectedGap && finalPoint, 'the line includes the first point, missing slot, and final point');
  assert.equal(connectedGap.split(',')[1], start.split(',')[1], 'the missing point connects at the last observed value');
  assert.notEqual(finalPoint.split(',')[1], connectedGap.split(',')[1]);
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

test('chart data table renders source categories and values and hit-tests only its laid out bounds', () => {
  const values: Array<Array<string | number>> = [
    ['Quarter', 'Revenue', 'Status'],
    ['Q1', 12, 'Budget'],
    ['Q2', 7, 'Actual'],
  ];
  const source = {
    ...sourceSnapshot(),
    getCell: (row: number, column: number) => {
      const value = values[row]?.[column];
      return value === undefined ? undefined : { address: `${row}:${column}`, value: String(value) };
    },
  } satisfies CanvasSheetSnapshot;
  const payload: ChartDrawingPayload = {
    kind: 'chart', chartId: 'chart-data-table', chartType: 'column', subtype: 'column',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: source.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }] },
    elements: { hiddenData: 'show', legend: { visible: false, position: 'bottom' }, dataTable: { visible: true, showLegendKeys: true } },
  };
  const drawing: DrawingObject = {
    id: 'chart-data-table-drawing', sheetId: source.id, kind: 'chart', payloadId: payload.chartId,
    anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 400, height: 240, rotation: 0 }, zIndex: 0,
  };
  const data = resolveChartDataFromSources(payload, (sheetId) => sheetId === source.id ? source : undefined);
  const layout = buildChartLayout(payload, data, drawing.transform.width, drawing.transform.height);
  assert.equal(layout.status.kind, 'ready');
  assert.equal(layout.dataTable?.categoryCount, 2);
  assert.deepEqual(layout.dataTable?.series.map((entry) => entry.values), [['12', '7'], ['Budget', 'Actual']]);
  const [drawable] = createCanvasFloatingDrawables({
    drawings: [drawing], drawingPayloads: new Map([[payload.chartId, payload]]), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [], skeleton: {} as SheetSkeleton, imageCache: new Map(),
    requestRender: () => undefined, tables: [],
  });
  const { context, calls } = mockCanvasContext();
  drawable!.draw(context, drawing.transform);
  for (const text of ['Series', 'Q1', 'Q2', 'Revenue', '12', '7', 'Status', 'Budget', 'Actual']) {
    assert.ok(calls.includes(`fillText:${text}`), `expected chart table cell ${text}`);
  }
  assert.ok(!calls.some((call) => call.includes('Chart Data Table')));
  const tableBounds = layout.dataTable!.bounds;
  assert.deepEqual(drawable?.hitTest?.({ x: tableBounds.left + 1, y: tableBounds.top + 1 }), {
    action: 'chart.select-element', data: { kind: 'data-table' },
  });
  assert.notDeepEqual(drawable?.hitTest?.({ x: tableBounds.left - 1, y: tableBounds.top + 1 }), {
    action: 'chart.select-element', data: { kind: 'data-table' },
  });

  const noLegendKeysPayload: ChartDrawingPayload = {
    ...payload,
    chartId: 'chart-data-table-no-keys',
    elements: { ...payload.elements, dataTable: { visible: true, showLegendKeys: false } },
  };
  const [noLegendKeysDrawable] = createCanvasFloatingDrawables({
    drawings: [{ ...drawing, payloadId: noLegendKeysPayload.chartId }],
    drawingPayloads: new Map([[noLegendKeysPayload.chartId, noLegendKeysPayload]]), allSheets: [source], sheet: source,
    pivotResults: {}, sparklines: [], skeleton: {} as SheetSkeleton, imageCache: new Map(),
    requestRender: () => undefined, tables: [],
  });
  const noLegendKeysCanvas = mockCanvasContext();
  noLegendKeysDrawable!.draw(noLegendKeysCanvas.context, drawing.transform);
  assert.ok(noLegendKeysCanvas.calls.includes('fillText:Revenue'), 'series names remain visible when legend keys are disabled');
  assert.ok(noLegendKeysCanvas.calls.includes('fillText:Budget'), 'text-valued source cells remain in the data table');

  const narrowLayout = buildChartLayout(payload, data, 100, drawing.transform.height);
  assert.equal(narrowLayout.status.kind, 'unsupported', 'the renderer rejects cells that cannot display even one value glyph');
  assert.equal(narrowLayout.status.code, 'UNSUPPORTED_FEATURE');
  const invalidFontPayload: ChartDrawingPayload = {
    ...payload,
    elements: { ...payload.elements, dataTable: { visible: true, font: { fontSize: Number.NaN } } },
  };
  const invalidFontLayout = buildChartLayout(invalidFontPayload, data, drawing.transform.width, drawing.transform.height);
  assert.equal(invalidFontLayout.status.kind, 'invalid');
  assert.equal(invalidFontLayout.status.code, 'INVALID_CHART_SOURCE');
  const unsupportedStylePayload: ChartDrawingPayload = {
    ...payload,
    elements: { ...payload.elements, dataTable: { visible: true, font: { rotation: 45 } } },
  };
  const unsupportedStyleLayout = buildChartLayout(unsupportedStylePayload, data, drawing.transform.width, drawing.transform.height);
  assert.equal(unsupportedStyleLayout.status.kind, 'unsupported');
  assert.equal(unsupportedStyleLayout.status.code, 'UNSUPPORTED_FEATURE');

  const piePayload: ChartDrawingPayload = { ...payload, chartId: 'pie-data-table', chartType: 'pie', subtype: 'pie' };
  const unsupportedLayout = buildChartLayout(piePayload,
    resolveChartDataFromSources(piePayload, (sheetId) => sheetId === source.id ? source : undefined),
    drawing.transform.width, drawing.transform.height);
  assert.equal(unsupportedLayout.status.kind, 'unsupported');
  assert.equal(unsupportedLayout.status.code, 'UNSUPPORTED_FEATURE');
  assert.equal(unsupportedLayout.dataTable, undefined);
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
    arc: (x: number, y: number, radius: number) => calls.push(`arc:${x},${y},${radius}`),
    ellipse: () => calls.push('ellipse'),
    strokeRect: () => calls.push('strokeRect'),
    fillRect: (x: number, y: number, width: number, height: number) => calls.push(`fillRect:${x},${y},${width},${height}`),
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
