import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { createPivotMemberKey, pivotMemberKey, WorkbookModel, type PivotResultTree } from '@react-sheets/core-model';
import { registerDrawingFeature } from '../drawing';
import { buildChartLayout, buildPivotChartData, chartSourceRanges, resolveChartData, resolveChartDataFromSources, resolveChartTitleText, registerChartCommands, type ChartPayload, type ResolvedChartData } from './index';

function chartPair(sheetId: string, chartId: string, payload: ChartPayload) {
  return {
    sheetId,
    drawing: {
      id: `drawing-${chartId}`,
      sheetId,
      kind: 'chart' as const,
      payloadId: chartId,
      anchor: { kind: 'two-cell' as const, row: 1, column: 1, endRow: 8, endColumn: 8 },
      transform: { x: 40, y: 50, width: 360, height: 240, rotation: 0 },
      zIndex: 1,
    },
    payload,
  };
}

describe('chart feature', () => {
  it('indexes and resolves the linked title cell used by chart projection', () => {
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'linked-title', chartType: 'line', subtype: 'line',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'owner', startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] },
      elements: { hiddenData: 'show', titleText: { linkedFormula: "='Source'!$B$1" } },
    };
    const owner = { ownerSheetId: 'owner', sheetOrder: [{ id: 'owner', name: 'Owner' }, { id: 'source', name: 'Source' }] };
    const ranges = chartSourceRanges(payload, [], owner);
    assert.ok(ranges.some((range) => range.sheetId === 'source'
      && range.startRow === 0 && range.endRow === 0 && range.startColumn === 1 && range.endColumn === 1));
    assert.equal(resolveChartTitleText(payload, owner, (range) => `${range.sheetId}!${range.startRow}:${range.startColumn}`), 'source!0:1');
  });

  it('renders a structurally invalidated linked title as #REF! instead of stale cached text', () => {
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'deleted-title-source', chartType: 'line', subtype: 'line',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'owner', startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] },
      elements: { hiddenData: 'show', titleText: { linkedFormula: '=#REF!', text: 'stale cached title' } },
    };
    const owner = { ownerSheetId: 'owner', sheetOrder: [{ id: 'owner', name: 'Owner' }] };
    assert.equal(resolveChartTitleText(payload, owner, () => assert.fail('invalid references have no cell dependency')), '#REF!');
  });

  it('persists full chart payload through one canonical drawing aggregate', () => {
    const workbook = new WorkbookModel('chart-feature-test', 'Chart Feature');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    registerChartCommands(runtime);
    const payload: ChartPayload = {
      kind: 'chart',
      chartId: 'chart-1',
      chartType: 'combo',
      subtype: 'custom-combo',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }] },
      series: [
        { name: 'Revenue', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 1, endColumn: 1 }, chartType: 'column', axis: 'primary', color: '#2563eb' },
        { name: 'Margin', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 2, endColumn: 2 }, chartType: 'line', axis: 'secondary', color: '#dc2626', smooth: true },
      ],
      categoryRange: { sheetId: 'sheet-1', startRow: 1, endRow: 4, startColumn: 0, endColumn: 0 },
      stacked: 'none',
      elements: {
        title: 'Revenue',
        legend: { visible: true, position: 'bottom' },
        dataLabels: { visible: true },
        hiddenData: 'show',
        categoryAxis: { id: 'x', position: 'bottom', title: 'Month' },
        valueAxis: { id: 'y', position: 'left', title: 'Revenue', minimum: 0, maximum: 1000, majorUnit: 100 },
        secondaryValueAxis: { id: 'y2', position: 'right', title: 'Margin', minimum: 0, maximum: 1, scale: 'linear' },
      },
    };
    runtime.execute('chart.insert', chartPair('sheet-1', 'chart-1', payload));
    const sheet = workbook.getSheet('sheet-1');
    assert.deepEqual(sheet.drawings[0]?.anchor, { kind: 'two-cell', row: 1, column: 1, endRow: 8, endColumn: 8 });
    assert.deepEqual(sheet.drawingPayloads.get('chart-1'), payload);
    const reloaded = WorkbookModel.fromSnapshot(workbook.snapshot());
    assert.deepEqual(reloaded.getSheet('sheet-1').drawingPayloads.get('chart-1'), payload);
    assert.equal(runtime.getHistoryDepth().undo, 1);

    runtime.execute('chart.setSecondaryAxis', { sheetId: 'sheet-1', chartId: 'chart-1', seriesName: 'Revenue', enabled: true });
    assert.equal((sheet.drawingPayloads.get('chart-1') as ChartPayload).series?.[0]?.axis, 'secondary');
    assert.equal(runtime.undo(), true);
    assert.equal((sheet.drawingPayloads.get('chart-1') as ChartPayload).series?.[0]?.axis, 'primary');
    assert.equal(runtime.redo(), true);

    runtime.execute('chart.setElements', { sheetId: 'sheet-1', chartId: 'chart-1', elements: { hiddenData: 'hideRows', plotArea: { fill: '#f8fafc' }, valueAxis: { id: 'y', position: 'left', minimum: 0, maximum: 2000, majorGridlines: { visible: false } } } });
    runtime.execute('chart.setSeriesStyle', { sheetId: 'sheet-1', chartId: 'chart-1', seriesName: 'Revenue', style: { marker: { enabled: true, shape: 'circle', size: 6 }, trendlines: [{ type: 'linear', color: '#2563eb' }] } });
    const edited = sheet.drawingPayloads.get('chart-1') as ChartPayload;
    assert.equal(edited.elements.hiddenData, 'hideRows');
    assert.equal(edited.elements.valueAxis?.majorGridlines?.visible, false);
    assert.equal(edited.series?.[0]?.marker?.shape, 'circle');
    assert.equal(edited.series?.[0]?.trendlines?.[0]?.type, 'linear');

    const remoteWorkbook = new WorkbookModel('chart-feature-test', 'Chart Feature');
    const remoteRuntime = new CommandRuntime(remoteWorkbook);
    registerDrawingFeature(remoteRuntime);
    registerChartCommands(remoteRuntime);
    remoteRuntime.applyRemoteMutations(runtime.getUndoEntries().flatMap((entry) => entry.redo));
    assert.equal((remoteWorkbook.getSheet('sheet-1').drawingPayloads.get('chart-1') as ChartPayload).series?.[0]?.axis, 'secondary');
    assert.equal((remoteWorkbook.getSheet('sheet-1').drawingPayloads.get('chart-1') as ChartPayload).elements.hiddenData, 'hideRows');
  });

  it('supports local range data, scatter series, pivot result data and remote replay', () => {
    const workbook = new WorkbookModel('chart-data-test', 'Chart Data');
    const sheet = workbook.getSheet('sheet-1');
    const values = [
      ['Month', 'Revenue', 'Margin'],
      ['Jan', 100, 0.2],
      ['Feb', 120, 0.3],
      ['Mar', 150, 0.4],
    ];
    values.forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart',
      chartId: 'scatter-1',
      chartType: 'scatter',
      subtype: 'scatter-markers',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }] },
      elements: { hiddenData: 'show' },
      series: [
        { name: 'Revenue', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 1, endColumn: 1 }, xRange: { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }, yRange: { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }, chartType: 'scatter', axis: 'primary' },
        { name: 'Margin', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 2, endColumn: 2 }, xRange: { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }, yRange: { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 }, chartType: 'scatter', axis: 'secondary' },
      ],
    };
    const local = resolveChartData(workbook, payload);
    assert.deepEqual(local.categories, ['Jan', 'Feb', 'Mar']);
    assert.deepEqual(local.series[0]?.values, [100, 120, 150]);
    assert.equal(local.series[1]?.axis, 'secondary');

    const pivotTree: PivotResultTree = {
      schema: 'PivotResultTree',
      pivotId: 'pivot-1',
      fields: { fields: [{ fieldId: 'Month', name: 'Month', dataType: 'text', ordinal: 0 }] },
      columnPaths: [['Revenue'], ['Margin']],
      rows: [
        { kind: 'leaf', key: 'Jan', label: 'Jan', depth: 0, children: [], values: [{ columnPath: ['Revenue'], values: [100, 0.2], sourceRowPaths: [] }], subtotal: false, sourceRowPaths: [] },
        { kind: 'leaf', key: 'Feb', label: 'Feb', depth: 0, children: [], values: [{ columnPath: ['Revenue'], values: [120, 0.3], sourceRowPaths: [] }], subtotal: false, sourceRowPaths: [] },
      ],
      grandTotal: null,
      sourceRowPaths: [],
    };
    const pivotPayload: ChartPayload = { ...payload, chartId: 'pivot-chart', chartType: 'combo', subtype: 'custom-combo', source: { kind: 'pivot', pivotId: 'pivot-1' }, series: undefined };
    const pivotData = resolveChartData(workbook, pivotPayload, { 'pivot-1': pivotTree });
    assert.equal(pivotData.source, 'pivot');
    assert.deepEqual(pivotData.categories, ['Jan', 'Feb']);
    assert.deepEqual(pivotData.series[0]?.values, [100, 120]);

    const loadingPivotData = resolveChartDataFromSources(pivotPayload, () => undefined, {}, [], new Set(['pivot-1']));
    assert.equal(loadingPivotData.status.kind, 'loading');
    assert.match(loadingPivotData.status.message ?? '', /Loading PivotTable chart data/);
    const missingPivotData = resolveChartDataFromSources(pivotPayload, () => undefined);
    assert.equal(missingPivotData.status.kind, 'invalid');
    assert.equal(missingPivotData.status.code, 'PIVOT_REFERENCE_UNAVAILABLE');

    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    registerChartCommands(runtime);
    runtime.execute('chart.insert', chartPair('sheet-1', 'remote-chart', { ...payload, chartId: 'remote-chart' }));
    const remoteWorkbook = new WorkbookModel('chart-data-test', 'Chart Data');
    const remoteRuntime = new CommandRuntime(remoteWorkbook);
    registerDrawingFeature(remoteRuntime);
    registerChartCommands(remoteRuntime);
    const operation = runtime.getUndoEntries()[0]?.redo ?? [];
    remoteRuntime.applyRemoteMutations(operation);
    assert.equal(remoteWorkbook.getSheet('sheet-1').drawingPayloads.get('remote-chart')?.kind, 'chart');
    assert.equal(remoteWorkbook.getSheet('sheet-1').drawings[0]?.kind, 'chart');
  });

  it('projects the complete Pivot row-path × column-path × values matrix', () => {
    const member = (fieldId: string, value: string): string => `${fieldId}=${pivotMemberKey(createPivotMemberKey(value))}`;
    const pivotTree: PivotResultTree = {
      schema: 'PivotResultTree',
      pivotId: 'matrix-pivot',
      fields: {
        fields: [
          { fieldId: 'region', name: 'Region', dataType: 'text', ordinal: 0, values: ['East', 'West'] },
          { fieldId: 'product', name: 'Product', dataType: 'text', ordinal: 1, values: ['Widget', 'Gadget'] },
        ],
      },
      columnPaths: [['Jan'], ['Feb']],
      valueFields: [
        { valueId: 'value:sales', fieldId: 'sales', sourceFieldId: 'sales', displayName: 'Sales', summarizeBy: 'sum' },
        { valueId: 'value:count', fieldId: 'count', sourceFieldId: 'count', displayName: 'Orders', summarizeBy: 'count' },
      ],
      rows: [
        {
          kind: 'subtotal', key: 'East', label: 'East', depth: 0, subtotal: true,
          path: [member('region', 'East')], children: [
            { kind: 'leaf', key: 'Widget', label: 'Widget', depth: 1, subtotal: false, path: [member('region', 'East'), member('product', 'Widget')], children: [], values: [
              { columnPath: ['Jan'], values: [10, 1], sourceRowPaths: [] }, { columnPath: ['Feb'], values: [20, 2], sourceRowPaths: [] },
            ], sourceRowPaths: [] },
            { kind: 'leaf', key: 'Gadget', label: 'Gadget', depth: 1, subtotal: false, path: [member('region', 'East'), member('product', 'Gadget')], children: [], values: [
              { columnPath: ['Jan'], values: [11, 3], sourceRowPaths: [] }, { columnPath: ['Feb'], values: [21, 4], sourceRowPaths: [] },
            ], sourceRowPaths: [] },
          ], values: [], sourceRowPaths: [],
        },
        {
          kind: 'subtotal', key: 'West', label: 'West', depth: 0, subtotal: true,
          path: [member('region', 'West')], children: [
            { kind: 'leaf', key: 'Widget', label: 'Widget', depth: 1, subtotal: false, path: [member('region', 'West'), member('product', 'Widget')], children: [], values: [
              { columnPath: ['Jan'], values: [30, 5], sourceRowPaths: [] }, { columnPath: ['Feb'], values: [40, 6], sourceRowPaths: [] },
            ], sourceRowPaths: [] },
          ], values: [], sourceRowPaths: [],
        },
      ],
      grandTotal: null,
      sourceRowPaths: [],
    };
    const projected = buildPivotChartData(pivotTree);
    assert.deepEqual(projected.categories.map((category) => category.label), ['East / Widget', 'East / Gadget', 'West / Widget']);
    assert.deepEqual(projected.series.map((entry) => entry.name), ['Jan Sales', 'Jan Orders', 'Feb Sales', 'Feb Orders']);
    assert.deepEqual(projected.series.map((entry) => entry.values), [[10, 11, 30], [1, 3, 5], [20, 21, 40], [2, 4, 6]]);
    assert.notEqual(projected.categories[0]?.id, projected.categories[2]?.id);

    const noRows: PivotResultTree = {
      schema: 'PivotResultTree', pivotId: 'root-pivot', fields: { fields: [] }, columnPaths: [[]],
      valueFields: [{ valueId: 'value:amount', fieldId: 'amount', sourceFieldId: 'amount', displayName: 'Amount', summarizeBy: 'sum' }],
      rows: [{ kind: 'leaf', key: null, label: 'Values', depth: 0, path: ['__root__'], children: [], subtotal: false, values: [{ columnPath: [], values: [null], sourceRowPaths: [] }], sourceRowPaths: [] }],
      grandTotal: null, sourceRowPaths: [],
    };
    const rootProjection = buildPivotChartData(noRows);
    assert.deepEqual(rootProjection.categories.map((category) => category.label), ['Values']);
    assert.deepEqual(rootProjection.series[0]?.values, [null]);
  });

  it('resolves XY and Bubble bindings positionally and applies empty-cell policy before layout', () => {
    const workbook = new WorkbookModel('chart-xy-test', 'XY Charts');
    const sheet = workbook.getSheet('sheet-1');
    [
      ['Label', 'X', 'Y', 'Size'],
      ['A', 1, 10, 5],
      ['B', 2, null, 12],
      ['C', 4, 40, 8],
    ].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const range = { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 };
    const xRange = { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 };
    const yRange = { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 };
    const sizeRange = { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 3, endColumn: 3 };
    const scatter: ChartPayload = {
      kind: 'chart', chartId: 'xy-chart', chartType: 'scatter', subtype: 'scatter-markers', source: { kind: 'worksheet-ranges', ranges: [range] },
      series: [{ id: 'xy-series', name: 'Y', range: yRange, xRange, yRange, chartType: 'scatter' }], categoryRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }, elements: { hiddenData: 'show', emptyCells: 'gap' },
    };
    const scatterData = resolveChartData(workbook, scatter);
    assert.deepEqual(scatterData.series[0]?.xValues, [1, 2, 4]);
    assert.deepEqual(scatterData.series[0]?.values, [10, null, 40]);
    const scatterLayout = buildChartLayout(scatter, scatterData, 400, 240);
    assert.equal(scatterLayout.status.kind, 'ready');
    assert.deepEqual(scatterLayout.series[0]?.points.map((point) => point.xValue), [1, 2, 4]);
    assert.equal(scatterLayout.series[0]?.points[1]?.visible, false);

    const bubble: ChartPayload = { ...scatter, chartId: 'bubble-chart', chartType: 'bubble', subtype: 'bubble', series: [{ id: 'bubble-series', name: 'Y', range: yRange, xRange, yRange, sizeRange, chartType: 'bubble' }] };
    const bubbleData = resolveChartData(workbook, bubble);
    assert.deepEqual(bubbleData.series[0]?.sizeValues, [5, 12, 8]);
    const bubbleLayout = buildChartLayout(bubble, bubbleData, 400, 240);
    assert.equal(bubbleLayout.status.kind, 'ready');
    assert.deepEqual(bubbleLayout.series[0]?.points.map((point) => point.sizeValue), [5, 12, 8]);

    const zeroData = resolveChartData(workbook, { ...scatter, chartId: 'xy-zero', elements: { hiddenData: 'show', emptyCells: 'zero' } });
    assert.deepEqual(zeroData.series[0]?.values, [10, 0, 40]);
    assert.deepEqual(zeroData.series[0]?.missing, [false, false, false]);
  });

  it('separates clustered horizontal bars and derives positive automatic logarithmic bounds', () => {
    const workbook = new WorkbookModel('chart-axis-and-bars', 'Chart geometry');
    const sheet = workbook.getSheet('sheet-1');
    [
      ['Category', 'Series A', 'Series B'],
      ['East', 2, 6],
      ['West', 4, 8],
    ].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));

    const bar: ChartPayload = {
      kind: 'chart', chartId: 'clustered-bars', chartType: 'bar', subtype: 'clustered',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }] },
      elements: { hiddenData: 'show' },
    };
    const barLayout = buildChartLayout(bar, resolveChartData(workbook, bar), 400, 240);
    assert.equal(barLayout.status.kind, 'ready');
    const eastA = barLayout.series[0]?.bars.find((entry) => entry.index === 0);
    const eastB = barLayout.series[1]?.bars.find((entry) => entry.index === 0);
    assert.ok(eastA && eastB);
    assert.ok(eastA.y < eastB.y, 'series in one category occupy separate vertical slots');
    assert.equal(eastA.height, eastB.height);
    assert.ok(eastA.y + eastA.height <= eastB.y);

    const line: ChartPayload = {
      kind: 'chart', chartId: 'log-axis', chartType: 'line', subtype: 'line',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show', valueAxis: { id: 'y', position: 'left', scale: 'logarithmic' } },
    };
    const logarithmicLayout = buildChartLayout(line, resolveChartData(workbook, line), 400, 240);
    assert.equal(logarithmicLayout.status.kind, 'ready');
    assert.ok(logarithmicLayout.valueAxis!.minimum > 0);
    assert.ok(logarithmicLayout.valueAxis!.maximum > logarithmicLayout.valueAxis!.minimum);

    const invalidBounds = { ...line, elements: { ...line.elements, valueAxis: { id: 'y', position: 'left' as const, scale: 'logarithmic' as const, maximum: 1 } } };
    assert.equal(buildChartLayout(invalidBounds, resolveChartData(workbook, invalidBounds), 400, 240).status.kind, 'invalid');
  });

  it('uses signed value-axis geometry for radar points and preserves missing categories', () => {
    const workbook = new WorkbookModel('radar-signed-values', 'Radar signed values');
    const sheet = workbook.getSheet('sheet-1');
    [
      ['', 'Series'],
      ['Negative', -10],
      ['Zero', 0],
      ['Positive', 10],
      ['Missing', null],
    ].forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
      if (value !== null) sheet.cells.set(rowIndex, columnIndex, { value });
    }));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'radar-signed', chartType: 'radar', subtype: 'radar',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    const radar = layout.radar!;
    const vertices = radar.points[0]!.vertices;
    assert.equal(vertices.length, 4);
    assert.deepEqual(vertices.map((vertex) => vertex.visible), [true, true, true, false]);
    const radii = vertices.slice(0, 3).map((vertex) => Math.hypot(vertex.x - radar.centerX, vertex.y - radar.centerY));
    assert.notEqual(radii[0], radii[2], 'negative and positive values must not collapse to the same radius');
    assert.ok(radii[0]! < radii[1]! && radii[1]! < radii[2]!, 'negative, zero, and positive values follow the signed value axis');
  });

  it('keeps waterfall delta direction for connectors and publishes visible bar geometry', () => {
    const workbook = new WorkbookModel('waterfall-directed-geometry', 'Waterfall geometry');
    const sheet = workbook.getSheet('sheet-1');
    [['', 'Change'], ['Start', 10], ['Decrease', -3], ['Flat', 0], ['Total', 7]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'waterfall-directed', chartType: 'waterfall', subtype: 'waterfall',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      waterfallOptions: { connectorLines: true, totalPointIndexes: [3] },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    const bars = layout.waterfallBars!;
    assert.equal(bars[1]!.connector!.y, bars[1]!.geometry.y, 'a negative delta connects at the prior cumulative value on the top of the bar');
    assert.equal(bars[1]!.connector!.startX, bars[0]!.geometry.x + bars[0]!.geometry.width, 'connectors begin at the previous bar edge');
    assert.equal(bars[1]!.connector!.endX, bars[1]!.geometry.x, 'connectors end at the next bar edge');
    assert.equal(bars[2]!.geometry.height, 1, 'zero-value bars retain the renderer minimum hit target height');
    assert.equal(bars[2]!.geometry.width, bars[0]!.geometry.width);
    assert.equal(bars[3]!.connector!.y, bars[3]!.geometry.y, 'a total bar connects from the preceding cumulative total, not the zero baseline');
  });

  it('builds a chart above the engine argument-expansion limit without spreading point arrays', () => {
    const pointCount = 130_000;
    const categories = Array.from({ length: pointCount }, (_value, index) => index);
    const values = Array.from({ length: pointCount }, (_value, index) => index % 31);
    const series = { id: 'large-series', name: 'Large series', values, axis: 'primary' as const };
    const data: ResolvedChartData = {
      categories,
      series: [series],
      source: 'range',
      binding: { source: 'range', orientation: 'columns', categories, series: [series], hierarchyLevels: [], nonContiguous: false },
      status: { kind: 'ready' },
    };
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'large-line', chartType: 'line', subtype: 'line',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'large-sheet', startRow: 0, endRow: pointCount, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
    };

    const layout = buildChartLayout(payload, data, 480, 260);

    assert.equal(layout.status.kind, 'ready');
    assert.equal(layout.series[0]!.points.length, pointCount);
  });

  it('projects box-whisker inner points and mean-marker options into chart facts', () => {
    const workbook = new WorkbookModel('box-whisker-options', 'Box plot options');
    const sheet = workbook.getSheet('sheet-1');
    [['', 'Score'], ['A', 1], ['B', 2], ['C', 3], ['D', 4], ['E', 100]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'box-whisker-options', chartType: 'box-whisker', subtype: 'box-whisker',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      boxWhiskerOptions: { quartile: 'inclusive-median', showInnerPoints: true, showOutlierPoints: true, showMeanMarkers: true },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.deepEqual(layout.boxes![0]!.innerPoints, [1, 2, 3, 4]);
    assert.equal(layout.boxes![0]!.showMeanMarker, true);
    assert.equal(layout.boxes![0]!.mean, 22);
    assert.deepEqual(layout.boxes![0]!.outliers, [100]);
  });

  it('groups duplicate text categories and sums their values in by-category histograms', () => {
    const workbook = new WorkbookModel('histogram-by-category', 'Histogram categories');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Weight'], ['North', 2], ['South', 3], ['North', 4]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'category-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      histogramOptions: { mode: 'by-category' },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.deepEqual(layout.histogramBins!.map((bin) => [bin.kind, bin.category, bin.count, bin.value, bin.label]), [
      ['category', 'North', 2, 6, 'North'],
      ['category', 'South', 1, 3, 'South'],
    ]);
    assert.ok(layout.histogramBins!.every((bin) => bin.geometry.height > 0));
  });

  it('preserves underflow, overflow, and exact upper-bound values in numeric histograms', () => {
    const workbook = new WorkbookModel('histogram-boundaries', 'Histogram boundaries');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['A', -10], ['B', -5], ['C', 0], ['D', 5], ['E', 10], ['F', 15]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'numeric-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 6, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      histogramOptions: { mode: 'bin-width', binWidth: 5, underflow: 0, overflow: 10 },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.equal(layout.series[0]!.points.length, 0);
    assert.deepEqual(layout.histogramBins!.map((bin) => [bin.kind, bin.start, bin.end, bin.count, bin.label, bin.boundary]), [
      ['numeric', 0, 0, 3, '≤ 0', 'underflow'],
      ['numeric', 0, 5, 1, '(0, 5]', undefined],
      ['numeric', 5, 10, 1, '(5, 10]', undefined],
      ['numeric', 10, 10, 1, '> 10', 'overflow'],
    ]);
    assert.equal(layout.histogramBins!.reduce((sum, bin) => sum + bin.count, 0), 6);
  });

  it('counts underflow and overflow bins inside an explicit bin count', () => {
    const workbook = new WorkbookModel('histogram-bin-count', 'Histogram bin count');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['A', 0], ['B', 1], ['C', 2], ['D', 3]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'counted-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      histogramOptions: { mode: 'bin-count', binCount: 4, underflow: 0, overflow: 3 },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.equal(layout.histogramBins!.length, 4);
    assert.deepEqual(layout.histogramBins!.map((bin) => bin.count), [1, 1, 2, 0]);
    assert.equal(layout.histogramBins!.reduce((sum, bin) => sum + bin.count, 0), 4);
  });

  it('allows equal tail thresholds because underflow and overflow remain disjoint', () => {
    const workbook = new WorkbookModel('histogram-equal-tails', 'Histogram equal tail thresholds');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['A', -1], ['B', 0], ['C', 1]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'equal-tail-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      histogramOptions: { mode: 'bin-count', binCount: 2, underflow: 0, overflow: 0 },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.deepEqual(layout.histogramBins!.map((bin) => [bin.boundary, bin.count]), [['underflow', 2], ['overflow', 1]]);
  });

  it('rejects mismatched category and value vectors instead of truncating the longer source', () => {
    const workbook = new WorkbookModel('histogram-category-mismatch', 'Histogram category mismatch');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['A', 1], ['B', 2]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'mismatched-category-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' }, histogramOptions: { mode: 'by-category' },
    };
    const data = resolveChartData(workbook, payload);
    data.categories = data.categories.slice(0, 1);

    const layout = buildChartLayout(payload, data, 400, 240);
    assert.equal(layout.status.kind, 'invalid');
    assert.equal(layout.histogramBins, undefined);
  });

  it('rejects by-category aggregate and geometry overflow before emitting non-finite bars', () => {
    const workbook = new WorkbookModel('histogram-category-overflow', 'Histogram category overflow');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['A', Number.MAX_VALUE], ['A', Number.MAX_VALUE]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'overflowing-category-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' }, histogramOptions: { mode: 'by-category' },
    };

    const aggregateLayout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(aggregateLayout.status.kind, 'invalid');
    assert.equal(aggregateLayout.histogramBins, undefined);

    sheet.cells.set(2, 0, { value: 'B' });
    sheet.cells.set(1, 1, { value: -Number.MAX_VALUE });
    sheet.cells.set(2, 1, { value: Number.MAX_VALUE });
    const geometryLayout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(geometryLayout.status.kind, 'invalid');
    assert.equal(geometryLayout.histogramBins, undefined);
  });

  it('uses explicit tail boundaries without requiring an overflowing full-data span', () => {
    const workbook = new WorkbookModel('histogram-extreme-tails', 'Histogram extreme tails');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['Low', -1e308], ['High', 1e308]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'extreme-tail-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
      histogramOptions: { mode: 'bin-width', binWidth: 1, underflow: 0, overflow: 1 },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.deepEqual(layout.histogramBins!.map((bin) => bin.count), [1, 0, 1]);
  });

  it('rejects invalid bin settings and bin widths beyond drawable resolution before materializing bins', () => {
    const workbook = new WorkbookModel('histogram-bounds', 'Histogram bounds');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['A', 0], ['B', 1], ['C', 2], ['D', 3]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const source = { kind: 'worksheet-ranges' as const, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }] };
    const invalid: ChartPayload = {
      kind: 'chart', chartId: 'invalid-width', chartType: 'histogram', subtype: 'histogram', source,
      elements: { hiddenData: 'show' }, histogramOptions: { mode: 'bin-width', binWidth: 0 },
    };
    const excessive: ChartPayload = {
      kind: 'chart', chartId: 'excessive-width', chartType: 'histogram', subtype: 'histogram', source,
      elements: { hiddenData: 'show' }, histogramOptions: { mode: 'bin-width', binWidth: 1e-18 },
    };

    const invalidLayout = buildChartLayout(invalid, resolveChartData(workbook, invalid), 400, 240);
    const excessiveLayout = buildChartLayout(excessive, resolveChartData(workbook, excessive), 400, 240);
    assert.equal(invalidLayout.status.kind, 'invalid');
    assert.equal(excessiveLayout.status.kind, 'unsupported');
    assert.equal(excessiveLayout.histogramBins, undefined);
  });

  it('rejects multiple visible source series instead of silently histogramming only the first', () => {
    const workbook = new WorkbookModel('histogram-series', 'Histogram series');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'First', 'Second'], ['A', 1, 10], ['B', 2, 20], ['C', 3, 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'multi-series-histogram', chartType: 'histogram', subtype: 'histogram',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }] },
      elements: { hiddenData: 'show' },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'invalid');
    assert.equal(layout.histogramBins, undefined);
  });

  it('projects pie data-label text and hit bounds from the same slice facts', () => {
    const workbook = new WorkbookModel('pie-data-labels', 'Pie labels');
    const sheet = workbook.getSheet('sheet-1');
    [['Category', 'Value'], ['North', 2], ['South', 3]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = {
      kind: 'chart', chartId: 'pie-labels', chartType: 'pie', subtype: 'pie',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show', dataLabels: { visible: true, showCategoryName: true, showValue: true, showPercentage: true, separator: ' · ', position: 'outside-end' } },
    };

    const layout = buildChartLayout(payload, resolveChartData(workbook, payload), 400, 240);
    assert.equal(layout.status.kind, 'ready');
    assert.deepEqual(layout.pieSlices!.map((slice) => slice.dataLabelText), ['North · 2 · 40%', 'South · 3 · 60%']);
    assert.ok(layout.pieSlices!.every((slice) => slice.dataLabelBounds !== undefined));

    const noLabelsPayload: ChartPayload = {
      ...payload,
      chartId: 'pie-no-label-fields',
      elements: { hiddenData: 'show', dataLabels: { visible: true, showSeriesName: false, showCategoryName: false, showValue: false, showPercentage: false } },
    };
    const noLabelsLayout = buildChartLayout(noLabelsPayload, resolveChartData(workbook, noLabelsPayload), 400, 240);
    assert.ok(noLabelsLayout.pieSlices!.every((slice) => slice.dataLabelText === undefined), 'explicitly disabling every label field produces no label');
  });

  it('switches row-oriented worksheet matrices without converting categories into X coordinates', () => {
    const workbook = new WorkbookModel('chart-row-orientation', 'Row Orientation');
    const sheet = workbook.getSheet('sheet-1');
    [['', 'Jan', 'Feb', 'Mar'], ['Revenue', 10, 20, 30], ['Cost', 4, 8, 12]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const payload: ChartPayload = { kind: 'chart', chartId: 'row-chart', chartType: 'line', subtype: 'line-markers', dataOrientation: 'rows', source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 3 }] }, elements: { hiddenData: 'show' } };
    const data = resolveChartData(workbook, payload);
    assert.deepEqual(data.categories, ['Jan', 'Feb', 'Mar']);
    assert.deepEqual(data.series.map((series) => series.values), [[10, 20, 30], [4, 8, 12]]);
    assert.deepEqual(data.series.map((series) => series.name), ['Revenue', 'Cost']);
  });

  it('rejects a chart that declares a semantic XY family without independent bindings', () => {
    const workbook = new WorkbookModel('chart-reject-test', 'Chart Reject');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    registerChartCommands(runtime);
    const payload: ChartPayload = { kind: 'chart', chartId: 'invalid-xy', chartType: 'scatter', subtype: 'scatter-markers', source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }] }, series: [{ name: 'Y', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 1, endColumn: 1 }, chartType: 'scatter' }], elements: { hiddenData: 'show' } };
    assert.throws(() => runtime.execute('chart.insert', chartPair('sheet-1', 'invalid-xy', payload)), /explicit X\/Y range bindings/);
  });
});
