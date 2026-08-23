import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type PivotResultTree } from '@react-sheets/core-model';
import { registerDrawingFeature } from '../drawing';
import { resolveChartData, registerChartCommands, type ChartPayload } from './index';

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
  it('persists full chart payload through one canonical drawing aggregate', () => {
    const workbook = new WorkbookModel('chart-feature-test', 'Chart Feature');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    registerChartCommands(runtime);
    const payload: ChartPayload = {
      kind: 'chart',
      chartId: 'chart-1',
      chartType: 'combo',
      title: 'Revenue',
      sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }],
      series: [
        { name: 'Revenue', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 1, endColumn: 1 }, chartType: 'column', axis: 'primary', color: '#2563eb' },
        { name: 'Margin', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 2, endColumn: 2 }, chartType: 'line', axis: 'secondary', color: '#dc2626', smooth: true },
      ],
      categoryRange: { sheetId: 'sheet-1', startRow: 1, endRow: 4, startColumn: 0, endColumn: 0 },
      categoryAxis: { id: 'x', position: 'bottom', title: 'Month' },
      valueAxis: { id: 'y', position: 'left', title: 'Revenue', minimum: 0, maximum: 1000, majorUnit: 100 },
      secondaryValueAxis: { id: 'y2', position: 'right', title: 'Margin', minimum: 0, maximum: 1, scale: 'linear' },
      legendPosition: 'bottom',
      showDataLabels: true,
      stacked: 'none',
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

    const remoteWorkbook = new WorkbookModel('chart-feature-test', 'Chart Feature');
    const remoteRuntime = new CommandRuntime(remoteWorkbook);
    registerDrawingFeature(remoteRuntime);
    registerChartCommands(remoteRuntime);
    remoteRuntime.applyRemoteMutations(runtime.getUndoEntries().flatMap((entry) => entry.redo));
    assert.equal((remoteWorkbook.getSheet('sheet-1').drawingPayloads.get('chart-1') as ChartPayload).series?.[0]?.axis, 'secondary');
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
      sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }],
      series: [
        { name: 'Revenue', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 1, endColumn: 1 }, chartType: 'scatter', axis: 'primary' },
        { name: 'Margin', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 2, endColumn: 2 }, chartType: 'scatter', axis: 'secondary' },
      ],
    };
    const local = resolveChartData(workbook, payload);
    assert.deepEqual(local.categories, ['Jan', 'Feb', 'Mar']);
    assert.deepEqual(local.series[0]?.values, [100, 120, 150]);
    assert.equal(local.series[1]?.axis, 'secondary');

    const pivotTree: PivotResultTree = {
      schema: 'PivotResultTree',
      pivotId: 'pivot-1',
      fields: { fields: [{ id: 'Month', name: 'Month', dataType: 'text', ordinal: 0 }] },
      columnPaths: [['Revenue'], ['Margin']],
      rows: [
        { kind: 'leaf', key: 'Jan', label: 'Jan', depth: 0, children: [], values: [{ columnPath: ['Revenue'], values: [100, 0.2], sourceRowPaths: [] }], subtotal: false, sourceRowPaths: [] },
        { kind: 'leaf', key: 'Feb', label: 'Feb', depth: 0, children: [], values: [{ columnPath: ['Revenue'], values: [120, 0.3], sourceRowPaths: [] }], subtotal: false, sourceRowPaths: [] },
      ],
      grandTotal: null,
      sourceRowPaths: [],
    };
    const pivotPayload: ChartPayload = { ...payload, chartId: 'pivot-chart', chartType: 'combo', pivotId: 'pivot-1', series: undefined };
    const pivotData = resolveChartData(workbook, pivotPayload, { 'pivot-1': pivotTree });
    assert.equal(pivotData.source, 'pivot');
    assert.deepEqual(pivotData.categories, ['Jan', 'Feb']);
    assert.deepEqual(pivotData.series[0]?.values, [100, 120]);

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
});
