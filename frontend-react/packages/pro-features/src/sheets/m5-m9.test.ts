import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerChartFeature } from '../../../spreadsheet-app/src/features/chart';
import { registerPivotFeature } from '../../../spreadsheet-app/src/features/pivot';
import { registerSparklineFeature } from '../../../spreadsheet-app/src/features/sparkline';
import { buildPivotPanelState, setPivotAggregate } from './pivot-panel-state';
import { CHART_COMMAND_IDS } from '../../../spreadsheet-app/src/features/chart';

function registerSheetFeatures(runtime: CommandRuntime): void {
  registerChartFeature(runtime);
  registerPivotFeature(runtime);
  registerSparklineFeature(runtime);
}

describe('M6-M8 sheet feature commands', () => {
  it('inserts chart drawing objects and exposes P1 chart command ids', () => {
    const workbook = new WorkbookModel('chart-test', 'Charts');
    const runtime = new CommandRuntime(workbook);
    registerSheetFeatures(runtime);

    assert.equal(runtime.registry.hasCommand('pro.chart.add'), false);
    assert.equal(runtime.registry.hasCommand('pro.pivot.add'), false);
    assert.equal(runtime.registry.hasCommand('pro.sparkline.add'), false);

    runtime.execute('chart.insert.column', {
      sheetId: 'sheet-1',
      chartId: 'chart-1',
      drawingId: 'draw-chart-1',
      bounds: { x: 40, y: 40, width: 320, height: 220 },
      sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }],
      title: 'Sales',
    });

    const sheet = workbook.getSheet('sheet-1');
    assert.equal(sheet.drawings[0]?.kind, 'chart');
    assert.equal(sheet.charts.length, 0);
    assert.equal(sheet.drawingPayloads.get('chart-1')?.kind, 'chart');
    assert.ok(CHART_COMMAND_IDS.every((commandId) => runtime.registry.hasCommand(commandId)));
  });

  it('preserves combo and stacked chart semantics in the canonical payload', () => {
    const workbook = new WorkbookModel('chart-fidelity-test', 'Chart Fidelity');
    const runtime = new CommandRuntime(workbook);
    registerSheetFeatures(runtime);
    const sourceRanges = [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }];

    runtime.execute('chart.insert.combo', {
      sheetId: 'sheet-1',
      chartId: 'combo-1',
      drawingId: 'draw-combo-1',
      bounds: { x: 10, y: 10, width: 320, height: 220 },
      sourceRanges,
      stacked: 'percent',
    });
    const sheet = workbook.getSheet('sheet-1');
    const payload = sheet.drawingPayloads.get('combo-1');
    assert.equal(payload?.kind, 'chart');
    if (payload?.kind !== 'chart') throw new Error('Expected chart payload');
    assert.equal(payload.chartType, 'combo');
    assert.equal(payload.stacked, 'percent');

    runtime.execute('chart.setType', { sheetId: 'sheet-1', chartId: 'combo-1', chartType: 'combo', stacked: 'stacked' });
    assert.equal((sheet.drawingPayloads.get('combo-1') as typeof payload)?.stacked, 'stacked');
    assert.equal(runtime.undo(), true);
    assert.equal((sheet.drawingPayloads.get('combo-1') as typeof payload)?.stacked, 'percent');
  });

  it('restores a removed chart drawing and payload exactly on undo', () => {
    const workbook = new WorkbookModel('chart-remove-test', 'Chart Remove');
    const runtime = new CommandRuntime(workbook);
    registerSheetFeatures(runtime);
    runtime.execute('chart.insert.column', {
      sheetId: 'sheet-1',
      chartId: 'chart-1',
      drawingId: 'draw-chart-1',
      bounds: { x: 40, y: 40, width: 320, height: 220 },
      sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }],
      stacked: 'stacked',
    });
    const beforeDrawings = structuredClone(workbook.getSheet('sheet-1').drawings);
    const beforePayloads = structuredClone([...workbook.getSheet('sheet-1').drawingPayloads.entries()]);
    runtime.execute('chart.remove', { sheetId: 'sheet-1', chartId: 'chart-1' });
    assert.equal(workbook.getSheet('sheet-1').drawings.length, 0);
    assert.equal(runtime.undo(), true);
    const sheet = workbook.getSheet('sheet-1');
    assert.deepEqual(sheet.drawings, beforeDrawings);
    assert.deepEqual([...sheet.drawingPayloads.entries()], beforePayloads);
  });

  it('updates pivot layout through direct PivotModel helpers and commands', () => {
    const workbook = new WorkbookModel('pivot-ext-test', 'Pivot Ext');
    const runtime = new CommandRuntime(workbook);
    registerSheetFeatures(runtime);
    const pivot = {
      id: 'pivot-1',
      sheetId: 'sheet-1',
      sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      layout: {
        rows: [{ field: 'Region' }],
        columns: [],
        filters: [],
        values: [{ field: 'Amount', summarizeBy: 'sum' as const }],
        showSubtotals: true,
        showGrandTotals: true,
        compact: true,
        repeatLabels: false,
      },
    };
    runtime.execute('pivot.add', pivot);
    runtime.execute('pivot.setAggregate', { sheetId: 'sheet-1', pivotId: pivot.id, field: 'Amount', summarizeBy: 'average' });
    runtime.execute('pivot.slicer.set', {
      sheetId: 'sheet-1',
      pivotId: pivot.id,
      slicer: { id: 'slicer-1', field: 'Region', selected: ['North'] },
    });

    const updated = workbook.getSheet('sheet-1').pivots[0]!;
    assert.equal(updated.layout.values[0]?.summarizeBy, 'average');
    assert.equal(updated.slicers?.[0]?.selected[0], 'North');
    const panel = buildPivotPanelState(workbook, updated);
    assert.equal(panel.fieldCatalog.fields.length > 0, true);
    assert.equal(setPivotAggregate(updated.layout, 'Amount', 'max').values[0]?.summarizeBy, 'max');
  });

  it('creates sparkline groups and inserts by data/location ranges', () => {
    const workbook = new WorkbookModel('sparkline-test', 'Sparkline');
    const runtime = new CommandRuntime(workbook);
    registerSheetFeatures(runtime);

    runtime.execute('sparkline.insertDataLocation', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-1',
      dataRange: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 },
      location: { row: 2, column: 1 },
      type: 'line',
    });
    runtime.execute('sparkline.group.create', {
      sheetId: 'sheet-1',
      group: { id: 'group-1', sheetId: 'sheet-1', sparklineIds: ['spark-1'], showAxis: true, showMarkers: true },
    });
    runtime.execute('sparkline.group.update', {
      sheetId: 'sheet-1',
      groupId: 'group-1',
      patch: { showMarkers: false },
    });

    const sparkline = workbook.getSheet('sheet-1').sparklines[0]!;
    assert.equal(sparkline.groupId, 'group-1');
    assert.equal(sparkline.showAxis, true);
    assert.equal(sparkline.showMarkers, false);
  });
});
