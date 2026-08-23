import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerProSheetCommands } from './index';
import { buildPivotPanelState, setPivotAggregate } from './pivot-panel-state';
import { P1_CHART_COMMAND_IDS } from './chart-commands';

describe('M6-M8 pro sheet commands', () => {
  it('inserts chart drawing objects and exposes P1 chart command ids', () => {
    const workbook = new WorkbookModel('chart-test', 'Charts');
    const runtime = new CommandRuntime(workbook);
    registerProSheetCommands(runtime);

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
    assert.equal(sheet.charts[0]?.type, 'column');
    assert.ok(P1_CHART_COMMAND_IDS.every((commandId) => runtime.registry.hasCommand(commandId)));
  });

  it('updates pivot layout through direct PivotModel helpers and commands', () => {
    const workbook = new WorkbookModel('pivot-ext-test', 'Pivot Ext');
    const runtime = new CommandRuntime(workbook);
    registerProSheetCommands(runtime);
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
    runtime.execute('pro.pivot.add', pivot);
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
    registerProSheetCommands(runtime);

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
