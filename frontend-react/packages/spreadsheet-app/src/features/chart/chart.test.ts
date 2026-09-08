import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createPivotMemberKey, pivotMemberKey, type PivotResultTree } from '@react-sheets/core-model';
import { buildChartLayout, buildPivotChartData, registerChartCommands, type ChartPayload } from './index';
import { resolveChartDataFromSources } from './data';
import type { ResolvedVisibility } from '@react-sheets/sheet-features';
import { registerDrawingFeature } from '../drawing';
import { openCanonicalTestRuntime, seedCanonicalCells } from '../../../../core-model/src/canonical-test-runtime.test';

function chartPair(sheetId: string, chartId: string, payload: ChartPayload) {
  return {
    sheetId,
    drawing: {
      id: `drawing-${chartId}`, sheetId, kind: 'chart' as const, payloadId: chartId,
      anchor: { kind: 'two-cell' as const, row: 1, column: 1, endRow: 8, endColumn: 8 },
      transform: { x: 40, y: 50, width: 360, height: 240, rotation: 0 }, zIndex: 1,
    },
    payload,
  };
}

async function setup(unitId: string) {
  const fixture = await openCanonicalTestRuntime(unitId, 'Chart');
  registerDrawingFeature(fixture.runtime);
  registerChartCommands(fixture.runtime);
  return fixture;
}

function visibility(revision: number): ResolvedVisibility {
  const rows = new Map<number, { manualHidden: boolean; filterHidden: boolean; outlineHidden: boolean }>();
  const columns = new Map<number, { manualHidden: boolean }>();
  return { revision, rows, columns, isRowHidden: (row) => rows.has(row), isColumnHidden: (column) => columns.has(column) };
}

describe('chart feature', () => {
  it('persists chart inserts and edits through canonical drawing mutations', async () => {
    const { workbook, runtime, close } = await setup('chart-canonical-mutations');
    try {
      const payload: ChartPayload = {
        kind: 'chart', chartId: 'chart-1', chartType: 'combo', subtype: 'custom-combo',
        source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }] },
        series: [
          { name: 'Revenue', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 1, endColumn: 1 }, chartType: 'column', axis: 'primary', color: '#2563eb' },
          { name: 'Margin', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 2, endColumn: 2 }, chartType: 'line', axis: 'secondary', color: '#dc2626' },
        ],
        categoryRange: { sheetId: 'sheet-1', startRow: 1, endRow: 4, startColumn: 0, endColumn: 0 },
        stacked: 'none', elements: { title: 'Revenue', hiddenData: 'show', legend: { visible: true, position: 'bottom' } },
      };
      const planned: string[] = [];
      runtime.onMutation((mutation) => planned.push(mutation.id));
      await runtime.execute('chart.insert', chartPair('sheet-1', 'chart-1', payload));
      await runtime.execute('chart.setSecondaryAxis', { sheetId: 'sheet-1', chartId: 'chart-1', seriesName: 'Revenue', enabled: true });
      await runtime.execute('chart.setElements', { sheetId: 'sheet-1', chartId: 'chart-1', elements: { hiddenData: 'hideRows' } });
      const sheet = workbook.getSheet('sheet-1');
      assert.deepEqual(planned, ['drawing.add', 'drawing.payload.update', 'drawing.payload.update']);
      assert.deepEqual(sheet.drawings[0]?.anchor, { kind: 'two-cell', row: 1, column: 1, endRow: 8, endColumn: 8 });
      const edited = sheet.drawingPayloads.get('chart-1') as ChartPayload;
      assert.equal(edited.elements.hiddenData, 'hideRows');
      assert.equal(edited.series?.[0]?.axis, 'secondary');
      assert.equal(await runtime.undo(), true);
      assert.equal((sheet.drawingPayloads.get('chart-1') as ChartPayload).elements.hiddenData, 'show');
      assert.equal(await runtime.redo(), true);
    } finally {
      close();
    }
  });

  it('resolves worksheet ranges and XY bindings from canonical cells', async () => {
    const { workbook, runtime, close } = await setup('chart-canonical-data');
    try {
      await seedCanonicalCells(runtime, 'sheet-1', [
        { row: 0, column: 0, value: 'Month' }, { row: 0, column: 1, value: 'Revenue' }, { row: 0, column: 2, value: 'Margin' },
        { row: 1, column: 0, value: 'Jan' }, { row: 1, column: 1, value: 100 }, { row: 1, column: 2, value: 0.2 },
        { row: 2, column: 0, value: 'Feb' }, { row: 2, column: 1, value: 120 }, { row: 2, column: 2, value: 0.3 },
        { row: 3, column: 0, value: 'Mar' }, { row: 3, column: 1, value: 150 }, { row: 3, column: 2, value: 0.4 },
      ]);
      const range = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 };
      const xRange = { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 };
      const yRange = { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 };
      const payload: ChartPayload = {
        kind: 'chart', chartId: 'scatter-1', chartType: 'scatter', subtype: 'scatter-markers', source: { kind: 'worksheet-ranges', ranges: [range] },
        series: [{ name: 'Margin', range: yRange, xRange, yRange, chartType: 'scatter' }],
        categoryRange: { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }, elements: { hiddenData: 'show', emptyCells: 'gap' },
      };
      const sheet = workbook.getSheet('sheet-1');
      const resolvedVisibility = visibility(sheet.cells.revision);
      const data = resolveChartDataFromSources(payload, () => ({
        getCell: (row: number, column: number) => sheet.cells.get(row, column),
        resolvedVisibility,
        revision: sheet.cells.revision,
      }));
      assert.deepEqual(data.categories, ['Jan', 'Feb', 'Mar']);
      assert.deepEqual(data.series[0]?.xValues, [100, 120, 150]);
      assert.deepEqual(data.series[0]?.values, [0.2, 0.3, 0.4]);
      const layout = buildChartLayout(payload, data, 400, 240);
      assert.equal(layout.status.kind, 'ready');
    } finally {
      close();
    }
  });

  it('projects the complete pivot row-path by column-path matrix', () => {
    const member = (fieldId: string, value: string): string => `${fieldId}=${pivotMemberKey(createPivotMemberKey(value))}`;
    const pivotTree: PivotResultTree = {
      schema: 'PivotResultTree', pivotId: 'matrix-pivot', fields: { fields: [{ fieldId: 'region', name: 'Region', dataType: 'text', ordinal: 0 }] },
      columnPaths: [['Jan'], ['Feb']],
      valueFields: [{ valueId: 'value:sales', fieldId: 'sales', sourceFieldId: 'sales', displayName: 'Sales', summarizeBy: 'sum' }],
      rows: [
        { kind: 'leaf', key: 'East', label: 'East', depth: 0, subtotal: false, path: [member('region', 'East')], children: [], values: [{ columnPath: ['Jan'], values: [10], sourceRowPaths: [] }, { columnPath: ['Feb'], values: [20], sourceRowPaths: [] }], sourceRowPaths: [] },
        { kind: 'leaf', key: 'West', label: 'West', depth: 0, subtotal: false, path: [member('region', 'West')], children: [], values: [{ columnPath: ['Jan'], values: [30], sourceRowPaths: [] }, { columnPath: ['Feb'], values: [40], sourceRowPaths: [] }], sourceRowPaths: [] },
      ],
      grandTotal: null, sourceRowPaths: [],
    };
    const projected = buildPivotChartData(pivotTree);
    assert.deepEqual(projected.categories.map((category) => category.label), ['East', 'West']);
    assert.deepEqual(projected.series.map((entry) => entry.name), ['Jan Sales', 'Feb Sales']);
    assert.deepEqual(projected.series.map((entry) => entry.values), [[10, 30], [20, 40]]);
  });

  it('rejects an XY chart without independent bindings before commit', async () => {
    const { runtime, close } = await setup('chart-canonical-rejection');
    try {
      const payload: ChartPayload = {
        kind: 'chart', chartId: 'invalid-xy', chartType: 'scatter', subtype: 'scatter-markers',
        source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }] },
        series: [{ name: 'Y', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 1, endColumn: 1 }, chartType: 'scatter' }], elements: { hiddenData: 'show' },
      };
      await assert.rejects(() => runtime.execute('chart.insert', chartPair('sheet-1', 'invalid-xy', payload)), /explicit X\/Y range bindings/);
    } finally {
      close();
    }
  });
});
