import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { ChartDrawingPayload, DrawingObject, WorkbookTableModel } from '@react-sheets/core-model';
import { WorkbookSession } from '../../workbook-session';

describe('ProjectionRuntime chart dependencies', () => {
  it('invalidates a cross-sheet chart when its workbook table source is removed', async () => {
    const app = new WorkbookSession();
    const sourceSheetId = app.getActiveSheetId();
    app.addSheet();
    const ownerSheetId = app.getActiveSheetId();
    const runtime = app['runtime'];
    const table: WorkbookTableModel = {
      id: 'sales-table',
      name: 'Sales',
      sourceSheetId,
      sourceRange: { sheetId: sourceSheetId, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      rowCount: 3,
      fields: [
        { id: 'category', name: 'Category', ordinal: 0, type: 'text' },
        { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' },
      ],
      blockSize: 1024,
      blocks: [],
      revision: 0,
    };
    runtime.model.addTable(table);

    const chartId = 'sales-chart';
    const owner = runtime.model.getSheet(ownerSheetId);
    const drawing: DrawingObject = {
      id: 'sales-chart-drawing',
      sheetId: ownerSheetId,
      kind: 'chart',
      payloadId: chartId,
      anchor: { kind: 'absolute' },
      transform: { x: 0, y: 0, width: 320, height: 200, rotation: 0 },
      zIndex: 0,
    };
    const payload: ChartDrawingPayload = {
      kind: 'chart',
      chartId,
      chartType: 'column',
      subtype: 'clustered',
      source: {
        kind: 'table',
        tableId: table.id,
        bindings: {
          values: [{ area: 'values', fieldId: 'amount', aggregate: 'sum' }],
          category: [{ area: 'category', fieldId: 'category', aggregate: 'none' }],
          details: [],
          color: [],
          size: [],
          tooltip: [],
          filter: [],
        },
      },
      elements: { hiddenData: 'show' },
    };
    owner.drawings.push(drawing);
    owner.drawingPayloads.set(chartId, payload);

    const projection = app['projection'];
    projection.invalidateDependentChartProjections([]);
    const beforeRemoval = projection.getCanvasProjection(owner);

    await app.removeDataTable(table.id);

    const afterRemoval = projection.getCanvasProjection(owner);
    assert.notEqual(afterRemoval, beforeRemoval);
  });
});
