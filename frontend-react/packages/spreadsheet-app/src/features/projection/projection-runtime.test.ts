import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { ChartDrawingPayload, DrawingObject, ShapeDrawingPayload, StructuralFormulaOwnerDelta, WorkbookTableModel } from '@react-sheets/core-model';
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

describe('ProjectionRuntime structural formula owners', () => {
  it('invalidates the formula owner sheet when a structural patch targets another sheet', () => {
    const app = new WorkbookSession();
    const sourceSheetId = app.getActiveSheetId();
    app.addSheet();
    const ownerSheetId = app.getActiveSheetId();
    const runtime = app['runtime'];
    const owner = runtime.model.getSheet(ownerSheetId);
    const projection = app['projection'];
    const beforeFormula = '=Sales[Amount]';
    const afterFormula = '=Orders[Amount]';
    owner.cells.set(0, 0, { value: null, formula: beforeFormula });
    const before = projection.getCanvasProjection(owner);

    owner.cells.set(0, 0, { value: null, formula: afterFormula });
    const delta: StructuralFormulaOwnerDelta = {
      kind: 'formula-cell',
      beforeAddress: { sheetId: ownerSheetId, row: 0, column: 0 },
      afterAddress: { sheetId: ownerSheetId, row: 0, column: 0 },
      before: { formula: beforeFormula, sourceFormula: null, barcodeFormula: null },
      after: { formula: afterFormula, sourceFormula: null, barcodeFormula: null },
    };
    projection.invalidateProjectionMutations([{
      id: 'sheetTable.update',
      unitId: runtime.model.unitId,
      sheetId: sourceSheetId,
      params: { tableId: 'sales-table' },
      affectedRanges: [],
      structuralFormulaOwnerDeltas: [delta],
    }]);

    const after = projection.getCanvasProjection(owner);
    assert.notEqual(after, before);
    assert.equal(after.getCell(0, 0)?.formula, afterFormula);
  });

  it('invalidates cached conditional-format, drawing, and Table Sheet owners on their own sheet', () => {
    const app = new WorkbookSession();
    const sourceSheetId = app.getActiveSheetId();
    app.addSheet();
    const ownerSheetId = app.getActiveSheetId();
    const runtime = app['runtime'];
    const owner = runtime.model.getSheet(ownerSheetId);
    const projection = app['projection'];
    const beforeFormula = '=Sales[Amount]';
    const afterFormula = '=Orders[Amount]';
    const ranges = [{ sheetId: ownerSheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
    owner.conditionalFormats.push({
      id: 'linked-rule', sheetId: ownerSheetId, ranges, type: 'highlight', operator: 'greaterThan', value1: beforeFormula,
    });
    const shapeDrawing: DrawingObject = {
      id: 'linked-shape-drawing', sheetId: ownerSheetId, kind: 'shape', payloadId: 'linked-shape',
      anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 100, height: 40, rotation: 0 }, zIndex: 0,
    };
    const shapePayload: ShapeDrawingPayload = {
      kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#000000', propertyFormula: beforeFormula,
    };
    owner.drawings.push(shapeDrawing);
    owner.drawingPayloads.set(shapeDrawing.payloadId, shapePayload);
    owner.tableSheet = { viewId: 'sales-view', columns: [{ fieldId: 'amount', caption: 'Amount', formula: beforeFormula }], grouping: [] };

    const invalidate = (delta: StructuralFormulaOwnerDelta) => projection.invalidateProjectionMutations([{
      id: 'sheetTable.update',
      unitId: runtime.model.unitId,
      sheetId: sourceSheetId,
      params: { tableId: 'sales-table' },
      affectedRanges: [],
      structuralFormulaOwnerDeltas: [delta],
    }]);

    let before = projection.getCanvasProjection(owner);
    owner.conditionalFormats[0]!.value1 = afterFormula;
    invalidate({
      kind: 'formula-rule', sheetId: ownerSheetId, ruleKind: 'conditional-format', ruleId: 'linked-rule', field: 'value1',
      beforeFormula, afterFormula, beforeRanges: ranges, afterRanges: ranges,
    });
    let after = projection.getCanvasProjection(owner);
    assert.notEqual(after, before);
    assert.equal(after.conditionalFormats[0]?.value1, afterFormula);

    before = after;
    shapePayload.propertyFormula = afterFormula;
    invalidate({ kind: 'formula-object', ownerKind: 'shape-property', sheetId: ownerSheetId, payloadId: 'linked-shape', beforeFormula, afterFormula });
    after = projection.getCanvasProjection(owner);
    assert.notEqual(after, before);
    assert.equal((after.drawingPayloads.get('linked-shape') as ShapeDrawingPayload).propertyFormula, afterFormula);

    before = after;
    owner.tableSheet!.columns[0]!.formula = beforeFormula;
    invalidate({ kind: 'formula-object', ownerKind: 'table-sheet-column', sheetId: ownerSheetId, fieldId: 'amount', beforeFormula: afterFormula, afterFormula: beforeFormula });
    after = projection.getCanvasProjection(owner);
    assert.notEqual(after, before);
    assert.equal(after.tableSheet?.columns[0]?.formula, beforeFormula);
  });
});
