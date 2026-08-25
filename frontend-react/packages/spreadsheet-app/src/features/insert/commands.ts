import type { BarcodeCellPresentation, CellData, DataChartBindingArea, DataChartDrawingPayload, DrawingObject, FormControlDrawingPayload, ImageCellPresentation, RangeRef, SheetSnapshot, WorkbookTableModel, WorksheetModel } from '@react-sheets/core-model';
import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';

export interface AdvancedSheetCreateParams {
  sheet: SheetSnapshot;
  index?: number;
  table?: WorkbookTableModel;
}

export interface BarcodeApplyParams {
  sheetId: string;
  ranges: RangeRef[];
  presentation: BarcodeCellPresentation;
}

export interface DataChartCreateParams {
  sheetId: string;
  drawing: DrawingObject;
  payload: DataChartDrawingPayload;
  table?: WorkbookTableModel;
}

export interface DataChartUpdateParams {
  sheetId: string;
  drawingId: string;
  payload: DataChartDrawingPayload;
}

export interface CellImageApplyParams { sheetId: string; row: number; column: number; presentation: ImageCellPresentation }

export interface FormControlActivateParams { sheetId: string; drawingId: string }

function executeAdvancedSheetCreate(params: AdvancedSheetCreateParams, context: CommandContext) {
  if (!['table-sheet', 'gantt-sheet', 'report-sheet'].includes(params.sheet.kind)) throw new Error('Advanced sheet kind is invalid');
  if (context.workbook.sheets.has(params.sheet.id)) throw new Error(`Sheet already exists: ${params.sheet.id}`);
  const affectedRanges: RangeRef[] = [];
  const table = params.table;
  const addsTable = Boolean(table && !context.workbook.dataModel.tables.has(table.id));
  if (addsTable && table) {
    const sourceRanges = table.sourceRange ? [table.sourceRange] : [];
    context.applyMutation({
      id: 'table.add', unitId: context.workbook.unitId, sheetId: table.sourceSheetId ?? params.sheet.id, params: table,
      affectedRanges: sourceRanges,
      inverse: [{ id: 'table.remove', unitId: context.workbook.unitId, sheetId: table.sourceSheetId ?? params.sheet.id, params: { tableId: table.id, range: table.sourceRange }, affectedRanges: sourceRanges }],
      apply: () => context.workbook.addTable(table),
    });
  }
  context.applyMutation({
    id: 'sheet.restore', unitId: context.workbook.unitId, sheetId: params.sheet.id, params: { sheet: structuredClone(params.sheet), index: params.index }, affectedRanges,
    inverse: [{ id: 'sheet.remove', unitId: context.workbook.unitId, sheetId: params.sheet.id, params: { id: params.sheet.id }, affectedRanges }],
    apply: () => context.workbook.restoreSheetSnapshot(params.sheet, params.index),
  });
  return { operationId: context.operationId, mutationCount: addsTable ? 2 : 1, affectedRanges };
}

function executeDataChartCreate(params: DataChartCreateParams, context: CommandContext) {
  if (params.drawing.kind !== 'data-chart' || params.payload.kind !== 'data-chart') throw new Error('DataChart drawing and payload kinds must match');
  const boundTable = params.payload.source.kind === 'table' ? params.table ?? context.workbook.dataModel.tables.get(params.payload.source.tableId) : undefined;
  if (params.payload.source.kind === 'table' && !boundTable) throw new Error(`Data Chart table binding not found: ${params.payload.source.tableId}`);
  validateDataChartPayload(params.payload, boundTable);
  const affectedRanges: RangeRef[] = params.payload.source.kind === 'report-sheet'
    ? [params.payload.source.range]
    : params.table?.sourceRange ? [params.table.sourceRange] : [{ sheetId: params.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
  const table = params.table;
  const addsTable = Boolean(table && !context.workbook.dataModel.tables.has(table.id));
  if (addsTable && table) {
    context.applyMutation({ id: 'table.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: table, affectedRanges,
      inverse: [{ id: 'table.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { tableId: table.id, range: table.sourceRange }, affectedRanges }],
      apply: () => context.workbook.addTable(table),
    });
  }
  context.applyMutation({
    id: 'drawing.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawing: params.drawing, payload: params.payload }, affectedRanges,
    inverse: [{ id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: params.drawing.id }, affectedRanges }],
    apply: () => {
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.drawings.push(structuredClone(params.drawing));
      sheet.drawingPayloads.set(params.drawing.payloadId, structuredClone(params.payload));
    },
  });
  return { operationId: context.operationId, mutationCount: addsTable ? 2 : 1, affectedRanges };
}

const DATA_CHART_AREAS: readonly DataChartBindingArea[] = ['values', 'category', 'details', 'color', 'size', 'tooltip', 'filter'];
const DATA_CHART_AGGREGATES = new Set(['sum', 'average', 'count', 'min', 'max', 'none']);
const DATA_CHART_PLOT_TYPES = new Set(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar', 'treemap', 'funnel']);

function validateDataChartPayload(payload: DataChartDrawingPayload, table?: WorkbookTableModel): void {
  if (payload.kind !== 'data-chart' || !DATA_CHART_PLOT_TYPES.has(payload.plotType)) throw new Error('Data Chart payload plot type is invalid');
  if (payload.source.kind === 'table') {
    if (!payload.source.tableId.trim()) throw new Error('Data Chart table binding is required');
    if (table && table.id !== payload.source.tableId) throw new Error('Data Chart table binding does not match the inserted table');
  } else {
    const range = payload.source.range;
    if (!range.sheetId || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn) throw new Error('Data Chart report-sheet binding range is invalid');
  }
  if (!payload.bindings || typeof payload.bindings !== 'object') throw new Error('Data Chart bindings are required');
  for (const area of DATA_CHART_AREAS) {
    const entries = payload.bindings[area];
    if (!Array.isArray(entries)) throw new Error(`Data Chart binding area is missing: ${area}`);
    for (const entry of entries) {
      if (!entry || typeof entry.fieldId !== 'string' || !entry.fieldId.trim() || entry.area !== area || !DATA_CHART_AGGREGATES.has(entry.aggregate)) throw new Error(`Data Chart binding is invalid: ${area}`);
      if (entry.sort !== undefined && entry.sort !== 'asc' && entry.sort !== 'desc') throw new Error(`Data Chart sort is invalid: ${area}`);
    }
  }
  const inspector = payload.inspector;
  if (!inspector || !['top', 'bottom', 'left', 'right', 'none'].includes(inspector.legendPosition) || typeof inspector.showDataLabels !== 'boolean') throw new Error('Data Chart inspector configuration is invalid');
  if (!inspector.chartArea || !inspector.plotArea || !inspector.axis || typeof inspector.axis.showGridlines !== 'boolean') throw new Error('Data Chart inspector style is invalid');
}

function findDataChart(sheet: WorksheetModel, drawingId: string): { drawing: DrawingObject; payload: DataChartDrawingPayload } {
  const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
  const payload = drawing ? sheet.drawingPayloads.get(drawing.payloadId) : undefined;
  if (!drawing || drawing.kind !== 'data-chart' || payload?.kind !== 'data-chart') throw new Error(`Unknown Data Chart drawing: ${drawingId}`);
  return { drawing, payload };
}

function executeDataChartUpdate(params: DataChartUpdateParams, context: CommandContext) {
  const sheet = context.workbook.getSheet(params.sheetId);
  const current = findDataChart(sheet, params.drawingId);
  const boundTable = params.payload.source.kind === 'table' ? context.workbook.dataModel.tables.get(params.payload.source.tableId) : undefined;
  if (params.payload.source.kind === 'table' && !boundTable) throw new Error(`Data Chart table binding not found: ${params.payload.source.tableId}`);
  validateDataChartPayload(params.payload, boundTable);
  const affectedRanges: RangeRef[] = params.payload.source.kind === 'report-sheet'
    ? [params.payload.source.range]
    : (() => {
      const table = context.workbook.dataModel.tables.get(params.payload.source.tableId);
      return table?.sourceRange ? [table.sourceRange] : [{ sheetId: params.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
    })();
  context.applyMutation({
    id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId,
    params: { sheetId: params.sheetId, payloadId: current.drawing.payloadId, before: current.payload, after: params.payload }, affectedRanges,
    inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: current.drawing.payloadId, before: params.payload, after: current.payload }, affectedRanges }],
    apply: () => sheet.drawingPayloads.set(current.drawing.payloadId, structuredClone(params.payload)),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeBarcodeApply(params: BarcodeApplyParams, context: CommandContext) {
  const sheet = context.workbook.getSheet(params.sheetId);
  let mutationCount = 0;
  const affectedRanges = params.ranges.map((range) => ({ ...range, sheetId: params.sheetId }));
  for (const range of affectedRanges) {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const previous = structuredClone(sheet.cells.get(row, column));
        const next: CellData = structuredClone(previous ?? { value: null });
        next.presentation = structuredClone(params.presentation);
        const cellRange: RangeRef[] = [{ sheetId: params.sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
        context.applyMutation({
          id: 'cell.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row, column, value: next }, affectedRanges: cellRange,
          inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row, column, previous }, affectedRanges: cellRange }],
          apply: () => sheet.cells.set(row, column, structuredClone(next)),
        });
        mutationCount += 1;
      }
    }
  }
  return { operationId: context.operationId, mutationCount, affectedRanges };
}

export function registerInsertCommands(runtime: CommandRuntime): string[] {
  runtime.registry.registerCommand<AdvancedSheetCreateParams>({ id: 'sheet.create.advanced', execute: executeAdvancedSheetCreate });
  runtime.registry.registerCommand<BarcodeApplyParams>({ id: 'cell.barcode.apply', execute: executeBarcodeApply });
  runtime.registry.registerCommand<DataChartCreateParams>({ id: 'dataChart.create', execute: executeDataChartCreate });
  runtime.registry.registerCommand<DataChartUpdateParams>({ id: 'dataChart.update', execute: executeDataChartUpdate });
  runtime.registry.registerCommand<CellImageApplyParams>({
    id: 'cell.image.apply',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = structuredClone(sheet.cells.get(params.row, params.column));
      const next: CellData = structuredClone(previous ?? { value: null });
      next.presentation = structuredClone(params.presentation);
      const affectedRanges = [{ sheetId: params.sheetId, startRow: params.row, endRow: params.row, startColumn: params.column, endColumn: params.column }];
      context.applyMutation({ id: 'cell.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, value: next }, affectedRanges,
        inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, previous }, affectedRanges }],
        apply: () => sheet.cells.set(params.row, params.column, structuredClone(next)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<FormControlActivateParams>({
    id: 'formControl.activate',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
      const before = drawing ? sheet.drawingPayloads.get(drawing.payloadId) : undefined;
      if (!drawing || before?.kind !== 'form-control' || !before.enabled) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const after: FormControlDrawingPayload = structuredClone(before);
      if (after.controlType === 'checkbox' || after.controlType === 'option-button') after.value = !Boolean(after.value);
      else if (after.controlType === 'spin-button' || after.controlType === 'scrollbar') after.value = Number(after.value ?? 0) + 1;
      else if ((after.controlType === 'list-box' || after.controlType === 'combo-box') && after.inputRange) {
        const values: string[] = [];
        const source = context.workbook.getSheet(after.inputRange.sheetId);
        for (let row = after.inputRange.startRow; row <= after.inputRange.endRow; row += 1) for (let column = after.inputRange.startColumn; column <= after.inputRange.endColumn; column += 1) {
          const value = source.cells.get(row, column)?.value;
          if (value != null) values.push(String(value));
        }
        if (values.length) after.value = values[(Math.max(-1, values.indexOf(String(after.value ?? ''))) + 1) % values.length]!;
      } else return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = after.cellLink ? [{ sheetId: after.cellLink.sheetId, startRow: after.cellLink.row, endRow: after.cellLink.row, startColumn: after.cellLink.column, endColumn: after.cellLink.column }] : [{ sheetId: params.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
      context.applyMutation({ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after }, affectedRanges,
        inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before: after, after: before }, affectedRanges }],
        apply: () => sheet.drawingPayloads.set(drawing.payloadId, structuredClone(after)),
      });
      let mutationCount = 1;
      if (after.cellLink) {
        const linkedSheet = context.workbook.getSheet(after.cellLink.sheetId);
        const previous = structuredClone(linkedSheet.cells.get(after.cellLink.row, after.cellLink.column));
        const next: CellData = { ...(previous ?? { value: null }), value: after.value as CellData['value'] };
        context.applyMutation({ id: 'cell.set', unitId: context.workbook.unitId, sheetId: after.cellLink.sheetId, params: { sheetId: after.cellLink.sheetId, row: after.cellLink.row, column: after.cellLink.column, value: next }, affectedRanges,
          inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: after.cellLink.sheetId, params: { sheetId: after.cellLink.sheetId, row: after.cellLink.row, column: after.cellLink.column, previous }, affectedRanges }],
          apply: () => linkedSheet.cells.set(after.cellLink!.row, after.cellLink!.column, structuredClone(next)),
        });
        mutationCount += 1;
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });
  return ['sheet.create.advanced', 'cell.barcode.apply', 'dataChart.create', 'dataChart.update', 'cell.image.apply', 'formControl.activate'];
}
