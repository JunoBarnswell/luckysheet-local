import { BARCODE_SYMBOLOGIES, type BarcodeCellPresentation, type CellData, type DataChartBindingArea, type DataChartDrawingPayload, type DrawingObject, type FormControlDrawingPayload, type ImageCellPresentation, type ImageDrawingPayload, type ImageEffects, type RangeRef, type SheetSnapshot, type WorkbookTableModel, type WorksheetModel } from '@react-sheets/core-model';
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

export interface PictureConvertToCellParams {
  sheetId: string;
  drawingId: string;
  row: number;
  column: number;
}

export interface PictureConvertToFloatingParams {
  sheetId: string;
  row: number;
  column: number;
  drawingId: string;
  payloadId: string;
  transform?: DrawingObject['transform'];
}

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
  validateBarcodePresentation(params.presentation);
  const sheet = context.workbook.getSheet(params.sheetId);
  if (!Array.isArray(params.ranges) || params.ranges.length === 0) throw new Error('Barcode apply requires at least one range');
  let mutationCount = 0;
  const affectedRanges = params.ranges.map((range) => {
    if (!range || range.sheetId !== params.sheetId || !Number.isInteger(range.startRow) || !Number.isInteger(range.endRow)
      || !Number.isInteger(range.startColumn) || !Number.isInteger(range.endColumn)
      || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn
      || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) {
      throw new Error('Barcode apply range is outside the worksheet bounds');
    }
    return { ...range, sheetId: params.sheetId };
  });
  for (const range of affectedRanges) {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const previous = structuredClone(sheet.cells.get(row, column));
        validateBarcodeValue(params.presentation, previous?.value ?? null);
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

function validateBarcodePresentation(presentation: BarcodeCellPresentation): void {
  if (!presentation || presentation.kind !== 'barcode' || !BARCODE_SYMBOLOGIES.includes(presentation.symbology)) throw new Error(`Unsupported barcode symbology: ${presentation?.symbology ?? 'missing'}`);
  if (!presentation.source || (presentation.source.kind !== 'cell-value' && presentation.source.kind !== 'formula')) throw new Error('Barcode source is invalid');
  if (presentation.source.kind === 'formula' && !/^\s*=/.test(presentation.source.formula)) throw new Error('Barcode formula source must be a formula');
  if (!presentation.options || typeof presentation.options.showText !== 'boolean' || !/^#[0-9a-f]{6}$/i.test(presentation.options.foreground) || !/^#[0-9a-f]{6}$/i.test(presentation.options.background)) throw new Error('Barcode options are invalid');
  if (!Number.isInteger(presentation.options.quietZone) || presentation.options.quietZone < 0 || presentation.options.quietZone > 20) throw new Error('Barcode quiet zone must be an integer from 0 to 20');
  if (!['above', 'below', 'none'].includes(presentation.options.labelPosition)) throw new Error('Barcode label position is invalid');
  if (presentation.options.fontSize !== undefined && (!Number.isFinite(presentation.options.fontSize) || presentation.options.fontSize < 6 || presentation.options.fontSize > 48)) throw new Error('Barcode label font size is invalid');
  if (!presentation.parameters || presentation.parameters.symbology !== presentation.symbology) throw new Error('Barcode parameters do not match the selected symbology');
  if (presentation.symbology === 'qr') {
    if ('errorCorrection' in presentation.parameters && presentation.parameters.errorCorrection !== undefined && !['low', 'medium', 'quartile', 'high'].includes(presentation.parameters.errorCorrection)) throw new Error('Barcode error correction is invalid');
  } else if (presentation.symbology === 'pdf417') {
    if ('securityLevel' in presentation.parameters && presentation.parameters.securityLevel !== undefined && (!Number.isInteger(presentation.parameters.securityLevel) || presentation.parameters.securityLevel < 0 || presentation.parameters.securityLevel > 8)) throw new Error('PDF417 security level must be an integer from 0 to 8');
  } else if (presentation.symbology === 'ean13' || presentation.symbology === 'ean8' || presentation.symbology === 'upca') {
    if ('addOnText' in presentation.parameters && presentation.parameters.addOnText !== undefined && !/^\d{2,5}$/.test(presentation.parameters.addOnText)) throw new Error('Barcode add-on text must contain 2 to 5 digits');
    if ('includeCheckDigit' in presentation.parameters && presentation.parameters.includeCheckDigit !== undefined && typeof presentation.parameters.includeCheckDigit !== 'boolean') throw new Error('Barcode check-digit option is invalid');
  } else {
    if ('wideNarrowRatio' in presentation.parameters && presentation.parameters.wideNarrowRatio !== undefined && (!Number.isFinite(presentation.parameters.wideNarrowRatio) || presentation.parameters.wideNarrowRatio < 1 || presentation.parameters.wideNarrowRatio > 4)) throw new Error('Barcode wide/narrow ratio is invalid');
    if ('fullAscii' in presentation.parameters && presentation.parameters.fullAscii !== undefined && typeof presentation.parameters.fullAscii !== 'boolean') throw new Error('Barcode full-ASCII option is invalid');
    if ('includeCheckDigit' in presentation.parameters && presentation.parameters.includeCheckDigit !== undefined && typeof presentation.parameters.includeCheckDigit !== 'boolean') throw new Error('Barcode check-digit option is invalid');
  }
}

function validateBarcodeValue(presentation: BarcodeCellPresentation, rawValue: CellData['value']): void {
  if (presentation.source.kind === 'formula') return;
  const value = rawValue == null ? '' : String(rawValue);
  if (!value.trim()) throw new Error(`Barcode source value is empty for ${presentation.symbology}`);
  const valid = presentation.symbology === 'ean13' ? /^\d{12,13}$/.test(value)
    : presentation.symbology === 'ean8' ? /^\d{7,8}$/.test(value)
      : presentation.symbology === 'upca' ? /^\d{11,12}$/.test(value)
        : presentation.symbology === 'code39' ? /^[0-9A-Z .$/+%\-]+$/.test(value)
          : presentation.symbology === 'codabar' ? /^[0-9A-D\-\$:/.+]+$/i.test(value)
            : value.length <= 4096;
  if (!valid) throw new Error(`Barcode source value is invalid for ${presentation.symbology}`);
}

function validateImagePresentation(presentation: ImageCellPresentation): void {
  if (!presentation || presentation.kind !== 'image' || typeof presentation.src !== 'string' || presentation.src.length === 0) throw new Error('Image presentation source is required');
  if (!['contain', 'cover', 'stretch'].includes(presentation.fit)) throw new Error('Image fit mode is invalid');
  const crop = presentation.crop;
  if (crop && (![crop.left, crop.top, crop.right, crop.bottom].every((value) => Number.isFinite(value) && value >= 0) || crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1)) throw new Error('Image crop is invalid');
  const effects = presentation.effects;
  if (effects && (!isImageEffects(effects))) throw new Error('Image effects are invalid');
}

function isImageEffects(value: ImageEffects): boolean {
  return (value.brightness === undefined || (Number.isFinite(value.brightness) && value.brightness >= -1 && value.brightness <= 1))
    && (value.contrast === undefined || (Number.isFinite(value.contrast) && value.contrast >= -1 && value.contrast <= 1))
    && (value.transparency === undefined || (Number.isFinite(value.transparency) && value.transparency >= 0 && value.transparency <= 1));
}

function imageRange(sheetId: string, row: number, column: number): RangeRef[] {
  return [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
}

function assertCellPosition(sheet: WorksheetModel, row: number, column: number): void {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0 || row >= sheet.rowCount || column >= sheet.columnCount) throw new Error('Picture cell position is outside worksheet bounds');
}

function removeDrawingForPicture(sheet: WorksheetModel, drawingId: string): void {
  const index = sheet.drawings.findIndex((drawing) => drawing.id === drawingId);
  if (index < 0) throw new Error(`Unknown drawing: ${drawingId}`);
  const drawing = sheet.drawings[index]!;
  sheet.drawings.splice(index, 1);
  sheet.drawingPayloads.delete(drawing.payloadId);
}

function convertCellImageToPayload(presentation: ImageCellPresentation): ImageDrawingPayload {
  validateImagePresentation(presentation);
  return {
    kind: 'image',
    src: presentation.src,
    altText: presentation.altText,
    crop: presentation.crop ? structuredClone(presentation.crop) : undefined,
    effects: presentation.effects ? structuredClone(presentation.effects) : undefined,
  };
}

function convertDrawingImageToPresentation(payload: ImageDrawingPayload): ImageCellPresentation {
  return {
    kind: 'image',
    src: payload.src,
    altText: payload.altText,
    fit: 'contain',
    crop: payload.crop ? structuredClone(payload.crop) : undefined,
    effects: payload.effects ? structuredClone(payload.effects) : undefined,
  };
}

function executePictureConvertToCell(params: PictureConvertToCellParams, context: CommandContext) {
  const sheet = context.workbook.getSheet(params.sheetId);
  assertCellPosition(sheet, params.row, params.column);
  const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
  const payload = drawing ? sheet.drawingPayloads.get(drawing.payloadId) : undefined;
  if (!drawing || drawing.kind !== 'image' || payload?.kind !== 'image') throw new Error(`Unknown image drawing: ${params.drawingId}`);
  const previousCell = structuredClone(sheet.cells.get(params.row, params.column));
  if (previousCell && (previousCell.value !== null || previousCell.presentation !== undefined)) throw new Error('Picture conversion target cell is not empty');
  const nextCell: CellData = { ...(previousCell ?? { value: null }), presentation: convertDrawingImageToPresentation(payload) };
  const affectedRanges = imageRange(params.sheetId, params.row, params.column);
  const drawingSnapshot = structuredClone(drawing);
  const payloadSnapshot = structuredClone(payload);
  context.applyMutation({
    id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId,
    params: { sheetId: params.sheetId, drawingId: params.drawingId }, affectedRanges,
    inverse: [{ id: 'drawing.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawing: drawingSnapshot, payload: payloadSnapshot }, affectedRanges }],
    apply: () => removeDrawingForPicture(sheet, params.drawingId),
  });
  context.applyMutation({
    id: 'cell.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
    params: { sheetId: params.sheetId, row: params.row, column: params.column, value: nextCell }, affectedRanges,
    inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, previous: previousCell }, affectedRanges }],
    apply: () => sheet.cells.set(params.row, params.column, structuredClone(nextCell)),
  });
  return { operationId: context.operationId, mutationCount: 2, affectedRanges };
}

function executePictureConvertToFloating(params: PictureConvertToFloatingParams, context: CommandContext) {
  const sheet = context.workbook.getSheet(params.sheetId);
  assertCellPosition(sheet, params.row, params.column);
  if (sheet.drawings.some((drawing) => drawing.id === params.drawingId)) throw new Error(`Drawing already exists: ${params.drawingId}`);
  if (sheet.drawingPayloads.has(params.payloadId)) throw new Error(`Drawing payload already exists: ${params.payloadId}`);
  const previousCell = structuredClone(sheet.cells.get(params.row, params.column));
  if (!previousCell?.presentation || previousCell.presentation.kind !== 'image') throw new Error('Picture conversion source cell has no image');
  const payload = convertCellImageToPayload(previousCell.presentation);
  const drawing: DrawingObject = {
    id: params.drawingId,
    sheetId: params.sheetId,
    kind: 'image',
    anchor: { kind: 'one-cell', row: params.row, column: params.column },
    transform: structuredClone(params.transform ?? { x: 96, y: 96, width: 320, height: 200, rotation: 0 }),
    zIndex: 0,
    payloadId: params.payloadId,
  };
  const nextCell = structuredClone(previousCell);
  delete nextCell.presentation;
  const affectedRanges = imageRange(params.sheetId, params.row, params.column);
  context.applyMutation({
    id: 'cell.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
    params: { sheetId: params.sheetId, row: params.row, column: params.column, value: nextCell }, affectedRanges,
    inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, previous: previousCell }, affectedRanges }],
    apply: () => sheet.cells.set(params.row, params.column, structuredClone(nextCell)),
  });
  context.applyMutation({
    id: 'drawing.add', unitId: context.workbook.unitId, sheetId: params.sheetId,
    params: { sheetId: params.sheetId, drawing, payload }, affectedRanges,
    inverse: [{ id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: drawing.id }, affectedRanges }],
    apply: () => {
      sheet.drawings.push(structuredClone(drawing));
      sheet.drawingPayloads.set(drawing.payloadId, structuredClone(payload));
    },
  });
  return { operationId: context.operationId, mutationCount: 2, affectedRanges };
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
      assertCellPosition(sheet, params.row, params.column);
      validateImagePresentation(params.presentation);
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
  runtime.registry.registerCommand<PictureConvertToCellParams>({ id: 'picture.convertToCell', execute: executePictureConvertToCell });
  runtime.registry.registerCommand<PictureConvertToFloatingParams>({ id: 'picture.convertToFloating', execute: executePictureConvertToFloating });
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
  return ['sheet.create.advanced', 'cell.barcode.apply', 'dataChart.create', 'dataChart.update', 'cell.image.apply', 'picture.convertToCell', 'picture.convertToFloating', 'formControl.activate'];
}
