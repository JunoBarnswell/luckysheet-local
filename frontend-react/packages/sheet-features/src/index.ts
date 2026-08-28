import type {
  BandedRule,
  CellData,
  RichTextRun,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
  AutoFilterModel,
  WorksheetPane,
  MergeSpan,
  RangeRef,
  WorkbookTableModel,
  WorkbookModel,
  WorksheetModel,
  StructuralTransformParams,
  DefinedNameModel,
  TableSheetDefinition,
  GanttSheetDefinition,
  ReportSheetDefinition,
} from '@react-sheets/core-model';
import {
  clearFormulaProvenance,
  MAX_SHEET_COLUMN_COUNT,
  MAX_SHEET_ROW_COUNT,
  StructuralTransform,
  normalizeDefinedNameModel,
  normalizeFontFamily,
  planBorderChange,
  isBorderPlacement,
  isBorderLine,
  type BorderPlacement,
  type BorderLine,
} from '@react-sheets/core-model';
import { isHorizontalAlignment, isReadingOrder, isVerticalAlignment } from '@react-sheets/core-model';
import type { CommandContext, CommandResult, CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import { isSpillChild } from '@react-sheets/formula-engine';
import { buildCellFromText, type CellInputInterpretationContext } from './text-input';
import { registerEditingCommands } from './editing';
import { registerDataToolCommands, findValidationRule, normalizeConditionalFormatRule, normalizeDataValidationRule, validateDataInput } from './data-features';
import { registerSheetTableCommands } from './sheet-table-commands';
import { planSheetTableAutoExpansion, validateFilterOwnership } from './sheet-table-features';
import { registerOutlineCommands } from './outline-commands';
import { registerHomeCommands } from './home-commands';
import { registerPhoneticCommands } from './phonetic-commands';
import { normalizeCheckboxCellValue, registerCellTemplateCommands } from './cell-template-commands';
import { applyClearRangePlan, createClearRangePlan, restoreClearRangeSnapshot, type ClearRangeParams, type ClearRangeSnapshot } from './clear-planner';
import { assertCellWriteAuthority, createCellSetMutationParams, isCellSetMutationParams, type CellSetMutationParams } from './cell-write-authority';
import { CellEntryError } from './cell-entry-error';

function snapshotCellRegion(
  sheet: WorksheetModel,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): Array<{ row: number; column: number; cell: CellData }> {
  const extracted: Array<{ row: number; column: number; cell: CellData }> = [];
  sheet.cells.forEach((cell, row, column) => {
    if (row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) {
      extracted.push({ row, column, cell: structuredClone(cell) });
    }
  });
  return extracted;
}

function applyStructuralTransform(workbook: WorkbookModel, params: StructuralTransformParams): void {
  StructuralTransform.apply(workbook, params);
}

export * from './clipboard';
export * from './data-features';
export * from './editing';
export * from './sheet-table-features';
export * from './sheet-table-commands';
export * from './outline-commands';
export * from './outline-features';
export * from './text-input';
export * from './home-commands';
export * from './phonetic-commands';
export * from './auto-sum-contract';
export * from './fill-series';
export * from './flash-fill';
export * from './find-replace';
export * from './cell-template-commands';
export * from './clear-planner';
export * from './data-region-context';
export * from './cell-write-authority';
export * from './cell-entry-error';


export interface SetCellValueParams {
  sheetId: string;
  row: number;
  column: number;
  value: CellData;
}

/** Host/UI text commit contract. The command owns parsing, validation and
 * content/presentation merging; callers must pass the raw text unchanged. */
export interface CommitTextParams {
  sheetId: string;
  row: number;
  column: number;
  text: string;
  inputContext: CellInputInterpretationContext;
  style?: Partial<CellStyle>;
  /** Required only when a warning/information validation rule is overridden. */
  validationConfirmation?: boolean;
}

export interface CommitTypedValueParams {
  sheetId: string;
  row: number;
  column: number;
  value: CellData['value'];
  validationConfirmation?: boolean;
}

export interface CommitRichTextParams {
  sheetId: string;
  row: number;
  column: number;
  text: string;
  runs: RichTextRun[];
  validationConfirmation?: boolean;
}

export interface CommitTextCellsParams {
  text: string;
  targets: Array<{
    sheetId: string;
    row: number;
    column: number;
    inputContext: CellInputInterpretationContext;
    style?: Partial<CellStyle>;
  }>;
  validationConfirmation?: boolean;
}

export interface CommitTypedValueCellsParams {
  value: CellData['value'];
  targets: Array<{ sheetId: string; row: number; column: number }>;
  validationConfirmation?: boolean;
}

export interface CommitRichTextCellsParams {
  text: string;
  runs: RichTextRun[];
  targets: Array<{ sheetId: string; row: number; column: number }>;
  validationConfirmation?: boolean;
}

export interface AddTableParams extends WorkbookTableModel {}

export interface TableSheetUpdateParams {
  sheetId: string;
  definition: TableSheetDefinition;
}

export interface GanttSheetUpdateParams {
  sheetId: string;
  definition: GanttSheetDefinition;
}

export interface ReportSheetUpdateParams {
  sheetId: string;
  definition: ReportSheetDefinition;
}

export interface SetRangeValuesParams {
  sheetId: string;
  startRow: number;
  startColumn: number;
  values: CellData[][];
}

interface ClearRangeRestoreParams {
  sheetId: string;
  range: RangeRef;
  snapshot: ClearRangeSnapshot;
}

export interface AddSheetParams {
  id: string;
  name: string;
  rowCount?: number;
  columnCount?: number;
}

export interface SheetExtentParams {
  sheetId: string;
  rowCount: number;
  columnCount: number;
}

export interface RenameSheetParams {
  sheetId: string;
  name: string;
}

export interface RenameWorkbookParams { name: string; }

export interface InsertRowParams {
  sheetId: string;
  row: number;
  count: number;
}

export interface DeleteRowParams {
  sheetId: string;
  row: number;
  count: number;
}

export interface ResizeRowParams {
  sheetId: string;
  row: number;
  heightPx: number;
}

export interface InsertColumnParams {
  sheetId: string;
  column: number;
  count: number;
}

export interface DeleteColumnParams {
  sheetId: string;
  column: number;
  count: number;
}

export interface ResizeColumnParams {
  sheetId: string;
  column: number;
  widthPx: number;
}

export interface ColumnsVisibilityParams {
  sheetId: string;
  states: Array<{ column: number; hidden: boolean }>;
}

export interface RowsVisibilityParams {
  sheetId: string;
  states: Array<{ row: number; hidden: boolean }>;
}

export interface SetMergeParams {
  sheetId: string;
  range: RangeRef;
}

export interface RemoveMergeParams {
  sheetId: string;
  range: RangeRef;
}

export interface SetFreezeParams {
  sheetId: string;
  pane: WorksheetPane;
}

export interface SetRangeStyleParams {
  sheetId: string;
  range: RangeRef;
  style: Partial<CellStyle>;
  /** Canonical authored number format stored on CellData. */
  numberFormat?: string;
  /** Replace the authored style instead of merging a patch. */
  replaceStyle?: boolean;
  /** Remove the top-level number format when the source has no format. */
  clearNumberFormat?: boolean;
}

/** Canonical topology command shared by HOME and Format Cells. */
export interface SetBorderParams {
  sheetId: string;
  range?: RangeRef;
  ranges?: RangeRef[];
  placement: BorderPlacement;
  line?: BorderLine;
}

export interface SortRangeParams {
  sheetId: string;
  range: RangeRef;
  sortColumn: number;
  ascending: boolean;
  hasHeader?: boolean;
}

export interface AddConditionalFormatParams {
  sheetId: string;
  rule: ConditionalFormatRule;
}

export interface UpdateConditionalFormatParams {
  sheetId: string;
  ruleId: string;
  patch: Partial<ConditionalFormatRule>;
}

interface ConditionalFormatUpdateMutationParams {
  sheetId: string;
  before: ConditionalFormatRule;
  after: ConditionalFormatRule;
  ranges: RangeRef[];
}

export interface AddDataValidationParams {
  sheetId: string;
  rule: DataValidationRule;
}

function cellRange(params: Pick<SetCellValueParams, 'sheetId' | 'row' | 'column'>): RangeRef[] {
  return [
    {
      sheetId: params.sheetId,
      startRow: params.row,
      endRow: params.row,
      startColumn: params.column,
      endColumn: params.column,
    },
  ];
}

interface CellEntryTarget {
  sheetId: string;
  row: number;
  column: number;
  validationConfirmation?: boolean;
}

function rejectCellEntry(target: Pick<CellEntryTarget, 'sheetId' | 'row' | 'column'>, message: string, recovery: string): never {
  throw new CellEntryError({ code: 'CELL_ENTRY_INVALID_INPUT', message, ...target, recovery });
}

interface PreparedCellEntry {
  params: CellEntryTarget;
  next: CellData;
  previous: CellData | undefined;
  sheet: WorksheetModel;
  affectedRanges: RangeRef[];
}

function prepareCellEntry(
  params: CellEntryTarget,
  next: CellData,
  context: CommandContext,
): PreparedCellEntry {
  if (!Number.isSafeInteger(params.row) || params.row < 0 || !Number.isSafeInteger(params.column) || params.column < 0) {
    throw new CellEntryError({
      code: 'CELL_ENTRY_INVALID_INPUT',
      message: 'Cell row and column must be non-negative integers',
      sheetId: params.sheetId,
      row: params.row,
      column: params.column,
      recovery: 'Resolve a canonical non-negative cell address before committing.',
    });
  }
  const sheet = context.workbook.getSheet(params.sheetId);
  assertCanonicalCheckboxCell(next);
  for (const spill of sheet.spillRanges) {
    if (isSpillChild(spill, params.row, params.column)) {
      throw new CellEntryError({
        code: 'CELL_ENTRY_SPILL_CHILD',
        message: 'Cannot edit a dynamic-array spill child',
        sheetId: params.sheetId,
        row: params.row,
        column: params.column,
        recovery: 'Edit the dynamic-array source formula instead of a spill child.',
      });
    }
  }
  if (!next.formula) {
    const validation = validateDataInput(sheet, params.row, params.column, next.value);
    const rule = validation.ruleId ? findValidationRule(sheet, params.row, params.column) : undefined;
    if (validation.blocking) {
      throw new CellEntryError({
        code: 'CELL_ENTRY_VALIDATION_BLOCKED',
        message: validation.message ?? 'Cell value failed data validation',
        sheetId: params.sheetId,
        row: params.row,
        column: params.column,
        recovery: 'Correct the draft so it satisfies the target validation rule.',
        alertStyle: 'stop',
        ...(rule?.errorTitle ? { title: rule.errorTitle } : {}),
      });
    }
    if (!validation.valid && params.validationConfirmation !== true) {
      const alertStyle = validation.alertStyle === 'information' ? 'information' : 'warning';
      throw new CellEntryError({
        code: 'CELL_ENTRY_CONFIRMATION_REQUIRED',
        message: validation.message ?? 'Cell value requires explicit validation confirmation',
        sheetId: params.sheetId,
        row: params.row,
        column: params.column,
        recovery: 'Confirm the warning/information decision or return to the active edit session.',
        alertStyle,
        ...(rule?.errorTitle ? { title: rule.errorTitle } : {}),
      });
    }
  }

  return {
    params,
    next: structuredClone(next),
    previous: sheet.cells.get(params.row, params.column),
    sheet,
    affectedRanges: cellRange(params),
  };
}

function applyPreparedCellEntry(prepared: PreparedCellEntry, context: CommandContext): void {
  const { params, next, previous, sheet, affectedRanges } = prepared;
  context.applyMutation({
    id: 'cell.set',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: createCellSetMutationParams(sheet, {
      sheetId: params.sheetId,
      row: params.row,
      column: params.column,
      value: structuredClone(next),
    }, 'direct-entry', params.validationConfirmation === true),
    affectedRanges,
    inverse: [{
      id: 'cell.restore',
      unitId: context.workbook.unitId,
      sheetId: params.sheetId,
      params: {
        sheetId: params.sheetId,
        row: params.row,
        column: params.column,
        previous: previous ? structuredClone(previous) : undefined,
      },
      affectedRanges,
    }],
    apply: () => sheet.cells.set(params.row, params.column, structuredClone(next)),
  });
}

function commitCellEntry(
  params: CellEntryTarget,
  next: CellData,
  context: CommandContext,
): CommandResult {
  const prepared = prepareCellEntry(params, next, context);
  applyPreparedCellEntry(prepared, context);
  return { operationId: context.operationId, mutationCount: 1, affectedRanges: prepared.affectedRanges };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isWorksheetPane(value: unknown): value is WorksheetPane {
  if (!isRecord(value) || !['none', 'frozen', 'split'].includes(String(value.kind))) return false;
  if (value.kind === 'none') return true;
  return typeof value.xSplit === 'number' && Number.isFinite(value.xSplit) && value.xSplit >= 0
    && typeof value.ySplit === 'number' && Number.isFinite(value.ySplit) && value.ySplit >= 0
    && Number.isSafeInteger(value.startRow) && Number(value.startRow) >= 0
    && Number.isSafeInteger(value.startColumn) && Number(value.startColumn) >= 0;
}

function isColumnVisibilityMutation(value: unknown): value is ColumnsVisibilityParams {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.states) && value.states.length > 0
    && value.states.every((state) => isRecord(state) && Number.isSafeInteger(state.column) && Number(state.column) >= 0 && typeof state.hidden === 'boolean');
}

function isRowVisibilityMutation(value: unknown): value is RowsVisibilityParams {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.states) && value.states.length > 0
    && value.states.every((state) => isRecord(state) && Number.isSafeInteger(state.row) && Number(state.row) >= 0 && typeof state.hidden === 'boolean');
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && Number.isInteger(value.startRow) && Number.isInteger(value.endRow)
    && Number.isInteger(value.startColumn) && Number.isInteger(value.endColumn)
    && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
    && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
}

function isCellData(value: unknown): value is CellData {
  return isRecord(value)
    && ('value' in value || typeof value.formula === 'string');
}

function isCellSetMutation(value: unknown): value is SetCellValueParams {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && Number.isInteger(value.row) && Number(value.row) >= 0
    && Number.isInteger(value.column) && Number(value.column) >= 0
    && isCellData(value.value);
}

function isCellRestoreMutation(value: unknown): value is { sheetId: string; row: number; column: number; previous?: CellData } {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && Number.isInteger(value.row) && Number(value.row) >= 0
    && Number.isInteger(value.column) && Number(value.column) >= 0
    && (value.previous === undefined || isCellData(value.previous));
}

function sheetScopeRange(value: { sheetId: string }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function isRenameWorkbookMutation(value: unknown): value is RenameWorkbookParams {
  return isRecord(value) && typeof value.name === 'string' && value.name.trim().length > 0;
}

function isRenameSheetMutation(value: unknown): value is RenameSheetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.name === 'string' && value.name.trim().length > 0;
}

function isAddSheetMutation(value: unknown): value is AddSheetParams {
  return isRecord(value) && typeof value.id === 'string' && value.id.length > 0 && typeof value.name === 'string'
    && (value.rowCount === undefined || (Number.isInteger(value.rowCount) && Number(value.rowCount) > 0))
    && (value.columnCount === undefined || (Number.isInteger(value.columnCount) && Number(value.columnCount) > 0));
}

function isSheetIdMutation(value: unknown): value is { id: string } {
  return isRecord(value) && typeof value.id === 'string' && value.id.length > 0;
}

function isSheetExtentMutation(value: unknown): value is SheetExtentParams {
  return isRecord(value)
    && typeof value.sheetId === 'string' && value.sheetId.length > 0
    && Number.isSafeInteger(value.rowCount) && Number(value.rowCount) > 0 && Number(value.rowCount) <= MAX_SHEET_ROW_COUNT
    && Number.isSafeInteger(value.columnCount) && Number(value.columnCount) > 0 && Number(value.columnCount) <= MAX_SHEET_COLUMN_COUNT;
}

function isSheetRestoreMutation(value: unknown): value is { sheet: import('@react-sheets/core-model').SheetSnapshot; index?: number } {
  return isRecord(value) && isRecord(value.sheet) && typeof value.sheet.id === 'string' && typeof value.sheet.name === 'string'
    && (value.index === undefined || (Number.isSafeInteger(value.index) && Number(value.index) >= 0));
}

function isWorkbookTableMutation(value: unknown): value is WorkbookTableModel {
  return isRecord(value) && typeof value.id === 'string' && typeof value.name === 'string'
    && Number.isInteger(value.rowCount) && Number(value.rowCount) >= 0
    && Number.isInteger(value.blockSize) && Number(value.blockSize) > 0
    && Number.isInteger(value.revision) && Number(value.revision) >= 0
    && Array.isArray(value.fields) && Array.isArray(value.blocks);
}

function isTableRemoveMutation(value: unknown): value is { tableId: string; range?: RangeRef } {
  return isRecord(value) && typeof value.tableId === 'string' && (value.range === undefined || isRange(value.range));
}

function workbookTableRanges(value: WorkbookTableModel): RangeRef[] {
  return value.sourceRange ? [structuredClone(value.sourceRange)] : [];
}

function tableRemoveRanges(value: { range?: RangeRef }): RangeRef[] {
  return value.range ? [structuredClone(value.range)] : [];
}

function isSetRangeMutation(value: unknown): value is SetRangeValuesParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string'
    || !Number.isInteger(value.startRow) || Number(value.startRow) < 0
    || !Number.isInteger(value.startColumn) || Number(value.startColumn) < 0
    || !Array.isArray(value.values)) return false;
  return value.values.every((row) => Array.isArray(row) && row.every(isCellData));
}

function setRangeAffectedRanges(value: SetRangeValuesParams): RangeRef[] {
  const rowCount = value.values.length;
  const columnCount = Math.max(1, ...value.values.map((row) => row.length));
  return [{
    sheetId: value.sheetId,
    startRow: value.startRow,
    endRow: value.startRow + Math.max(0, rowCount - 1),
    startColumn: value.startColumn,
    endColumn: value.startColumn + columnCount - 1,
  }];
}

function isClearRangeMutation(value: unknown): value is ClearRangeParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range)
    && (value.family === 'all' || value.family === 'contents' || value.family === 'formats' || value.family === 'comments-and-notes' || value.family === 'hyperlinks');
}

function isClearRangeRestoreMutation(value: unknown): value is ClearRangeRestoreParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range)
    && isRecord(value.snapshot)
    && Array.isArray(value.snapshot.cells) && value.snapshot.cells.every((entry) => isRecord(entry) && Number.isInteger(entry.row) && Number.isInteger(entry.column) && (entry.value === undefined || isCellData(entry.value)))
    && Array.isArray(value.snapshot.notes) && Array.isArray(value.snapshot.hyperlinks) && Array.isArray(value.snapshot.comments)
    && (value.snapshot.conditionalFormats === undefined || Array.isArray(value.snapshot.conditionalFormats))
    && (value.snapshot.dataValidations === undefined || Array.isArray(value.snapshot.dataValidations));
}

function isStyleMutation(value: unknown): value is SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string } {
  if (!isRecord(value) || typeof value.sheetId !== 'string') return false;
  const flagsValid = (value.replaceStyle === undefined || typeof value.replaceStyle === 'boolean')
    && (value.clearNumberFormat === undefined || typeof value.clearNumberFormat === 'boolean');
  if (isRange(value.range)) return flagsValid
    && (value.style === undefined || isCellStylePatch(value.style))
    && (value.numberFormat === undefined || typeof value.numberFormat === 'string');
  return Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange)
    && flagsValid && typeof value.numberFormat === 'string' && (value.style === undefined || isCellStylePatch(value.style));
}

function isCellStylePatch(value: unknown): value is Partial<CellStyle> {
  if (!isRecord(value)) return false;
  if (value.horizontalAlignment !== undefined && !isHorizontalAlignment(value.horizontalAlignment)) return false;
  if (value.verticalAlignment !== undefined && !isVerticalAlignment(value.verticalAlignment)) return false;
  if (value.shrinkToFit !== undefined && typeof value.shrinkToFit !== 'boolean') return false;
  for (const key of ['bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript', 'wrapText', 'locked', 'formulaHidden'] as const) if (value[key] !== undefined && typeof value[key] !== 'boolean') return false;
  if (value.underlineStyle !== undefined && !['single', 'double', 'singleAccounting', 'doubleAccounting'].includes(String(value.underlineStyle))) return false;
  if (value.textDirection !== undefined && !['context', 'ltr', 'rtl'].includes(String(value.textDirection))) return false;
  if (value.readingOrder !== undefined && !isReadingOrder(value.readingOrder)) return false;
  if (value.textOrientation !== undefined && !['horizontal', 'stacked', 'rotateUp', 'rotateDown'].includes(String(value.textOrientation))) return false;
  if (value.indent !== undefined && (!Number.isInteger(value.indent) || Number(value.indent) < 0 || Number(value.indent) > 250)) return false;
  if (value.textRotate !== undefined && (typeof value.textRotate !== 'number' || !Number.isFinite(value.textRotate) || value.textRotate < 0 || value.textRotate > 180)) return false;
  if (value.fill !== undefined && !isCellFill(value.fill)) return false;
  if (value.numberFormatSpec !== undefined && !isCellNumberFormatSpec(value.numberFormatSpec)) return false;
  return value.unsupportedAlignment === undefined;
}

function isCellFill(value: unknown): value is NonNullable<CellStyle['fill']> {
  if (!isRecord(value) || !['solid', 'pattern', 'gradient'].includes(String(value.kind))) return false;
  if (value.foreground !== undefined && typeof value.foreground !== 'string') return false;
  if (value.background !== undefined && typeof value.background !== 'string') return false;
  if (value.kind === 'pattern' && !['solid', 'none', 'gray125', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray0625', 'lightGray', 'darkGray', 'mediumGray'].includes(String(value.pattern))) return false;
  if (value.kind === 'gradient') {
    if (value.gradientType !== undefined && value.gradientType !== 'linear' && value.gradientType !== 'path') return false;
    if (value.degree !== undefined && (typeof value.degree !== 'number' || !Number.isFinite(value.degree))) return false;
    if (!Array.isArray(value.stops) || value.stops.length < 2 || !value.stops.every((stop) => isRecord(stop) && typeof stop.color === 'string' && typeof stop.position === 'number' && Number.isFinite(stop.position) && stop.position >= 0 && stop.position <= 1)) return false;
  }
  return true;
}

function isCellNumberFormatSpec(value: unknown): value is NonNullable<CellStyle['numberFormatSpec']> {
  if (!isRecord(value)) return false;
  const categories = ['general', 'number', 'currency', 'accounting', 'date', 'time', 'percentage', 'fraction', 'scientific', 'text', 'special', 'custom'];
  if (value.category !== undefined && !categories.includes(String(value.category))) return false;
  if (value.locale !== undefined && typeof value.locale !== 'string') return false;
  if (value.decimalPlaces !== undefined && (!Number.isSafeInteger(value.decimalPlaces) || Number(value.decimalPlaces) < 0 || Number(value.decimalPlaces) > 30)) return false;
  if (value.useThousandsSeparator !== undefined && typeof value.useThousandsSeparator !== 'boolean') return false;
  if (value.negativeStyle !== undefined && !['minus', 'parentheses', 'red-minus', 'red-parentheses'].includes(String(value.negativeStyle))) return false;
  if (value.currencySymbol !== undefined && typeof value.currencySymbol !== 'string') return false;
  if (value.sample !== undefined && typeof value.sample !== 'string') return false;
  return true;
}

function normalizeStyleFontFamily(style: Partial<CellStyle> | undefined): Partial<CellStyle> | undefined {
  if (!style || style.fontFamily === undefined) return style;
  return { ...style, fontFamily: normalizeFontFamily(style.fontFamily) };
}

function isBorderCommand(value: unknown): value is SetBorderParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || !isBorderPlacement(value.placement)) return false;
  if (value.range !== undefined && value.ranges !== undefined) return false;
  const ranges = value.range !== undefined
    ? (isRange(value.range) ? [value.range] : undefined)
    : (Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange) ? value.ranges : undefined);
  if (!ranges || ranges.some((range) => range.sheetId !== value.sheetId)) return false;
  if (value.placement === 'none') return value.line === undefined;
  return isBorderLine(value.line);
}

function styleAffectedRanges(value: SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string }): RangeRef[] {
  return 'range' in value ? [structuredClone(value.range)] : value.ranges.map((range) => structuredClone(range));
}

function isSheetAtCountMutation(value: unknown): value is { sheetId: string; at: number; count: number } {
  return isRecord(value) && typeof value.sheetId === 'string'
    && Number.isInteger(value.at) && Number(value.at) >= 0
    && Number.isInteger(value.count) && Number(value.count) > 0;
}

function structuralAffectedRanges(value: { sheetId: string; at: number; count: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: value.at, endRow: value.at + value.count - 1, startColumn: 0, endColumn: 0 }];
}

function columnStructuralAffectedRanges(value: { sheetId: string; at: number; count: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: value.at, endColumn: value.at + value.count - 1 }];
}

function isSheetIndexMutation(value: unknown): value is { sheetId: string; index: number } {
  return isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.index) && Number(value.index) >= 0;
}

function isSheetIndicesMutation(value: unknown): value is { sheetId: string; indices: number[] } {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.indices)
    && value.indices.every((index) => Number.isInteger(index) && Number(index) >= 0);
}

function isSelectedDimensionCommand(value: unknown): value is { sheetId: string; indices: number[] } {
  return isSheetIndicesMutation(value) && value.indices.length > 0 && new Set(value.indices).size === value.indices.length;
}

function executeSelectedDimensionCommand(
  runtime: CommandRuntime,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  params: { sheetId: string; indices: number[] },
  context: CommandContext,
): CommandResult {
  if (!isSelectedDimensionCommand(params)) throw new Error(`Invalid selected ${axis} ${operation} command payload`);
  const sheet = context.workbook.getSheet(params.sheetId);
  const limit = axis === 'row' ? sheet.rowCount : sheet.columnCount;
  if (params.indices.some((index) => index >= limit)) throw new Error(`Selected ${axis} index is outside the worksheet bounds`);
  const commandId = axis === 'row'
    ? operation === 'insert' ? 'sheet.rows.insert' : 'sheet.rows.delete'
    : operation === 'insert' ? 'sheet.columns.insert' : 'sheet.columns.delete';
  const ordered = [...params.indices].sort((left, right) => right - left);
  const affectedRanges: RangeRef[] = [];
  let mutationCount = 0;
  for (const index of ordered) {
    const result = runtime.execute(commandId, { sheetId: params.sheetId, at: index, count: 1 });
    mutationCount += result.mutationCount;
    affectedRanges.push(...result.affectedRanges);
  }
  return { operationId: context.operationId, mutationCount, affectedRanges };
}

function rowAffectedRange(value: { sheetId: string; row: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: value.row, endRow: value.row, startColumn: 0, endColumn: 0 }];
}

function rowsVisibilityAffectedRanges(value: RowsVisibilityParams): RangeRef[] {
  return value.states.map((state) => rowAffectedRange({ sheetId: value.sheetId, row: state.row })[0]!);
}

function columnAffectedRange(value: { sheetId: string; column: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: value.column, endColumn: value.column }];
}

function isSheetViewMutation(value: unknown): value is { sheetId: string; showGridlines?: boolean; showHeaders?: boolean; zoom?: number } {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.showGridlines === undefined || typeof value.showGridlines === 'boolean')
    && (value.showHeaders === undefined || typeof value.showHeaders === 'boolean')
    && (value.zoom === undefined || (typeof value.zoom === 'number' && Number.isFinite(value.zoom) && value.zoom > 0));
}

function isTableSheetColumn(value: unknown): value is TableSheetDefinition['columns'][number] {
  if (!isRecord(value) || typeof value.fieldId !== 'string' || value.fieldId.trim().length === 0 || typeof value.caption !== 'string' || value.caption.trim().length === 0) return false;
  if (value.widthPx !== undefined && (typeof value.widthPx !== 'number' || !Number.isFinite(value.widthPx) || value.widthPx <= 0)) return false;
  if (value.type !== undefined && typeof value.type !== 'string') return false;
  return value.formula === undefined || typeof value.formula === 'string';
}

function isTableSheetDefinition(value: unknown): value is TableSheetDefinition {
  if (!isRecord(value) || typeof value.viewId !== 'string' || value.viewId.trim().length === 0 || !Array.isArray(value.columns) || !Array.isArray(value.grouping)) return false;
  if (!value.columns.every(isTableSheetColumn)) return false;
  const columnIds = new Set(value.columns.map((column) => column.fieldId));
  if (columnIds.size !== value.columns.length) return false;
  if (!value.grouping.every((group) => isRecord(group) && typeof group.fieldId === 'string' && columnIds.has(group.fieldId) && (group.collapsed === undefined || typeof group.collapsed === 'boolean'))) return false;
  if (new Set(value.grouping.map((group) => group.fieldId)).size !== value.grouping.length) return false;
  if (value.sortState !== undefined) {
    if (!Array.isArray(value.sortState) || !value.sortState.every((sort) => isRecord(sort) && typeof sort.fieldId === 'string' && columnIds.has(sort.fieldId) && (sort.direction === 'asc' || sort.direction === 'desc'))) return false;
    if (new Set(value.sortState.map((sort) => sort.fieldId)).size !== value.sortState.length) return false;
  }
  return true;
}

function normalizeTableSheetDefinition(workbook: WorkbookModel, params: TableSheetUpdateParams): TableSheetDefinition {
  if (!isTableSheetDefinition(params.definition)) throw new Error('TableSheet definition is invalid');
  const sheet = workbook.getSheet(params.sheetId);
  if (sheet.kind !== 'table-sheet' || !sheet.tableSheet) throw new Error('TableSheet definition can only be updated on a table-sheet');
  const table = workbook.dataModel.tables.get(params.definition.viewId);
  if (!table) throw new Error(`TableSheet binding table is unavailable: ${params.definition.viewId}`);
  const fieldIds = new Set(table.fields.map((field) => field.id));
  if (params.definition.columns.length === 0 || params.definition.columns.some((column) => !fieldIds.has(column.fieldId))) throw new Error('TableSheet columns must reference fields from the binding table');
  if (params.definition.grouping.some((group) => !fieldIds.has(group.fieldId)) || params.definition.sortState?.some((sort) => !fieldIds.has(sort.fieldId))) throw new Error('TableSheet grouping and sorting must reference binding-table fields');
  return structuredClone(params.definition);
}

function isGanttSheetDefinition(value: unknown): value is GanttSheetDefinition {
  if (!isRecord(value) || typeof value.viewId !== 'string' || !isRecord(value.fieldMap)
    || !isRecord(value.calendar) || !isRecord(value.timeline) || !isRecord(value.dependencyStyle)) return false;
  const fieldMap = value.fieldMap;
  if (['id', 'title', 'start', 'end', 'progress'].some((key) => typeof fieldMap[key] !== 'string' || String(fieldMap[key]).trim().length === 0)) return false;
  if (fieldMap.parentId !== undefined && typeof fieldMap.parentId !== 'string') return false;
  if (fieldMap.dependencies !== undefined && typeof fieldMap.dependencies !== 'string') return false;
  const calendar = value.calendar;
  if (!Array.isArray(calendar.workingDays) || !calendar.workingDays.every((day) => Number.isInteger(day) && day >= 0 && day <= 6)) return false;
  if (typeof calendar.dayStartHour !== 'number' || !Number.isFinite(calendar.dayStartHour)
    || typeof calendar.dayEndHour !== 'number' || !Number.isFinite(calendar.dayEndHour)
    || calendar.dayStartHour < 0 || calendar.dayEndHour > 24 || calendar.dayStartHour >= calendar.dayEndHour) return false;
  const timeline = value.timeline;
  if (!['day', 'week', 'month', 'quarter'].includes(String(timeline.unit))) return false;
  if (timeline.start !== undefined && typeof timeline.start !== 'string') return false;
  if (timeline.end !== undefined && typeof timeline.end !== 'string') return false;
  const style = value.dependencyStyle;
  return typeof style.color === 'string' && style.color.trim().length > 0 && typeof style.width === 'number' && Number.isFinite(style.width) && style.width > 0;
}

function normalizeGanttSheetDefinition(workbook: WorkbookModel, params: GanttSheetUpdateParams): GanttSheetDefinition {
  if (!isGanttSheetDefinition(params.definition)) throw new Error('GanttSheet definition is invalid');
  const sheet = workbook.getSheet(params.sheetId);
  if (sheet.kind !== 'gantt-sheet' || !sheet.ganttSheet) throw new Error('GanttSheet definition can only be updated on a gantt-sheet');
  const table = workbook.dataModel.tables.get(params.definition.viewId);
  if (!table) throw new Error(`GanttSheet binding table is unavailable: ${params.definition.viewId}`);
  const fieldIds = new Set(table.fields.map((field) => field.id));
  const mapping = params.definition.fieldMap;
  for (const key of ['id', 'title', 'start', 'end', 'progress'] as const) {
    if (!fieldIds.has(mapping[key])) throw new Error(`GanttSheet field mapping ${key} is unavailable`);
  }
  for (const key of ['parentId', 'dependencies'] as const) {
    if (mapping[key] !== undefined && !fieldIds.has(mapping[key]!)) throw new Error(`GanttSheet field mapping ${key} is unavailable`);
  }
  return structuredClone(params.definition);
}

function isReportSheetDefinition(value: unknown): value is ReportSheetDefinition {
  if (!isRecord(value) || typeof value.templateSheetId !== 'string' || (value.tableId !== undefined && typeof value.tableId !== 'string')
    || !Array.isArray(value.bindings) || !isRecord(value.pagination) || !isRecord(value.layout) || !Array.isArray(value.dataEntry)) return false;
  if (!['design', 'preview', 'paginated'].includes(String(value.renderMode))) return false;
  if (value.pagination.enabled !== undefined && typeof value.pagination.enabled !== 'boolean') return false;
  if (value.pagination.rowsPerPage !== undefined && (!Number.isInteger(value.pagination.rowsPerPage) || Number(value.pagination.rowsPerPage) <= 0)) return false;
  if (value.pagination.repeatHeaderRows !== undefined && (!Array.isArray(value.pagination.repeatHeaderRows) || !value.pagination.repeatHeaderRows.every((row) => Number.isInteger(row) && row >= 0))) return false;
  const layout = value.layout;
  if (!['portrait', 'landscape'].includes(String(layout.orientation))) return false;
  for (const key of ['marginTopPx', 'marginRightPx', 'marginBottomPx', 'marginLeftPx']) if (typeof layout[key] !== 'number' || !Number.isFinite(layout[key]) || layout[key] < 0) return false;
  for (const raw of value.bindings) {
    if (!isRecord(raw) || !isRecord(raw.cell) || !Number.isInteger(raw.cell.row) || !Number.isInteger(raw.cell.column) || Number(raw.cell.row) < 0 || Number(raw.cell.column) < 0 || typeof raw.expression !== 'string' || !raw.expression.trim() || !['static', 'field', 'formula', 'group', 'summary'].includes(String(raw.kind))) return false;
    if (raw.direction !== undefined && !['vertical', 'horizontal'].includes(String(raw.direction))) return false;
    if (raw.fill !== undefined && !['none', 'down', 'right'].includes(String(raw.fill))) return false;
    if (raw.summary !== undefined && !['sum', 'count', 'average', 'min', 'max'].includes(String(raw.summary))) return false;
  }
  return value.dataEntry.every((entry) => isRecord(entry) && typeof entry.fieldId === 'string' && typeof entry.writable === 'boolean' && (entry.required === undefined || typeof entry.required === 'boolean'));
}

function normalizeReportSheetDefinition(workbook: WorkbookModel, params: ReportSheetUpdateParams): ReportSheetDefinition {
  if (!isReportSheetDefinition(params.definition)) throw new Error('ReportSheet definition is invalid');
  const sheet = workbook.getSheet(params.sheetId);
  if (sheet.kind !== 'report-sheet' || !sheet.reportSheet) throw new Error('ReportSheet definition can only be updated on a report-sheet');
  if (!workbook.sheets.has(params.definition.templateSheetId)) throw new Error(`ReportSheet template sheet is unavailable: ${params.definition.templateSheetId}`);
  const table = params.definition.tableId ? workbook.dataModel.tables.get(params.definition.tableId) : undefined;
  if (params.definition.tableId && !table) throw new Error(`ReportSheet binding table is unavailable: ${params.definition.tableId}`);
  const fieldIds = new Set(table?.fields.map((field) => field.id) ?? []);
  for (const binding of params.definition.bindings) if (['field', 'group', 'summary'].includes(binding.kind) && !fieldIds.has(binding.expression)) throw new Error(`ReportSheet binding field is unavailable: ${binding.expression}`);
  if (params.definition.dataEntry.some((entry) => !fieldIds.has(entry.fieldId))) throw new Error('ReportSheet data-entry fields must belong to the binding table');
  return structuredClone(params.definition);
}

function tableSheetAffectedRange(workbook: WorkbookModel, params: { sheetId: string }): RangeRef[] {
  const sheet = workbook.getSheet(params.sheetId);
  return [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
}

function isFilterMutation(value: unknown): value is { sheetId: string; autoFilter: AutoFilterModel } {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.autoFilter)
    && typeof value.autoFilter.sheetId === 'string' && isRange(value.autoFilter.range)
    && isRecord(value.autoFilter.columns);
}

function isFilterRemoveMutation(value: unknown): value is { sheetId: string; range?: RangeRef } {
  return isRecord(value) && typeof value.sheetId === 'string' && (value.range === undefined || isRange(value.range));
}

function isConditionalAddMutation(value: unknown): value is AddConditionalFormatParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.rule)
    && typeof value.rule.id === 'string' && typeof value.rule.sheetId === 'string'
    && Array.isArray(value.rule.ranges) && value.rule.ranges.length > 0 && value.rule.ranges.every(isRange);
}

function isConditionalFormatUpdateParams(value: unknown): value is UpdateConditionalFormatParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.ruleId === 'string' && value.ruleId.length > 0 && isRecord(value.patch);
}

function isConditionalFormatUpdateMutation(value: unknown): value is ConditionalFormatUpdateMutationParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && isConditionalAddMutation({ sheetId: value.sheetId, rule: value.before })
    && isConditionalAddMutation({ sheetId: value.sheetId, rule: value.after })
    && Array.isArray(value.ranges) && value.ranges.every(isRange);
}

function isRuleRemoveMutation(value: unknown): value is { sheetId: string; ruleId: string } {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.ruleId === 'string'
    && (value.ranges === undefined || (Array.isArray(value.ranges) && value.ranges.every(isRange)));
}

function isSheetRangesMutation(value: unknown): value is { sheetId: string; ranges: RangeRef[] } {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.ranges) && value.ranges.every(isRange);
}

function isDataValidationAddMutation(value: unknown): value is AddDataValidationParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.rule)
    && typeof value.rule.id === 'string' && typeof value.rule.sheetId === 'string'
    && Array.isArray(value.rule.ranges) && value.rule.ranges.length > 0 && value.rule.ranges.every(isRange);
}

function isBandedMutation(value: unknown): value is { sheetId: string; rule: BandedRule | null } {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.rule === null || (isRecord(value.rule)
      && isRange(value.rule.range)
      && typeof value.rule.firstColor === 'string' && typeof value.rule.secondColor === 'string'));
}

function isNameSetMutation(value: unknown): value is { model: DefinedNameModel } {
  return isRecord(value) && isRecord(value.model)
    && typeof value.model.name === 'string' && value.model.name.trim().length > 0
    && typeof value.model.formula === 'string' && value.model.formula.trim().length > 0
    && (value.model.scope === 'workbook' || value.model.scope === 'sheet')
    && (value.model.scope === 'workbook' || typeof value.model.sheetId === 'string');
}

function isNameRemoveMutation(value: unknown): value is { name: string } {
  return isRecord(value) && typeof value.name === 'string' && value.name.trim().length > 0;
}

function ruleRanges(value: { rule: { ranges: RangeRef[] } }): RangeRef[] {
  return value.rule.ranges.map((range) => structuredClone(range));
}

function removeRuleRanges(value: { ranges?: RangeRef[] }): RangeRef[] {
  return value.ranges?.map((range) => structuredClone(range)) ?? [];
}

function restoreCell(
  workbook: WorkbookModel,
  item: MutationInfo<{ row: number; column: number; previous?: CellData }>,
): void {
  const sheet = workbook.getSheet(item.sheetId);
  const { row, column, previous } = item.params;
  if (previous) sheet.cells.set(row, column, previous);
  else sheet.cells.delete(row, column);
}

function assertCanonicalCheckboxCell(cell: CellData | undefined): void {
  if (cell?.editor?.kind !== 'checkbox') return;
  const normalized = normalizeCheckboxCellValue(cell, cell.editor);
  if (!Object.is(normalized, cell.value)) throw new Error('Checkbox cell value must match one configured canonical state');
}

export function registerSheetCommands(runtime: CommandRuntime): void {
  registerEditingCommands(runtime);
  registerDataToolCommands(runtime);
  registerSheetTableCommands(runtime);
  registerOutlineCommands(runtime);
  registerHomeCommands(runtime);
  registerPhoneticCommands(runtime);
  registerCellTemplateCommands(runtime);

  runtime.registry.registerMutation<SheetExtentParams>({
    id: 'sheet.extent.grow',
    handler: (item, context) => {
      if (!isSheetExtentMutation(item.params)) throw new Error('Invalid sheet.extent.grow mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      if (item.params.rowCount < sheet.rowCount || item.params.columnCount < sheet.columnCount) {
        throw new Error('Sheet extent growth cannot shrink a worksheet');
      }
      sheet.rowCount = item.params.rowCount;
      sheet.columnCount = item.params.columnCount;
    },
    metadata: {
      schema: { name: 'SheetExtentGrow', validate: isSheetExtentMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.extent.restore'],
    },
  });
  runtime.registry.registerMutation<SheetExtentParams>({
    id: 'sheet.extent.restore',
    handler: (item, context) => {
      if (!isSheetExtentMutation(item.params)) throw new Error('Invalid sheet.extent.restore mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      sheet.rowCount = item.params.rowCount;
      sheet.columnCount = item.params.columnCount;
    },
    metadata: {
      schema: { name: 'SheetExtentRestore', validate: isSheetExtentMutation },
      permission: { capability: 'navigate', roles: ['owner', 'editor', 'commenter', 'viewer'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.extent.grow'],
    },
  });
  runtime.registry.registerCommand<SheetExtentParams>({
    id: 'sheet.extent.grow',
    history: 'none',
    execute: (params, context) => {
      if (!isSheetExtentMutation(params)) throw new Error('Invalid sheet.extent.grow command payload');
      const sheet = context.workbook.getSheet(params.sheetId);
      const next = {
        sheetId: params.sheetId,
        rowCount: Math.max(sheet.rowCount, params.rowCount),
        columnCount: Math.max(sheet.columnCount, params.columnCount),
      };
      if (next.rowCount === sheet.rowCount && next.columnCount === sheet.columnCount) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const previous = { sheetId: params.sheetId, rowCount: sheet.rowCount, columnCount: sheet.columnCount };
      context.applyMutation({
        id: 'sheet.extent.grow',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: next,
        affectedRanges: [],
        inverse: [{
          id: 'sheet.extent.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: previous,
          affectedRanges: [],
        }],
        apply: () => {
          sheet.rowCount = next.rowCount;
          sheet.columnCount = next.columnCount;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  runtime.registry.registerMutation<RenameWorkbookParams>({
    id: 'workbook.renamed',
    handler: (item, context) => {
      if (!isRenameWorkbookMutation(item.params)) throw new Error('Invalid workbook.renamed mutation payload');
      context.workbook.name = item.params.name;
    },
    metadata: {
      schema: { name: 'RenameWorkbook', validate: isRenameWorkbookMutation },
      permission: { capability: 'workbook.rename', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['workbook.renamed'],
    },
  });
  runtime.registry.registerCommand<RenameWorkbookParams>({
    id: 'workbook.rename',
    execute: (params, context) => {
      const name = params.name.trim();
      if (!name) throw new Error('Workbook name is required');
      const previous = context.workbook.name;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'workbook.renamed',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params: { name },
        affectedRanges,
        inverse: [{ id: 'workbook.renamed', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { name: previous }, affectedRanges }],
        apply: () => { context.workbook.name = name; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 1. Sheet mutations & commands
  runtime.registry.registerMutation<AddSheetParams>({
    id: 'sheet.add',
    handler: (item, context) => {
      if (!isAddSheetMutation(item.params)) throw new Error('Invalid sheet.add mutation payload');
      const params = item.params;
      context.workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount);
    },
    metadata: {
      schema: { name: 'AddSheet', validate: isAddSheetMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.remove'],
    },
  });
  runtime.registry.registerMutation<{ id: string }>({
    id: 'sheet.remove',
    handler: (item, context) => {
      if (!isSheetIdMutation(item.params)) throw new Error('Invalid sheet.remove mutation payload');
      context.workbook.removeSheet(item.params.id);
    },
    metadata: {
      schema: { name: 'RemoveSheet', validate: isSheetIdMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.restore'],
    },
  });
  runtime.registry.registerMutation<RenameSheetParams>({
    id: 'sheet.rename',
    handler: (item, context) => {
      if (!isRenameSheetMutation(item.params)) throw new Error('Invalid sheet.rename mutation payload');
      const params = item.params;
      context.workbook.renameSheet(params.sheetId, params.name);
    },
    metadata: {
      schema: { name: 'RenameSheet', validate: isRenameSheetMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.rename'],
    },
  });
  runtime.registry.registerMutation<{ sheet: import('@react-sheets/core-model').SheetSnapshot; index?: number }>({
    id: 'sheet.restore',
    handler: (item, context) => {
      if (!isSheetRestoreMutation(item.params)) throw new Error('Invalid sheet.restore mutation payload');
      context.workbook.restoreSheetSnapshot(item.params.sheet, item.params.index);
    },
    metadata: {
      schema: { name: 'RestoreSheet', validate: isSheetRestoreMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.remove'],
    },
  });
  runtime.registry.registerMutation<TableSheetUpdateParams>({
    id: 'tableSheet.update',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isTableSheetDefinition(item.params.definition)) throw new Error('Invalid tableSheet.update mutation payload');
      const params = item.params as TableSheetUpdateParams;
      const definition = normalizeTableSheetDefinition(context.workbook, params);
      context.workbook.getSheet(params.sheetId).tableSheet = definition;
    },
    metadata: {
      schema: { name: 'TableSheetDefinitionUpdate', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isTableSheetDefinition(value.definition) },
      permission: { capability: 'table-sheet.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => sheetScopeRange(params), mode: 'declared' },
      inverseIds: ['tableSheet.update'],
    },
  });
  runtime.registry.registerMutation<GanttSheetUpdateParams>({
    id: 'ganttSheet.update',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isGanttSheetDefinition(item.params.definition)) throw new Error('Invalid ganttSheet.update mutation payload');
      const params = item.params as GanttSheetUpdateParams;
      context.workbook.getSheet(params.sheetId).ganttSheet = normalizeGanttSheetDefinition(context.workbook, params);
    },
    metadata: {
      schema: { name: 'GanttSheetDefinitionUpdate', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isGanttSheetDefinition(value.definition) },
      permission: { capability: 'gantt-sheet.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => sheetScopeRange(params), mode: 'declared' },
      inverseIds: ['ganttSheet.update'],
    },
  });
  runtime.registry.registerMutation<ReportSheetUpdateParams>({
    id: 'reportSheet.update',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isReportSheetDefinition(item.params.definition)) throw new Error('Invalid reportSheet.update mutation payload');
      const params = item.params as ReportSheetUpdateParams;
      context.workbook.getSheet(params.sheetId).reportSheet = normalizeReportSheetDefinition(context.workbook, params);
    },
    metadata: {
      schema: { name: 'ReportSheetDefinitionUpdate', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isReportSheetDefinition(value.definition) },
      permission: { capability: 'report-sheet.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => sheetScopeRange(params), mode: 'declared' },
      inverseIds: ['reportSheet.update'],
    },
  });

  runtime.registry.registerCommand<{ id: string }>({
    id: 'sheet.remove',
    execute: (paramsInput, context) => {
      const params = paramsInput as { id: string };
      const workbook = context.workbook;
      if (workbook.getSheets().length <= 1) {
        throw new Error('A workbook must keep at least one worksheet');
      }
      const index = workbook.sheetOrder.indexOf(params.id);
      const snapshot = workbook.getSheetSnapshot(params.id);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.remove',
        unitId: workbook.unitId,
        sheetId: params.id,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.restore',
            unitId: workbook.unitId,
            sheetId: params.id,
            params: { sheet: snapshot, index },
            affectedRanges,
          },
        ],
        apply: () => workbook.removeSheet(params.id),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<AddSheetParams>({
    id: 'sheet.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.add',
        unitId: context.workbook.unitId,
        sheetId: params.id,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.remove',
            unitId: context.workbook.unitId,
            sheetId: params.id,
            params: { id: params.id },
            affectedRanges,
          },
        ],
        apply: () =>
          context.workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<TableSheetUpdateParams>({
    id: 'tableSheet.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.tableSheet ? structuredClone(sheet.tableSheet) : undefined;
      if (sheet.kind !== 'table-sheet' || !previous) throw new Error('TableSheet definition can only be updated on a table-sheet');
      const definition = normalizeTableSheetDefinition(context.workbook, params);
      if (JSON.stringify(previous) === JSON.stringify(definition)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = tableSheetAffectedRange(context.workbook, params);
      context.applyMutation({
        id: 'tableSheet.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, definition },
        affectedRanges,
        inverse: [{
          id: 'tableSheet.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, definition: previous },
          affectedRanges,
        }],
        apply: () => { sheet.tableSheet = structuredClone(definition); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<GanttSheetUpdateParams>({
    id: 'ganttSheet.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.ganttSheet ? structuredClone(sheet.ganttSheet) : undefined;
      if (sheet.kind !== 'gantt-sheet' || !previous) throw new Error('GanttSheet definition can only be updated on a gantt-sheet');
      const definition = normalizeGanttSheetDefinition(context.workbook, params);
      if (JSON.stringify(previous) === JSON.stringify(definition)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = tableSheetAffectedRange(context.workbook, params);
      context.applyMutation({
        id: 'ganttSheet.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, definition },
        affectedRanges,
        inverse: [{
          id: 'ganttSheet.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, definition: previous },
          affectedRanges,
        }],
        apply: () => { sheet.ganttSheet = structuredClone(definition); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<ReportSheetUpdateParams>({
    id: 'reportSheet.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.reportSheet ? structuredClone(sheet.reportSheet) : undefined;
      if (sheet.kind !== 'report-sheet' || !previous) throw new Error('ReportSheet definition can only be updated on a report-sheet');
      const definition = normalizeReportSheetDefinition(context.workbook, params);
      if (JSON.stringify(previous) === JSON.stringify(definition)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = tableSheetAffectedRange(context.workbook, params);
      context.applyMutation({
        id: 'reportSheet.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, definition },
        affectedRanges,
        inverse: [{ id: 'reportSheet.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, definition: previous }, affectedRanges }],
        apply: () => { sheet.reportSheet = structuredClone(definition); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<RenameSheetParams>({
    id: 'sheet.rename',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousName = sheet.name;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.rename',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.rename',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, name: previousName },
            affectedRanges,
          },
        ],
        apply: () => context.workbook.renameSheet(params.sheetId, params.name),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<WorkbookTableModel>({
    id: 'table.add',
    handler: (item, context) => {
      if (!isWorkbookTableMutation(item.params)) throw new Error('Invalid table.add mutation payload');
      context.workbook.addTable(item.params);
    },
    metadata: {
      schema: { name: 'WorkbookTableModel', validate: isWorkbookTableMutation },
      permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: workbookTableRanges, mode: 'exact' },
      inverseIds: ['table.remove'],
    },
  });
  runtime.registry.registerMutation<{ tableId: string; range?: RangeRef }>({
    id: 'table.remove',
    handler: (item, context) => {
      if (!isTableRemoveMutation(item.params)) throw new Error('Invalid table.remove mutation payload');
      context.workbook.removeTable(item.params.tableId);
    },
    metadata: {
      schema: { name: 'TableRemove', validate: isTableRemoveMutation },
      permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: tableRemoveRanges, mode: 'declared' },
      inverseIds: ['table.add'],
    },
  });
  runtime.registry.registerCommand<AddTableParams>({
    id: 'table.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'table.add',
        unitId: context.workbook.unitId,
        sheetId: params.sourceSheetId ?? context.workbook.primarySheetId,
        params: structuredClone(params),
        affectedRanges,
        inverse: [{ id: 'table.remove', unitId: context.workbook.unitId, sheetId: params.sourceSheetId ?? context.workbook.primarySheetId, params: { tableId: params.id, range: params.sourceRange }, affectedRanges }],
        apply: () => context.workbook.addTable(params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ tableId: string; sheetId: string }>({
    id: 'table.remove',
    execute: (params, context) => {
      const previous = structuredClone(context.workbook.getTable(params.tableId));
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'table.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { tableId: params.tableId, range: previous.sourceRange },
        affectedRanges,
        inverse: [{ id: 'table.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => context.workbook.removeTable(params.tableId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 2. Cell mutations & commands
  runtime.registry.registerMutation<CellSetMutationParams>({
    id: 'cell.set',
    handler: (item, context) => {
      if (!isCellSetMutationParams(item.params)) throw new Error('Invalid cell.set mutation payload');
      const params = item.params;
      assertCellWriteAuthority(params, context.workbook.getSheet(params.sheetId));
      const value = clearFormulaProvenance(params.value);
      assertCanonicalCheckboxCell(value);
      context.workbook.getSheet(params.sheetId).cells.set(params.row, params.column, value);
    },
    metadata: {
      schema: { name: 'SetCellValue', validate: isCellSetMutationParams },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: cellRange, mode: 'exact' },
      inverseIds: ['cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; row: number; column: number; previous?: CellData }>({
    id: 'cell.restore',
    handler: (item, context) => {
      if (!isCellRestoreMutation(item.params)) throw new Error('Invalid cell.restore mutation payload');
      assertCanonicalCheckboxCell(item.params.previous);
      restoreCell(context.workbook, item as MutationInfo<{ row: number; column: number; previous?: CellData }>);
    },
    metadata: {
      schema: { name: 'RestoreCell', validate: isCellRestoreMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: cellRange, mode: 'declared' },
      inverseIds: ['cell.set'],
    },
  });

  runtime.registry.registerCommand<SetCellValueParams>({
    id: 'sheet.cell.set',
    execute: (params, context) => {
      if (!isCellSetMutation(params)) throw new Error('Invalid cell.set command payload');
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.cells.get(params.row, params.column);
      const affectedRanges = cellRange(params);
      const value = clearFormulaProvenance(params.value);
      const canonicalParams = createCellSetMutationParams(sheet, { ...params, value }, 'script');
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: canonicalParams,
        affectedRanges,
        inverse: [
          {
            id: 'cell.restore',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: params.row, column: params.column, previous },
            affectedRanges,
          },
        ],
        apply: () => sheet.cells.set(params.row, params.column, value),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<CommitTextParams>({
    id: 'sheet.cell.commitText',
    execute: (params, context) => {
      if (typeof params.text !== 'string') {
        throw new CellEntryError({
          code: 'CELL_ENTRY_INVALID_INPUT',
          message: 'Cell text must be a string',
          sheetId: params.sheetId,
          row: params.row,
          column: params.column,
          recovery: 'Pass the unchanged lexical input string to sheet.cell.commitText.',
        });
      }
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.cells.get(params.row, params.column);
      const next = buildCellFromText(params.text, previous, params.inputContext, params.style);
      if (previous?.editor?.kind === 'checkbox') next.value = normalizeCheckboxCellValue(next, previous.editor);
      return commitCellEntry(params, next, context);
    },
  });

  runtime.registry.registerCommand<CommitTypedValueParams>({
    id: 'sheet.cell.commitTypedValue',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.cells.get(params.row, params.column);
      const next = clearFormulaProvenance(previous ? structuredClone(previous) : { value: null });
      next.value = structuredClone(params.value);
      delete next.formula;
      delete next.formulaValue;
      delete next.displayValue;
      delete next.richText;
      return commitCellEntry(params, next, context);
    },
  });

  runtime.registry.registerCommand<CommitRichTextParams>({
    id: 'sheet.cell.commitRichText',
    execute: (params, context) => {
      if (typeof params.text !== 'string' || !Array.isArray(params.runs) || params.runs.some((run) => typeof run.text !== 'string')) {
        throw new CellEntryError({
          code: 'CELL_ENTRY_RICH_TEXT_INVALID',
          message: 'Rich-text commit requires text and canonical runs',
          sheetId: params.sheetId,
          row: params.row,
          column: params.column,
          recovery: 'Submit a canonical rich-text draft with text-preserving runs.',
        });
      }
      if (params.runs.map((run) => run.text).join('') !== params.text) {
        throw new CellEntryError({
          code: 'CELL_ENTRY_RICH_TEXT_INVALID',
          message: 'Rich-text runs do not reproduce the canonical plain text',
          sheetId: params.sheetId,
          row: params.row,
          column: params.column,
          recovery: 'Normalize the run sequence before committing.',
        });
      }
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.cells.get(params.row, params.column);
      const next = clearFormulaProvenance(previous ? structuredClone(previous) : { value: null });
      next.value = params.text;
      next.richText = structuredClone(params.runs);
      delete next.formula;
      delete next.formulaValue;
      delete next.displayValue;
      return commitCellEntry(params, next, context);
    },
  });

  runtime.registry.registerCommand<CommitTextCellsParams>({
    id: 'sheet.cells.commitText',
    execute: (params, context) => {
      if (typeof params.text !== 'string' || !Array.isArray(params.targets) || params.targets.length === 0) {
        throw new CellEntryError({
          code: 'CELL_ENTRY_INVALID_INPUT',
          message: 'Multi-cell text commit requires a raw text value and at least one target',
          sheetId: params.targets?.[0]?.sheetId ?? context.workbook.primarySheetId,
          row: params.targets?.[0]?.row ?? 0,
          column: params.targets?.[0]?.column ?? 0,
          recovery: 'Resolve the complete canonical selection before opening the transaction.',
        });
      }
      const identities = new Set<string>();
      const prepared = params.targets.map((target) => {
        const identity = `${target.sheetId}:${target.row}:${target.column}`;
        if (identities.has(identity)) {
          throw new CellEntryError({
            code: 'CELL_ENTRY_INVALID_INPUT',
            message: `Multi-cell text commit contains a duplicate target: ${identity}`,
            sheetId: target.sheetId,
            row: target.row,
            column: target.column,
            recovery: 'Normalize overlapping selection ranges into unique canonical targets.',
          });
        }
        identities.add(identity);
        const sheet = context.workbook.getSheet(target.sheetId);
        const previous = sheet.cells.get(target.row, target.column);
        const next = buildCellFromText(params.text, previous, target.inputContext, target.style);
        if (previous?.editor?.kind === 'checkbox') next.value = normalizeCheckboxCellValue(next, previous.editor);
        return prepareCellEntry({ ...target, validationConfirmation: params.validationConfirmation }, next, context);
      });
      for (const entry of prepared) applyPreparedCellEntry(entry, context);
      return {
        operationId: context.operationId,
        mutationCount: prepared.length,
        affectedRanges: prepared.flatMap((entry) => entry.affectedRanges),
      };
    },
  });

  runtime.registry.registerCommand<CommitTypedValueCellsParams>({
    id: 'sheet.cells.commitTypedValue',
    execute: (params, context) => {
      if (!Array.isArray(params.targets) || params.targets.length === 0) rejectCellEntry({ sheetId: context.workbook.primarySheetId, row: 0, column: 0 }, 'Multi-cell typed commit requires at least one target', 'Resolve the complete canonical selection before committing.');
      const identities = new Set<string>();
      const prepared = params.targets.map((target) => {
        const identity = `${target.sheetId}:${target.row}:${target.column}`;
        if (identities.has(identity)) rejectCellEntry(target, `Multi-cell typed commit contains a duplicate target: ${identity}`, 'Normalize overlapping selection ranges into unique canonical targets.');
        identities.add(identity);
        const sheet = context.workbook.getSheet(target.sheetId);
        const previous = sheet.cells.get(target.row, target.column);
        const next = clearFormulaProvenance(previous ? structuredClone(previous) : { value: null });
        next.value = structuredClone(params.value);
        delete next.formula;
        delete next.formulaValue;
        delete next.displayValue;
        delete next.richText;
        return prepareCellEntry({ ...target, validationConfirmation: params.validationConfirmation }, next, context);
      });
      for (const entry of prepared) applyPreparedCellEntry(entry, context);
      return { operationId: context.operationId, mutationCount: prepared.length, affectedRanges: prepared.flatMap((entry) => entry.affectedRanges) };
    },
  });

  runtime.registry.registerCommand<CommitRichTextCellsParams>({
    id: 'sheet.cells.commitRichText',
    execute: (params, context) => {
      if (!Array.isArray(params.targets) || params.targets.length === 0 || params.runs.map((run) => run.text).join('') !== params.text) rejectCellEntry(params.targets?.[0] ?? { sheetId: context.workbook.primarySheetId, row: 0, column: 0 }, 'Multi-cell rich-text commit payload is invalid', 'Submit text-preserving rich-text runs and at least one canonical target.');
      const identities = new Set<string>();
      const prepared = params.targets.map((target) => {
        const identity = `${target.sheetId}:${target.row}:${target.column}`;
        if (identities.has(identity)) rejectCellEntry(target, `Multi-cell rich-text commit contains a duplicate target: ${identity}`, 'Normalize overlapping selection ranges into unique canonical targets.');
        identities.add(identity);
        const sheet = context.workbook.getSheet(target.sheetId);
        const previous = sheet.cells.get(target.row, target.column);
        const next = clearFormulaProvenance(previous ? structuredClone(previous) : { value: null });
        next.value = params.text;
        next.richText = structuredClone(params.runs);
        delete next.formula;
        delete next.formulaValue;
        delete next.displayValue;
        return prepareCellEntry({ ...target, validationConfirmation: params.validationConfirmation }, next, context);
      });
      for (const entry of prepared) applyPreparedCellEntry(entry, context);
      return { operationId: context.operationId, mutationCount: prepared.length, affectedRanges: prepared.flatMap((entry) => entry.affectedRanges) };
    },
  });

  // 3. Range set
  runtime.registry.registerMutation<SetRangeValuesParams>({
    id: 'range.set',
    handler: (item, context) => {
      if (!isSetRangeMutation(item.params)) throw new Error('Invalid range.set mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
        const rowValues = params.values[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
          const value = rowValues[columnOffset];
          if (value) sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, clearFormulaProvenance(value));
        }
      }
    },
    metadata: {
      schema: { name: 'SetRangeValues', validate: isSetRangeMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: setRangeAffectedRanges, mode: 'declared' },
      inverseIds: ['cell.restore'],
    },
  });

  runtime.registry.registerCommand<SetRangeValuesParams>({
    id: 'sheet.range.set',
    execute: (params, context) => {
      if (!isSetRangeMutation(params)) throw new Error('Invalid range.set command payload');
      const sheet = context.workbook.getSheet(params.sheetId);
      const values = params.values.map((row) => row.map((value) => value ? clearFormulaProvenance(value) : value));
      const writeRange = setRangeAffectedRanges(params)[0];
      let tablePlansId = 0;
      const tablePlans = writeRange && params.values.length > 0
        ? planSheetTableAutoExpansion(sheet, writeRange, (prefix) => `${prefix}-${context.operationId}-${tablePlansId++}`)
        : [];
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = writeRange ? [structuredClone(writeRange)] : [];
      for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
        const rowValues = values[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
          const row = params.startRow + rowOffset;
          const column = params.startColumn + columnOffset;
          previous.push({ row, column, value: sheet.cells.get(row, column) });
          affectedRanges.push({
            sheetId: params.sheetId,
            startRow: row,
            endRow: row,
            startColumn: column,
            endColumn: column,
          });
        }
      }
      for (const plan of tablePlans) {
        const tableIndex = sheet.sheetTables.findIndex((table) => table.id === plan.previous.id);
        if (tableIndex < 0) throw new Error(`Sheet Table not found: ${plan.previous.id}`);
        const tableAffectedRanges = [structuredClone(plan.next.range)];
        context.applyMutation({
          id: 'sheetTable.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: plan.next,
          affectedRanges: tableAffectedRanges,
          inverse: [{
            id: 'sheetTable.update',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: plan.previous,
            affectedRanges: [structuredClone(plan.previous.range)],
          }],
          apply: () => { sheet.sheetTables[tableIndex] = structuredClone(plan.next); },
        });
        affectedRanges.push(...tableAffectedRanges);
      }
      context.applyMutation({
        id: 'range.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, startRow: params.startRow, startColumn: params.startColumn, values },
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row: item.row, column: item.column, previous: item.value },
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => {
          for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
            const rowValues = values[rowOffset] ?? [];
            for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
              const value = rowValues[columnOffset];
              if (value)
                sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, {
                  ...value,
                });
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1 + tablePlans.length, affectedRanges };
    },
  });

  runtime.registry.registerMutation<ClearRangeParams>({
    id: 'range.clear',
    handler: (item, context) => {
      if (!isClearRangeMutation(item.params)) throw new Error('Invalid range.clear mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      applyClearRangePlan(sheet, createClearRangePlan(sheet, params));
    },
    metadata: {
      schema: { name: 'ClearRange', validate: isClearRangeMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['range.clear.restore', 'cell.restore'],
    },
  });

  runtime.registry.registerMutation<ClearRangeRestoreParams>({
    id: 'range.clear.restore',
    handler: (item, context) => {
      if (!isClearRangeRestoreMutation(item.params)) throw new Error('Invalid range.clear.restore mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      restoreClearRangeSnapshot(sheet, params.range, params.snapshot);
    },
    metadata: {
      schema: { name: 'ClearRangeRestore', validate: isClearRangeRestoreMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['range.clear'],
    },
  });

  runtime.registry.registerMutation<SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string }>({
    id: 'style.set',
    handler: (item, context) => {
    if (!isStyleMutation(item.params)) throw new Error('Invalid style.set mutation payload');
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    const ranges = 'range' in params ? [params.range] : params.ranges;
    const style = 'numberFormat' in params && !('style' in params)
      ? { numberFormat: params.numberFormat }
      : normalizeStyleFontFamily((params as SetRangeStyleParams).style) ?? {};
    for (const range of ranges) {
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          const current = sheet.cells.get(row, column) ?? { value: null as CellData['value'] };
          const next = { ...current };
          if ('replaceStyle' in params && params.replaceStyle) {
            if (Object.keys(style).length > 0) next.style = structuredClone(style);
            else delete next.style;
          } else if (Object.keys(style).length > 0) {
            next.style = { ...(current.style ?? {}), ...style };
          }
          const numberFormat = 'numberFormat' in params && typeof params.numberFormat === 'string'
            ? params.numberFormat
            : style.numberFormat;
          if (numberFormat !== undefined) next.numberFormat = numberFormat;
          if ('clearNumberFormat' in params && params.clearNumberFormat) delete next.numberFormat;
          if ('replaceStyle' in params && params.replaceStyle) delete next.displayValue;
          sheet.cells.set(row, column, next);
        }
      }
    }
    },
    metadata: {
      schema: { name: 'StyleSet', validate: isStyleMutation },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: styleAffectedRanges, mode: 'declared' },
      inverseIds: ['cell.restore'],
    },
  });

  // 4. Range Clear
  runtime.registry.registerCommand<ClearRangeParams>({
    id: 'sheet.range.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (!isClearRangeMutation(params)) throw new Error('Invalid sheet.range.clear command payload');
      const plan = createClearRangePlan(sheet, params);
      const { range } = plan;
      const affectedRanges: RangeRef[] = [range];

      context.applyMutation({
        id: 'range.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, range },
        affectedRanges,
        inverse: [{
          id: 'range.clear.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, range, snapshot: plan.snapshot },
          affectedRanges,
        }],
        apply: () => runtime.registry.getMutation('range.clear')({ id: 'range.clear', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, range }, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 5. Range Style batch application
  runtime.registry.registerCommand<SetRangeStyleParams>({
    id: 'sheet.style.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const canonicalParams: SetRangeStyleParams = {
        ...params,
        style: normalizeStyleFontFamily(params.style) ?? {},
      };
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [canonicalParams.range];

      for (let r = canonicalParams.range.startRow; r <= canonicalParams.range.endRow; r++) {
        for (let c = canonicalParams.range.startColumn; c <= canonicalParams.range.endColumn; c++) {
          const cell = sheet.cells.get(r, c);
          previous.push({ row: r, column: c, value: cell ? structuredClone(cell) : undefined });
        }
      }

      context.applyMutation({
        id: 'style.set',
        unitId: context.workbook.unitId,
        sheetId: canonicalParams.sheetId,
        params: canonicalParams,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row: item.row, column: item.column, previous: item.value },
          permission: { capability: 'format', protectionAction: 'format', checksProtection: true, affectedRangeMode: 'declared', objectScope: 'range' },
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => runtime.registry.getMutation('style.set')({ id: 'style.set', unitId: context.workbook.unitId, sheetId: canonicalParams.sheetId, params: canonicalParams, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 5b. Canonical border topology. The planner is pure; this command expands
  // its side-aware patches into existing style.set mutations so the current
  // collaboration, permission, undo and OOXML style paths remain identical.
  runtime.registry.registerCommand<SetBorderParams>({
    id: 'sheet.borders.set',
    execute: (params, context) => {
      if (!isBorderCommand(params)) throw new Error('Invalid sheet.borders.set command payload');
      const sheet = context.workbook.getSheet(params.sheetId);
      const inputRanges = params.range ? [params.range] : (params.ranges ?? []);
      const plans = inputRanges.map((range) => planBorderChange(range, params.placement, params.line, {
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
      }));
      const affectedRanges = plans.map((plan) => structuredClone(plan.range));

      for (const plan of plans) {
        for (const plannedCell of plan.cells) {
          // Empty side sets are no-ops for perimeter/inside placements. The
          // explicit none placement is the only operation that clears a cell.
          if (plan.placement !== 'none' && Object.keys(plannedCell.sides).length === 0) continue;
          const previous = sheet.cells.get(plannedCell.row, plannedCell.column);
          const existingBorders = previous?.style?.borders ?? {};
          const borders = plan.placement === 'none'
            ? {}
            : { ...existingBorders, ...plannedCell.sides };
          const cellRange: RangeRef = {
            sheetId: params.sheetId,
            startRow: plannedCell.row,
            endRow: plannedCell.row,
            startColumn: plannedCell.column,
            endColumn: plannedCell.column,
          };
          const style: Partial<CellStyle> = { borders };
          context.applyMutation({
            id: 'style.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: cellRange, style },
            affectedRanges: [cellRange],
            inverse: [{
              id: 'cell.restore',
              unitId: context.workbook.unitId,
              sheetId: params.sheetId,
              params: { sheetId: params.sheetId, row: plannedCell.row, column: plannedCell.column, previous: previous ? structuredClone(previous) : undefined },
              permission: { capability: 'format', protectionAction: 'format', checksProtection: true, affectedRangeMode: 'declared', objectScope: 'range' },
              affectedRanges: [cellRange],
            }],
            apply: () => runtime.registry.getMutation('style.set')({
              id: 'style.set',
              unitId: context.workbook.unitId,
              sheetId: params.sheetId,
              params: { sheetId: params.sheetId, range: cellRange, style },
              affectedRanges: [cellRange],
            }, context),
          });
        }
      }
      return { operationId: context.operationId, mutationCount: 0, affectedRanges };
    },
  });

  // 6. Merge commands & mutations
  runtime.registry.registerMutation<SetMergeParams>({
    id: 'merge.set',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isRange(item.params.range)) throw new Error('Invalid merge.set mutation payload');
      const params = item.params as SetMergeParams;
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.merges.push({ range: params.range, anchor: { row: params.range.startRow, column: params.range.startColumn } });
    },
    metadata: {
      schema: { name: 'SetMerge', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range) },
      permission: { capability: 'sheet.merge.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['merge.remove'],
    },
  });
  runtime.registry.registerMutation<RemoveMergeParams>({
    id: 'merge.remove',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isRange(item.params.range)) throw new Error('Invalid merge.remove mutation payload');
      const params = item.params as RemoveMergeParams;
      const sheet = context.workbook.getSheet(params.sheetId);
      const idx = sheet.merges.findIndex((m) => m.range.startRow === params.range.startRow && m.range.startColumn === params.range.startColumn);
      if (idx >= 0) sheet.merges.splice(idx, 1);
    },
    metadata: {
      schema: { name: 'RemoveMerge', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range) },
      permission: { capability: 'sheet.merge.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['merge.set'],
    },
  });

  runtime.registry.registerCommand<SetMergeParams>({
    id: 'sheet.merge.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const span: MergeSpan = {
        range: params.range,
        anchor: { row: params.range.startRow, column: params.range.startColumn },
      };
      const affectedRanges = [params.range];

      context.applyMutation({
        id: 'merge.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'merge.remove',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: params.range },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.merges.push(span);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<RemoveMergeParams>({
    id: 'sheet.merge.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const idx = sheet.merges.findIndex(
        (m) =>
          m.range.startRow === params.range.startRow &&
          m.range.startColumn === params.range.startColumn,
      );
      const affectedRanges = [params.range];
      if (idx < 0)
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previousSpan = sheet.merges[idx]!;

      context.applyMutation({
        id: 'merge.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'merge.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: previousSpan.range },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.merges.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 7. Freeze commands
  runtime.registry.registerMutation<SetFreezeParams>({
    id: 'freeze.set',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isWorksheetPane(item.params.pane)) throw new Error('Invalid freeze.set mutation payload');
      const params = item.params as SetFreezeParams;
      context.workbook.getSheet(params.sheetId).pane = { ...params.pane };
    },
    metadata: {
      schema: { name: 'SetPane', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isWorksheetPane(value.pane) },
      permission: { capability: 'sheet.view.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['freeze.set'],
    },
  });

  runtime.registry.registerCommand<SetFreezeParams>({
    id: 'sheet.freeze.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousPane = { ...sheet.pane };
      const affectedRanges: RangeRef[] = [];

      context.applyMutation({
        id: 'freeze.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'freeze.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, pane: previousPane },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.pane = { ...params.pane };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 8. Row and Column resizing
  runtime.registry.registerMutation<ResizeRowParams>({
    id: 'row.resize',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !Number.isInteger(item.params.row) || typeof item.params.heightPx !== 'number' || item.params.heightPx <= 0) throw new Error('Invalid row.resize mutation payload');
      const params = item.params as ResizeRowParams;
      context.workbook.getSheet(params.sheetId).rowHeightsPx[params.row] = params.heightPx;
    },
    metadata: {
      schema: { name: 'ResizeRowPx', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.row) && typeof value.heightPx === 'number' && Number(value.row) >= 0 && Number(value.heightPx) > 0 },
      permission: { capability: 'sheet.dimension.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: rowAffectedRange, mode: 'declared' },
      inverseIds: ['row.resize'],
    },
  });
  runtime.registry.registerMutation<ResizeColumnParams>({
    id: 'column.resize',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !Number.isInteger(item.params.column) || typeof item.params.widthPx !== 'number' || item.params.widthPx <= 0) throw new Error('Invalid column.resize mutation payload');
      const params = item.params as ResizeColumnParams;
      context.workbook.getSheet(params.sheetId).columnWidthsPx[params.column] = params.widthPx;
    },
    metadata: {
      schema: { name: 'ResizeColumnPx', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.column) && typeof value.widthPx === 'number' && Number(value.column) >= 0 && Number(value.widthPx) > 0 },
      permission: { capability: 'sheet.dimension.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: columnAffectedRange, mode: 'declared' },
      inverseIds: ['column.resize'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; widthPx: number }>({
    id: 'column.defaultWidth.resize',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || typeof item.params.widthPx !== 'number' || !Number.isFinite(item.params.widthPx) || item.params.widthPx <= 0) throw new Error('Invalid column.defaultWidth.resize mutation payload');
      context.workbook.getSheet(item.params.sheetId).defaultColumnWidthPx = item.params.widthPx;
    },
    metadata: {
      schema: { name: 'ResizeDefaultColumnWidthPx', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && typeof value.widthPx === 'number' && Number.isFinite(value.widthPx) && value.widthPx > 0 },
      permission: { capability: 'sheet.dimension.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['column.defaultWidth.resize'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; rows?: Array<Omit<ResizeRowParams, 'sheetId'>>; columns?: Array<Omit<ResizeColumnParams, 'sheetId'>> }>({
    id: 'sheet.dimensions.apply',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [];
      let mutationCount = 0;
      for (const row of params.rows ?? []) {
        if (!Number.isSafeInteger(row.row) || row.row < 0 || !Number.isFinite(row.heightPx) || row.heightPx <= 0) throw new Error('Invalid row pixel size');
        const mutationParams: ResizeRowParams = { sheetId: params.sheetId, ...row };
        const previousHeightPx = sheet.rowHeightsPx[row.row] ?? sheet.defaultRowHeightPx;
        context.applyMutation({
          id: 'row.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: mutationParams, affectedRanges,
          inverse: [{ id: 'row.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: row.row, heightPx: previousHeightPx }, affectedRanges }],
          apply: () => { sheet.rowHeightsPx[row.row] = row.heightPx; },
        });
        mutationCount += 1;
      }
      for (const column of params.columns ?? []) {
        if (!Number.isSafeInteger(column.column) || column.column < 0 || !Number.isFinite(column.widthPx) || column.widthPx <= 0) throw new Error('Invalid column pixel size');
        const mutationParams: ResizeColumnParams = { sheetId: params.sheetId, ...column };
        const previousWidthPx = sheet.columnWidthsPx[column.column] ?? sheet.defaultColumnWidthPx;
        context.applyMutation({
            id: 'column.resize',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: mutationParams,
            affectedRanges,
            inverse: [{ id: 'column.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, column: column.column, widthPx: previousWidthPx }, affectedRanges }],
            apply: () => { sheet.columnWidthsPx[column.column] = column.widthPx; },
        });
        mutationCount += 1;
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; widthPx: number }>({
    id: 'sheet.column.defaultWidth.set',
    execute: (params, context) => {
      if (!Number.isFinite(params.widthPx) || params.widthPx <= 0) throw new Error('Default column width must be positive pixels');
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousWidthPx = sheet.defaultColumnWidthPx;
      const affectedRanges = sheetScopeRange(params);
      context.applyMutation({
        id: 'column.defaultWidth.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges,
        inverse: [{ id: 'column.defaultWidth.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, widthPx: previousWidthPx }, affectedRanges }],
        apply: () => { sheet.defaultColumnWidthPx = params.widthPx; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 9. Range Sorting
  runtime.registry.registerCommand<SortRangeParams>({
    id: 'sheet.sort',
    execute: (params, context) => runtime.execute('data.sort.rows', {
      sheetId: params.sheetId,
      range: params.range,
      criteria: [{ column: params.sortColumn, ascending: params.ascending }],
      hasHeader: params.hasHeader,
    }),
  });

  // 10. 结构操作:行/列插入与删除（统一走 StructuralTransform）
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'rows.inserted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid rows.inserted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'insert-rows', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'RowsInserted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: structuralAffectedRanges, mode: 'declared' },
      inverseIds: ['rows.deleted'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'rows.deleted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid rows.deleted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'delete-rows', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'RowsDeleted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: structuralAffectedRanges, mode: 'declared' },
      inverseIds: ['rows.inserted', 'cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'columns.inserted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid columns.inserted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'insert-columns', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'ColumnsInserted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: columnStructuralAffectedRanges, mode: 'declared' },
      inverseIds: ['columns.deleted'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'columns.deleted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid columns.deleted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'delete-columns', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'ColumnsDeleted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: columnStructuralAffectedRanges, mode: 'declared' },
      inverseIds: ['columns.inserted', 'cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'row.hidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid row.hidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenRows.add(params.index);
    },
    metadata: {
      schema: { name: 'RowHidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => rowAffectedRange({ sheetId: params.sheetId, row: params.index }), mode: 'declared' },
      inverseIds: ['row.unhidden'],
    },
  });
  runtime.registry.registerMutation<RowsVisibilityParams>({
    id: 'rows.visibility',
    handler: (item, context) => {
      if (!isRowVisibilityMutation(item.params)) throw new Error('Invalid rows.visibility mutation payload');
      const hiddenRows = context.workbook.getSheet(item.params.sheetId).hiddenRows;
      for (const state of item.params.states) {
        if (state.hidden) hiddenRows.add(state.row);
        else hiddenRows.delete(state.row);
      }
    },
    metadata: {
      schema: { name: 'RowsVisibility', validate: isRowVisibilityMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: rowsVisibilityAffectedRanges, mode: 'declared' },
      inverseIds: ['rows.visibility'],
    },
  });
  runtime.registry.registerMutation<ColumnsVisibilityParams>({
    id: 'columns.visibility',
    handler: (item, context) => {
      if (!isColumnVisibilityMutation(item.params)) throw new Error('Invalid columns.visibility mutation payload');
      const hiddenColumns = context.workbook.getSheet(item.params.sheetId).hiddenColumns;
      for (const state of item.params.states) {
        if (state.hidden) hiddenColumns.add(state.column);
        else hiddenColumns.delete(state.column);
      }
    },
    metadata: {
      schema: { name: 'ColumnsVisibility', validate: isColumnVisibilityMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.states.map((state) => columnAffectedRange({ sheetId: params.sheetId, column: state.column })[0]!), mode: 'declared' },
      inverseIds: ['columns.visibility'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'row.unhidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid row.unhidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenRows.delete(params.index);
    },
    metadata: {
      schema: { name: 'RowUnhidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => rowAffectedRange({ sheetId: params.sheetId, row: params.index }), mode: 'declared' },
      inverseIds: ['row.hidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'rows.unhidden.all',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid rows.unhidden.all mutation payload');
      context.workbook.getSheet(item.params.sheetId).hiddenRows.clear();
    },
    metadata: {
      schema: { name: 'RowsUnhiddenAll', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['rows.hidden.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; indices: number[] }>({
    id: 'rows.hidden.restore',
    handler: (item, context) => {
      if (!isSheetIndicesMutation(item.params)) throw new Error('Invalid rows.hidden.restore mutation payload');
      const params = item.params;
      const hiddenRows = context.workbook.getSheet(params.sheetId).hiddenRows;
      hiddenRows.clear();
      for (const index of params.indices) hiddenRows.add(index);
    },
    metadata: {
      schema: { name: 'RowsHiddenRestore', validate: isSheetIndicesMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['rows.unhidden.all'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'column.hidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid column.hidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenColumns.add(params.index);
    },
    metadata: {
      schema: { name: 'ColumnHidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => columnAffectedRange({ sheetId: params.sheetId, column: params.index }), mode: 'declared' },
      inverseIds: ['column.unhidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'column.unhidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid column.unhidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenColumns.delete(params.index);
    },
    metadata: {
      schema: { name: 'ColumnUnhidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => columnAffectedRange({ sheetId: params.sheetId, column: params.index }), mode: 'declared' },
      inverseIds: ['column.hidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'columns.unhidden.all',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid columns.unhidden.all mutation payload');
      context.workbook.getSheet(item.params.sheetId).hiddenColumns.clear();
    },
    metadata: {
      schema: { name: 'ColumnsUnhiddenAll', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['columns.hidden.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; indices: number[] }>({
    id: 'columns.hidden.restore',
    handler: (item, context) => {
      if (!isSheetIndicesMutation(item.params)) throw new Error('Invalid columns.hidden.restore mutation payload');
      const params = item.params;
      const hiddenColumns = context.workbook.getSheet(params.sheetId).hiddenColumns;
      hiddenColumns.clear();
      for (const index of params.indices) hiddenColumns.add(index);
    },
    metadata: {
      schema: { name: 'ColumnsHiddenRestore', validate: isSheetIndicesMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['columns.unhidden.all'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; index: number }>({
    id: 'sheet.row.hide',
    execute: (params, context) => {
      const hiddenRows = context.workbook.getSheet(params.sheetId).hiddenRows;
      if (hiddenRows.has(params.index)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: params.index, endRow: params.index, startColumn: 0, endColumn: Math.max(0, context.workbook.getSheet(params.sheetId).columnCount - 1) }];
      context.applyMutation({
        id: 'row.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'row.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => hiddenRows.add(params.index),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; rows: number[]; hidden: boolean }>({
    id: 'sheet.rows.visibility.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const rows = [...new Set(params.rows)].filter((row) => Number.isSafeInteger(row) && row >= 0 && row < sheet.rowCount);
      const changed = rows.filter((row) => sheet.hiddenRows.has(row) !== params.hidden);
      if (!changed.length) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const states = changed.map((row) => ({ row, hidden: params.hidden }));
      const inverseStates = changed.map((row) => ({ row, hidden: !params.hidden }));
      const affectedRanges = rowsVisibilityAffectedRanges({ sheetId: params.sheetId, states });
      context.applyMutation({
        id: 'rows.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, states }, affectedRanges,
        inverse: [{ id: 'rows.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, states: inverseStates }, affectedRanges }],
        apply: () => { for (const state of states) { if (state.hidden) sheet.hiddenRows.add(state.row); else sheet.hiddenRows.delete(state.row); } },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; columns: number[]; hidden: boolean }>({
    id: 'sheet.columns.visibility.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const columns = [...new Set(params.columns)].filter((column) => Number.isSafeInteger(column) && column >= 0 && column < sheet.columnCount);
      const changed = columns.filter((column) => sheet.hiddenColumns.has(column) !== params.hidden);
      if (!changed.length) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const states = changed.map((column) => ({ column, hidden: params.hidden }));
      const inverseStates = changed.map((column) => ({ column, hidden: !params.hidden }));
      const affectedRanges = states.map((state) => columnAffectedRange({ sheetId: params.sheetId, column: state.column })[0]!);
      context.applyMutation({
        id: 'columns.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, states }, affectedRanges,
        inverse: [{ id: 'columns.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, states: inverseStates }, affectedRanges }],
        apply: () => { for (const state of states) { if (state.hidden) sheet.hiddenColumns.add(state.column); else sheet.hiddenColumns.delete(state.column); } },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; index: number }>({
    id: 'sheet.column.hide',
    execute: (params, context) => {
      const hiddenColumns = context.workbook.getSheet(params.sheetId).hiddenColumns;
      if (hiddenColumns.has(params.index)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, context.workbook.getSheet(params.sheetId).rowCount - 1), startColumn: params.index, endColumn: params.index }];
      context.applyMutation({
        id: 'column.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'column.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => hiddenColumns.add(params.index),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.rows.unhide.all',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = [...sheet.hiddenRows].sort((left, right) => left - right);
      if (previous.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
      context.applyMutation({
        id: 'rows.unhidden.all',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'rows.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => sheet.hiddenRows.clear(),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.columns.unhide.all',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = [...sheet.hiddenColumns].sort((left, right) => left - right);
      if (previous.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
      context.applyMutation({
        id: 'columns.unhidden.all',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'columns.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => sheet.hiddenColumns.clear(),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.rows.insert',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: params.at,
        endRow: params.at + params.count - 1,
        startColumn: 0,
        endColumn: Math.max(0, sheet.columnCount - 1),
      }];
      context.applyMutation({
        id: 'rows.inserted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'rows.deleted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'insert-rows', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.rows.delete',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const end = params.at + params.count - 1;
      const removed = snapshotCellRegion(sheet, params.at, end, 0, Math.max(0, sheet.columnCount - 1));
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: params.at,
        endRow: end,
        startColumn: 0,
        endColumn: Math.max(0, sheet.columnCount - 1),
      }];
      context.applyMutation({
        id: 'rows.deleted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'rows.inserted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
          ...removed.map((entry) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.cell },
            affectedRanges: cellRange({ sheetId: params.sheetId, row: entry.row, column: entry.column }),
          })),
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'delete-rows', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.columns.insert',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: 0,
        endRow: Math.max(0, sheet.rowCount - 1),
        startColumn: params.at,
        endColumn: params.at + params.count - 1,
      }];
      context.applyMutation({
        id: 'columns.inserted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'columns.deleted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'insert-columns', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.columns.delete',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const end = params.at + params.count - 1;
      const removed = snapshotCellRegion(sheet, 0, Math.max(0, sheet.rowCount - 1), params.at, end);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: 0,
        endRow: Math.max(0, sheet.rowCount - 1),
        startColumn: params.at,
        endColumn: end,
      }];
      context.applyMutation({
        id: 'columns.deleted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'columns.inserted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
          ...removed.map((entry) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.cell },
            affectedRanges: cellRange({ sheetId: params.sheetId, row: entry.row, column: entry.column }),
          })),
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'delete-columns', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; indices: number[] }>({
    id: 'sheet.rows.insert.selected',
    execute: (params, context) => executeSelectedDimensionCommand(runtime, 'row', 'insert', params, context),
  });
  runtime.registry.registerCommand<{ sheetId: string; indices: number[] }>({
    id: 'sheet.rows.delete.selected',
    execute: (params, context) => executeSelectedDimensionCommand(runtime, 'row', 'delete', params, context),
  });
  runtime.registry.registerCommand<{ sheetId: string; indices: number[] }>({
    id: 'sheet.columns.insert.selected',
    execute: (params, context) => executeSelectedDimensionCommand(runtime, 'column', 'insert', params, context),
  });
  runtime.registry.registerCommand<{ sheetId: string; indices: number[] }>({
    id: 'sheet.columns.delete.selected',
    execute: (params, context) => executeSelectedDimensionCommand(runtime, 'column', 'delete', params, context),
  });

  // 12. 多列排序 / 转置 / 翻转 / 拆分
  runtime.registry.registerCommand<{
    sheetId: string;
    range: RangeRef;
    criteria: Array<{ column: number; ascending: boolean }>;
    hasHeader: boolean;
  }>({
    id: 'sheet.sort.multi',
    execute: (params, context) => runtime.execute('data.sort.rows', params),
  });

  runtime.registry.registerCommand<{ sheetId: string; row: number; column: number; delimiter: string; maxColumns?: number }>({
    id: 'sheet.splitColumn',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const cell = sheet.cells.get(params.row, params.column);
      const text = cell?.value == null ? '' : String(cell.value);
      const maxColumns = Math.max(2, params.maxColumns ?? 4);
      const parts = text.split(params.delimiter).slice(0, maxColumns);
      if (parts.length <= 1 && parts[0] === text) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const baseStyle = cell?.style ? structuredClone(cell.style) : undefined;
      const values: CellData[][] = [parts.map((part) => ({ value: coerceText(part, cell), style: baseStyle ? structuredClone(baseStyle) : undefined }))];
      while (values[0]!.length < 1) values[0]!.push({ value: null });
      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: params.row,
        startColumn: params.column,
        values,
      });
    },
  });

  // 13. 筛选 / 条件格式 / 数据验证 / 色带 / 名称
  runtime.registry.registerMutation<{ sheetId: string; autoFilter: AutoFilterModel }>({
    id: 'autoFilter.set',
    handler: (item, context) => {
      if (!isFilterMutation(item.params)) throw new Error('Invalid autoFilter.set mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.autoFilter = validateFilterOwnership(sheet, params.autoFilter, { kind: 'worksheet' });
    },
    metadata: {
      schema: { name: 'AutoFilterSet', validate: isFilterMutation },
      permission: { capability: 'sheet.autoFilter.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.autoFilter.range)], mode: 'exact' },
      inverseIds: ['autoFilter.set', 'autoFilter.remove'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; range?: RangeRef }>({
    id: 'autoFilter.remove',
    handler: (item, context) => {
      if (!isFilterRemoveMutation(item.params)) throw new Error('Invalid autoFilter.remove mutation payload');
      context.workbook.getSheet(item.params.sheetId).autoFilter = undefined;
    },
    metadata: {
      schema: { name: 'AutoFilterRemove', validate: isFilterRemoveMutation },
      permission: { capability: 'sheet.autoFilter.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.range ? [structuredClone(params.range)] : [], mode: 'declared' },
      inverseIds: ['autoFilter.set'],
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; autoFilter: AutoFilterModel }>({
    id: 'sheet.autoFilter.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const autoFilter = validateFilterOwnership(sheet, params.autoFilter, { kind: 'worksheet' });
      const previous = sheet.autoFilter;
      const affectedRanges: RangeRef[] = [structuredClone(autoFilter.range)];
      context.applyMutation({
        id: 'autoFilter.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, autoFilter },
        affectedRanges,
        inverse: previous
          ? [{
            id: 'autoFilter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, autoFilter: structuredClone(previous) },
            affectedRanges,
          }]
          : [{
            id: 'autoFilter.remove',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: autoFilter.range },
            affectedRanges,
          }],
        apply: () => {
          sheet.autoFilter = structuredClone(autoFilter);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.autoFilter.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.autoFilter;
      if (!previous) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [structuredClone(previous.range)];
      context.applyMutation({
        id: 'autoFilter.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, range: previous.range },
        affectedRanges,
        inverse: [
          {
            id: 'autoFilter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, autoFilter: structuredClone(previous) },
            affectedRanges,
          },
        ],
        apply: () => {
          context.workbook.getSheet(params.sheetId).autoFilter = undefined;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<AddConditionalFormatParams>({
    id: 'cf.add',
    handler: (item, context) => {
      if (!isConditionalAddMutation(item.params)) throw new Error('Invalid cf.add mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.rule.id);
      if (index >= 0) sheet.conditionalFormats[index] = structuredClone(params.rule);
      else sheet.conditionalFormats.push(structuredClone(params.rule));
    },
    metadata: {
      schema: { name: 'ConditionalFormatAdd', validate: isConditionalAddMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: ruleRanges, mode: 'exact' },
      inverseIds: ['cf.remove'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; ruleId: string; ranges?: RangeRef[] }>({
    id: 'cf.remove',
    handler: (item, context) => {
      if (!isRuleRemoveMutation(item.params)) throw new Error('Invalid cf.remove mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
      if (index >= 0) sheet.conditionalFormats.splice(index, 1);
    },
    metadata: {
      schema: { name: 'ConditionalFormatRemove', validate: isRuleRemoveMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: removeRuleRanges, mode: 'declared' },
      inverseIds: ['cf.add'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; ranges: RangeRef[] }>({
    id: 'cf.clear',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid cf.clear mutation payload');
      context.workbook.getSheet(item.params.sheetId).conditionalFormats.length = 0;
    },
    metadata: {
      schema: { name: 'ConditionalFormatClear', validate: isSheetRangesMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.ranges.map((range) => structuredClone(range)), mode: 'declared' },
      inverseIds: ['cf.add'],
    },
  });
  runtime.registry.registerCommand<AddConditionalFormatParams>({
    id: 'sheet.cf.add',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const normalizedRule = normalizeConditionalFormatRule(params.rule, sheet.conditionalFormats.length + 1);
      const normalizedParams = { ...params, rule: normalizedRule };
      const affectedRanges: RangeRef[] = structuredClone(normalizedRule.ranges);
      context.applyMutation({
        id: 'cf.add',
        unitId: context.workbook.unitId,
        sheetId: params.rule.sheetId,
        params: normalizedParams,
        affectedRanges,
        inverse: [
          {
            id: 'cf.remove',
            unitId: context.workbook.unitId,
            sheetId: params.rule.sheetId,
            params: { sheetId: normalizedRule.sheetId, ruleId: normalizedRule.id, ranges: normalizedRule.ranges },
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(normalizedRule.sheetId);
          const index = target.conditionalFormats.findIndex((rule) => rule.id === normalizedRule.id);
          if (index >= 0) target.conditionalFormats[index] = structuredClone(normalizedRule);
          else target.conditionalFormats.push(structuredClone(normalizedRule));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerMutation<ConditionalFormatUpdateMutationParams>({
    id: 'cf.update',
    handler: (item, context) => {
      if (!isConditionalFormatUpdateMutation(item.params)) throw new Error('Invalid cf.update mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === item.params.before.id);
      if (index < 0 || JSON.stringify(sheet.conditionalFormats[index]) !== JSON.stringify(item.params.before)) throw new Error(`Conditional format ${item.params.before.id} changed before update`);
      sheet.conditionalFormats[index] = structuredClone(item.params.after);
    },
    metadata: {
      schema: { name: 'ConditionalFormatUpdate', validate: isConditionalFormatUpdateMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.ranges.map((range) => structuredClone(range)), mode: 'declared' },
      inverseIds: ['cf.update'],
    },
  });
  runtime.registry.registerCommand<UpdateConditionalFormatParams>({
    id: 'sheet.cf.update',
    execute: (params, context) => {
      if (!isConditionalFormatUpdateParams(params)) throw new Error('Invalid conditional format update parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) throw new Error(`Unknown conditional format rule: ${params.ruleId}`);
      const before = structuredClone(sheet.conditionalFormats[index]!);
      const after = normalizeConditionalFormatRule({ ...before, ...structuredClone(params.patch), id: before.id, sheetId: before.sheetId, ranges: structuredClone(params.patch.ranges ?? before.ranges) }, before.priority ?? index + 1);
      const ranges = [...before.ranges, ...after.ranges].map((range) => structuredClone(range));
      context.applyMutation({
        id: 'cf.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, before, after, ranges },
        affectedRanges: ranges,
        inverse: [{ id: 'cf.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, before: after, after: before, ranges }, affectedRanges: ranges }],
        apply: () => {
          const current = sheet.conditionalFormats.findIndex((rule) => rule.id === before.id);
          if (current < 0 || JSON.stringify(sheet.conditionalFormats[current]) !== JSON.stringify(before)) throw new Error(`Conditional format ${before.id} changed before update`);
          sheet.conditionalFormats[current] = structuredClone(after);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: ranges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleId: string }>({
    id: 'sheet.cf.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.conditionalFormats[index]!);
      const affectedRanges: RangeRef[] = structuredClone(previous.ranges);
      context.applyMutation({
        id: 'cf.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: previous.ranges },
        affectedRanges,
        inverse: [
          {
            id: 'cf.add',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous } satisfies AddConditionalFormatParams,
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId);
          const idx = target.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
          if (idx >= 0) target.conditionalFormats.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.cf.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (sheet.conditionalFormats.length === 0) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const previous = structuredClone(sheet.conditionalFormats);
      const affectedRanges: RangeRef[] = previous.flatMap((rule) => structuredClone(rule.ranges));
      context.applyMutation({
        id: 'cf.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: affectedRanges },
        affectedRanges,
        inverse: previous.map((rule) => ({
          id: 'cf.add' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, rule } satisfies AddConditionalFormatParams,
          affectedRanges: [] as RangeRef[],
        })),
        apply: () => {
          context.workbook.getSheet(params.sheetId).conditionalFormats.length = 0;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<AddDataValidationParams>({
    id: 'dv.add',
    handler: (item, context) => {
      if (!isDataValidationAddMutation(item.params)) throw new Error('Invalid dv.add mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.rule.id);
      if (index >= 0) sheet.dataValidations[index] = structuredClone(params.rule);
      else sheet.dataValidations.push(structuredClone(params.rule));
    },
    metadata: {
      schema: { name: 'DataValidationAdd', validate: isDataValidationAddMutation },
      permission: { capability: 'sheet.data-validation.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: ruleRanges, mode: 'exact' },
      inverseIds: ['dv.remove'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; ruleId: string; ranges?: RangeRef[] }>({
    id: 'dv.remove',
    handler: (item, context) => {
      if (!isRuleRemoveMutation(item.params)) throw new Error('Invalid dv.remove mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.ruleId);
      if (index >= 0) sheet.dataValidations.splice(index, 1);
    },
    metadata: {
      schema: { name: 'DataValidationRemove', validate: isRuleRemoveMutation },
      permission: { capability: 'sheet.data-validation.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: removeRuleRanges, mode: 'declared' },
      inverseIds: ['dv.add'],
    },
  });
  runtime.registry.registerCommand<AddDataValidationParams>({
    id: 'sheet.dv.add',
    execute: (params, context) => {
      const normalizedRule = normalizeDataValidationRule(params.rule);
      const normalizedParams = { ...params, rule: normalizedRule };
      const affectedRanges: RangeRef[] = structuredClone(normalizedRule.ranges);
      context.applyMutation({
        id: 'dv.add',
        unitId: context.workbook.unitId,
        sheetId: params.rule.sheetId,
        params: normalizedParams,
        affectedRanges,
        inverse: [
          {
            id: 'dv.remove',
            unitId: context.workbook.unitId,
            sheetId: params.rule.sheetId,
            params: { sheetId: normalizedRule.sheetId, ruleId: normalizedRule.id, ranges: normalizedRule.ranges },
            affectedRanges,
          },
        ],
        apply: () => {
          const sheet = context.workbook.getSheet(normalizedRule.sheetId);
          const index = sheet.dataValidations.findIndex((rule) => rule.id === normalizedRule.id);
          if (index >= 0) sheet.dataValidations[index] = structuredClone(normalizedRule);
          else sheet.dataValidations.push(structuredClone(normalizedRule));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleId: string }>({
    id: 'sheet.dv.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.dataValidations[index]!);
      const affectedRanges: RangeRef[] = structuredClone(previous.ranges);
      context.applyMutation({
        id: 'dv.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: previous.ranges },
        affectedRanges,
        inverse: [
          {
            id: 'dv.add',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous } satisfies AddDataValidationParams,
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId);
          const idx = target.dataValidations.findIndex((rule) => rule.id === params.ruleId);
          if (idx >= 0) target.dataValidations.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string; rule: BandedRule | null }>({
    id: 'banded.set',
    handler: (item, context) => {
      if (!isBandedMutation(item.params)) throw new Error('Invalid banded.set mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.bandedRule = params.rule ? structuredClone(params.rule) : undefined;
    },
    metadata: {
      schema: { name: 'BandedRuleSet', validate: isBandedMutation },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.rule ? [structuredClone(params.rule.range)] : [], mode: 'declared' },
      inverseIds: ['banded.set'],
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; rule: BandedRule | null }>({
    id: 'sheet.banded.set',
    execute: (params, context) => {
      const previous = context.workbook.getSheet(params.sheetId).bandedRule;
      const affectedRanges: RangeRef[] = params.rule ? [structuredClone(params.rule.range)] : [];
      context.applyMutation({
        id: 'banded.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'banded.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous ? structuredClone(previous) : null },
            affectedRanges,
          },
        ],
        apply: () => {
          const sheet = context.workbook.getSheet(params.sheetId);
          sheet.bandedRule = params.rule ? structuredClone(params.rule) : undefined;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ model: DefinedNameModel }>({
    id: 'name.set',
    handler: (item, context) => {
      if (!isNameSetMutation(item.params)) throw new Error('Invalid name.set mutation payload');
      context.workbook.setDefinedName(item.params.model);
    },
    metadata: {
      schema: { name: 'DefinedNameSet', validate: isNameSetMutation },
      permission: { capability: 'workbook.defined-name.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['name.set', 'name.remove'],
    },
  });
  runtime.registry.registerMutation<{ name: string; scope?: 'workbook' | 'sheet'; sheetId?: string }>({
    id: 'name.remove',
    handler: (item, context) => {
      if (!isNameRemoveMutation(item.params)) throw new Error('Invalid name.remove mutation payload');
      const params = item.params;
      context.workbook.removeDefinedName(params.name, params.scope ?? 'workbook', params.sheetId);
    },
    metadata: {
      schema: { name: 'DefinedNameRemove', validate: isNameRemoveMutation },
      permission: { capability: 'workbook.defined-name.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['name.set'],
    },
  });
  runtime.registry.registerCommand<{
    name: string;
    value?: string;
    formula?: string;
    scope?: 'workbook' | 'sheet';
    sheetId?: string;
    hidden?: boolean;
    comment?: string;
  }>({
    id: 'workbook.name.set',
    execute: (params, context) => {
      const model: DefinedNameModel = {
        name: params.name,
        formula: params.formula ?? params.value ?? '',
        scope: params.scope ?? 'workbook',
        ...(params.sheetId ? { sheetId: params.sheetId } : {}),
        ...(params.hidden === undefined ? {} : { hidden: params.hidden }),
        ...(params.comment === undefined ? {} : { comment: params.comment }),
      };
      // Validate before opening a mutation so invalid scope/name input cannot
      // create a history entry or leave the legacy formula view half updated.
      const normalized = normalizeDefinedNameModel(model);
      const previous = context.workbook.getDefinedNameExact(normalized.name, normalized.scope, normalized.sheetId);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'name.set',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params: { model: normalized },
        affectedRanges,
        inverse: previous !== undefined
          ? [{ id: 'name.set', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { model: previous }, affectedRanges }]
          : [{ id: 'name.remove', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { name: normalized.name, scope: normalized.scope, sheetId: normalized.sheetId }, affectedRanges }],
        apply: () => {
          context.workbook.setDefinedName(normalized);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ name: string; scope?: 'workbook' | 'sheet'; sheetId?: string }>({
    id: 'workbook.name.remove',
    execute: (params, context) => {
      const previous = context.workbook.getDefinedNameExact(params.name, params.scope ?? 'workbook', params.sheetId);
      if (previous === undefined || previous.scope !== (params.scope ?? 'workbook') || previous.sheetId !== params.sheetId) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'name.remove',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params,
        affectedRanges,
        inverse: [
          { id: 'name.set', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { model: previous }, affectedRanges },
        ],
        apply: () => {
          context.workbook.removeDefinedName(params.name, previous.scope, previous.sheetId);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
}

function coerceText(text: string, previousCell: CellData | undefined): CellData['value'] {
  if (typeof previousCell?.value === 'number') {
    const numeric = Number(text.replace(/[$,%]/g, ''));
    if (Number.isFinite(numeric) && text.trim() !== '') return numeric;
  }
  return text;
}
