import type {
  CellData,
  CellStyle,
  CellHyperlink,
  CellNote,
  CommentThread,
  ConditionalFormatRule,
  DataValidationRule,
  WorksheetPane,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
  BorderLine,
  BorderPlacement,
} from '@react-sheets/core-model';
import { cellKey, clearFormulaProvenance, columnLabel, planCellShift, sheetRuleRegistry, type CellShiftSpec } from '@react-sheets/core-model';
import { StructuralTransform } from '@react-sheets/core-model';
import { formatValue } from '@react-sheets/number-format';
import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import {
  copyRangeToClipboardData,
  parseClipboardPayload,
  shiftFormula,
  type ClipboardTransfer,
  type ClipboardPayload,
  type PasteSpecialSpec,
  isPasteSpecialSpecSupported,
} from '../clipboard';
import { isCellInputInterpretationContext, type CellInputInterpretationContext } from '../text-input';
import type { CellEntryIntent } from '../index';

export type FreezePreset = 'none' | 'firstRow' | 'firstColumn' | 'both';
export type GoToSpecialKind =
  | 'blanks'
  | 'constants'
  | 'formulas'
  | 'comments'
  | 'comments-notes'
  | 'visible'
  | 'errors'
  | 'conditional-format'
  | 'data-validation'
  | 'current-region'
  | 'last-cell'
  | 'objects';

export interface MultiRangeStyleParams {
  sheetId: string;
  ranges: RangeRef[];
  style: Partial<CellStyle>;
}

export interface PasteRangeParams {
  sheetId: string;
  targetOrigin: { row: number; column: number };
  clipboard: ClipboardPayload;
  inputContext: CellInputInterpretationContext;
  entryIntent: CellEntryIntent;
  transfer: ClipboardTransfer;
  spec: PasteSpecialSpec;
}

export interface CutPasteRangeParams extends PasteRangeParams {
  /** Explicit source range is required for a cut transaction. */
  sourceRange: RangeRef;
}

export interface FormatCellsParams {
  sheetId: string;
  ranges: RangeRef[];
  numberFormat?: string;
  style?: Partial<CellStyle>;
  border?: { placement: BorderPlacement; line?: BorderLine };
}

export interface GoToParams {
  sheetId: string;
  reference: string;
}

export interface GoToSpecialParams {
  sheetId: string;
  range: RangeRef;
  kind: GoToSpecialKind;
}

export interface CellShiftParams extends CellShiftSpec {
  affectedBand: RangeRef;
}

export interface SheetViewParams {
  sheetId: string;
  showGridlines?: boolean;
  showHeaders?: boolean;
  zoom?: number;
}

function normalizeRanges(ranges: RangeRef[]): RangeRef[] {
  return ranges.map((range) => ({
    sheetId: range.sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  }));
}

function forEachCell(
  sheet: WorksheetModel,
  range: RangeRef,
  fn: (row: number, column: number, cell: CellData | undefined) => void,
): void {
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      fn(row, column, sheet.cells.get(row, column));
    }
  }
}

function parseA1Reference(reference: string): { row: number; column: number } | null {
  const match = reference.trim().match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) return null;
  const colLetters = match[1]!.toUpperCase();
  let column = 0;
  for (const char of colLetters) column = column * 26 + char.charCodeAt(0) - 64;
  column -= 1;
  const row = Number(match[2]) - 1;
  if (!Number.isFinite(row) || row < 0 || column < 0) return null;
  return { row, column };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
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
  return isRecord(value) && 'value' in value;
}

function isSheetViewMutation(value: unknown): value is SheetViewParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.showGridlines === undefined || typeof value.showGridlines === 'boolean')
    && (value.showHeaders === undefined || typeof value.showHeaders === 'boolean')
    && (value.zoom === undefined || (typeof value.zoom === 'number' && Number.isFinite(value.zoom) && value.zoom > 0));
}

type PasteMutationParams = PasteRangeParams & {
  sourceExtent: { rows: number; columns: number };
  sourceRange?: RangeRef;
  clearSource?: boolean;
  snapshot: PasteSnapshot;
  sourceSnapshot?: PasteSnapshot;
};

interface CellSnapshot {
  row: number;
  column: number;
  value?: CellData;
}

interface MetadataSnapshot<T> {
  key: string;
  value?: T;
}

interface PasteSnapshot {
  clearRanges?: RangeRef[];
  clearMetadataRanges?: RangeRef[];
  cells: CellSnapshot[];
  notes?: MetadataSnapshot<CellNote>[];
  hyperlinks?: MetadataSnapshot<CellHyperlink>[];
  commentCells?: string[];
  comments?: CommentThread[];
  validations?: DataValidationRule[];
  conditionalFormats?: ConditionalFormatRule[];
  columnWidths?: Array<{ column: number; widthPx?: number }>;
}

function isPasteMutation(value: unknown): value is PasteMutationParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && isRecord(value.targetOrigin) && Number.isInteger(value.targetOrigin.row) && Number.isInteger(value.targetOrigin.column)
    && isCellInputInterpretationContext(value.inputContext)
    && isRecord(value.entryIntent) && value.entryIntent.kind === 'paste'
    && (value.transfer === 'copy' || value.transfer === 'move')
    && isRecord(value.clipboard) && value.clipboard.transfer === value.transfer
    && value.clipboard.schema === 'SparseClipboardPayload'
    && isRecord(value.clipboard.sourceExtent)
    && Number.isInteger(value.clipboard.sourceExtent.rows) && Number(value.clipboard.sourceExtent.rows) > 0
    && Number.isInteger(value.clipboard.sourceExtent.columns) && Number(value.clipboard.sourceExtent.columns) > 0
    && Array.isArray(value.clipboard.occupiedCells)
    && value.clipboard.occupiedCells.every((cell) => isRecord(cell)
      && Number.isInteger(cell.rowOffset) && Number(cell.rowOffset) >= 0 && Number(cell.rowOffset) < Number(value.clipboard.sourceExtent!.rows)
      && Number.isInteger(cell.columnOffset) && Number(cell.columnOffset) >= 0 && Number(cell.columnOffset) < Number(value.clipboard.sourceExtent!.columns)
      && isCellData(cell.value))
    && isRecord(value.clipboard.rangeMetadata)
    && Array.isArray(value.clipboard.rangeMetadata.columnWidths)
    && Array.isArray(value.clipboard.rangeMetadata.validations)
    && Array.isArray(value.clipboard.rangeMetadata.conditionalFormats)
    && Array.isArray(value.clipboard.rangeMetadata.notes)
    && Array.isArray(value.clipboard.rangeMetadata.comments)
    && Array.isArray(value.clipboard.rangeMetadata.hyperlinks)
    && isPasteSpecialSpec(value.spec)
    && isPasteSpecialSpecSupported(value.spec, value.clipboard as unknown as ClipboardPayload)
    && isRecord(value.sourceExtent) && Number.isInteger(value.sourceExtent.rows) && Number.isInteger(value.sourceExtent.columns)
    && isPasteSnapshot(value.snapshot)
    && (value.sourceSnapshot === undefined || isPasteSnapshot(value.sourceSnapshot))
    && (value.transfer === 'move'
      ? isRange(value.sourceRange) && value.clearSource === true && (value.sourceRange.sheetId === value.sheetId || isPasteSnapshot(value.sourceSnapshot))
      : value.sourceRange === undefined && value.clearSource === false && value.sourceSnapshot === undefined);
}

function pasteAffectedRanges(value: PasteMutationParams): RangeRef[] {
  const rowCount = value.spec.transpose ? value.sourceExtent.columns : value.sourceExtent.rows;
  const columnCount = value.spec.transpose ? value.sourceExtent.rows : value.sourceExtent.columns;
  const ranges = [{ sheetId: value.sheetId, startRow: value.targetOrigin.row, endRow: value.targetOrigin.row + Math.max(0, rowCount - 1), startColumn: value.targetOrigin.column, endColumn: value.targetOrigin.column + Math.max(0, columnCount - 1) }];
  if (value.clearSource && value.sourceRange) ranges.push(structuredClone(value.sourceRange));
  return ranges;
}

function isPasteSpecialSpec(value: unknown): value is PasteSpecialSpec {
  if (!isRecord(value)) return false;
  const metadata = value.metadata;
  return (value.content === 'none' || value.content === 'all' || value.content === 'values' || value.content === 'formulas')
    && (value.formatting === 'all' || value.formatting === 'none' || value.formatting === 'number-format' || value.formatting === 'source-formatting' || value.formatting === 'all-except-borders' || value.formatting === 'source-theme')
    && isRecord(metadata)
    && typeof metadata.commentsNotes === 'boolean'
    && typeof metadata.validation === 'boolean'
    && typeof metadata.columnWidths === 'boolean'
    && typeof metadata.conditionalFormats === 'boolean'
    && typeof metadata.hyperlinks === 'boolean'
    && (value.operation === 'none' || value.operation === 'add' || value.operation === 'subtract' || value.operation === 'multiply' || value.operation === 'divide')
    && typeof value.skipBlanks === 'boolean'
    && typeof value.transpose === 'boolean'
    && typeof value.link === 'boolean';
}

function isPasteSnapshot(value: unknown): value is PasteSnapshot {
  if (!isRecord(value) || !Array.isArray(value.cells)) return false;
  return (value.clearRanges === undefined || (Array.isArray(value.clearRanges) && value.clearRanges.every(isRange)))
    && (value.clearMetadataRanges === undefined || (Array.isArray(value.clearMetadataRanges) && value.clearMetadataRanges.every(isRange)))
    && value.cells.every((entry) => isRecord(entry) && Number.isInteger(entry.row) && Number.isInteger(entry.column) && (entry.value === undefined || isCellData(entry.value)))
    && (value.notes === undefined || Array.isArray(value.notes))
    && (value.hyperlinks === undefined || Array.isArray(value.hyperlinks))
    && (value.commentCells === undefined || Array.isArray(value.commentCells))
    && (value.comments === undefined || Array.isArray(value.comments))
    && (value.validations === undefined || Array.isArray(value.validations))
    && (value.conditionalFormats === undefined || Array.isArray(value.conditionalFormats))
    && (value.columnWidths === undefined || Array.isArray(value.columnWidths));
}

function isCellShiftMutation(value: unknown): value is CellShiftParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range)
    && isRange(value.affectedBand)
    && value.range.sheetId === value.sheetId
    && value.affectedBand.sheetId === value.sheetId
    && (value.operation === 'insert' || value.operation === 'delete')
    && (value.axis === 'row' || value.axis === 'column');
}

type CellShiftRestoreParams = { spec: CellShiftParams; cells: Array<{ row: number; column: number; cell: CellData }> };

function isCellShiftRestoreMutation(value: unknown): value is CellShiftRestoreParams {
  return isRecord(value) && isCellShiftMutation(value.spec)
    && Array.isArray(value.cells) && value.cells.every((entry) => isRecord(entry) && Number.isInteger(entry.row) && Number.isInteger(entry.column) && isCellData(entry.cell));
}

function isSheetDuplicateMutation(value: unknown): value is { sourceSheetId: string; newId: string; newName: string } {
  return isRecord(value) && typeof value.sourceSheetId === 'string' && typeof value.newId === 'string' && typeof value.newName === 'string' && value.newId.length > 0;
}

function isSheetIdMutation(value: unknown): value is { sheetId: string } {
  return isRecord(value) && typeof value.sheetId === 'string' && value.sheetId.length > 0;
}

function isSheetReorderedMutation(value: unknown): value is { sheetId: string; toIndex: number } {
  return isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.toIndex) && Number(value.toIndex) >= 0;
}

function isTabColorMutation(value: unknown): value is { sheetId: string; color?: string } {
  return isRecord(value) && typeof value.sheetId === 'string' && (value.color === undefined || typeof value.color === 'string');
}

export function resolveGoTo(workbook: WorkbookModel, params: GoToParams): { row: number; column: number } | null {
  const resolved = resolveGoToRange(workbook, params);
  return resolved ? { row: resolved.startRow, column: resolved.startColumn } : null;
}

/** Resolve a Go To reference without losing a multi-cell selection. */
export function resolveGoToRange(workbook: WorkbookModel, params: GoToParams): RangeRef | null {
  let reference = params.reference.trim();
  let sheetId = params.sheetId;
  const named = workbook.getDefinedName(reference, sheetId)?.formula;
  if (named) reference = named;
  const qualified = reference.match(/^(?:'([^']+)'|([^!]+))!(.+)$/);
  if (qualified) {
    const sheetName = (qualified[1] ?? qualified[2] ?? '').trim();
    const targetSheet = workbook.getSheets().find((sheet) => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase() || sheet.id.toLocaleLowerCase() === sheetName.toLocaleLowerCase());
    if (!targetSheet) return null;
    sheetId = targetSheet.id;
    reference = qualified[3]!.trim();
  }
  const parts = reference.replace(/\$/g, '').split(':');
  const start = parseA1Reference(parts[0] ?? '');
  if (!start) return null;
  const end = parts.length > 1 ? parseA1Reference(parts[1] ?? '') : start;
  if (!end) return null;
  return {
    sheetId,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  };
}

function detectCurrentRegion(sheet: WorksheetModel, row: number, column: number): RangeRef {
  const hasValue = (targetRow: number, targetColumn: number): boolean => {
    const cell = sheet.cells.get(targetRow, targetColumn);
    return Boolean(cell && (cell.value !== null && cell.value !== undefined && cell.value !== '' || cell.formula));
  };
  if (!hasValue(row, column)) return { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column };
  let startRow = row;
  let endRow = row;
  let startColumn = column;
  let endColumn = column;
  while (startRow > 0 && hasValue(startRow - 1, column)) startRow -= 1;
  while (endRow + 1 < sheet.rowCount && hasValue(endRow + 1, column)) endRow += 1;
  while (startColumn > 0 && hasValue(row, startColumn - 1)) startColumn -= 1;
  while (endColumn + 1 < sheet.columnCount && hasValue(row, endColumn + 1)) endColumn += 1;
  // Expand through rows/columns that are connected to the first discovered
  // rectangle. This handles a populated rectangular table with an empty
  // corner in the active row without scanning the full worksheet.
  let changed = true;
  while (changed) {
    changed = false;
    if (startRow > 0 && Array.from({ length: endColumn - startColumn + 1 }, (_, offset) => hasValue(startRow - 1, startColumn + offset)).some(Boolean)) { startRow -= 1; changed = true; }
    if (endRow + 1 < sheet.rowCount && Array.from({ length: endColumn - startColumn + 1 }, (_, offset) => hasValue(endRow + 1, startColumn + offset)).some(Boolean)) { endRow += 1; changed = true; }
    if (startColumn > 0 && Array.from({ length: endRow - startRow + 1 }, (_, offset) => hasValue(startRow + offset, startColumn - 1)).some(Boolean)) { startColumn -= 1; changed = true; }
    if (endColumn + 1 < sheet.columnCount && Array.from({ length: endRow - startRow + 1 }, (_, offset) => hasValue(startRow + offset, endColumn + 1)).some(Boolean)) { endColumn += 1; changed = true; }
  }
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

export function resolveGoToSpecial(
  workbook: WorkbookModel,
  params: GoToSpecialParams,
): RangeRef[] {
  const sheet = workbook.getSheet(params.sheetId);
  const normalizedRange: RangeRef = {
    sheetId: params.sheetId,
    startRow: Math.min(params.range.startRow, params.range.endRow),
    endRow: Math.max(params.range.startRow, params.range.endRow),
    startColumn: Math.min(params.range.startColumn, params.range.endColumn),
    endColumn: Math.max(params.range.startColumn, params.range.endColumn),
  };
  if (params.kind === 'current-region') {
    const anchor = { row: normalizedRange.startRow, column: normalizedRange.startColumn };
    return [detectCurrentRegion(sheet, anchor.row, anchor.column)];
  }
  if (params.kind === 'last-cell') {
    let lastRow = 0;
    let lastColumn = 0;
    sheet.cells.forEach((cell, row, column) => {
      if (cell && (cell.value !== null && cell.value !== undefined || cell.formula || cell.style || cell.numberFormat)) {
        if (row > lastRow || (row === lastRow && column > lastColumn)) {
          lastRow = row;
          lastColumn = column;
        }
      }
    });
    return [{ sheetId: params.sheetId, startRow: lastRow, endRow: lastRow, startColumn: lastColumn, endColumn: lastColumn }];
  }
  const hits: RangeRef[] = [];
  const conditionalCells = new Set<string>();
  if (params.kind === 'conditional-format') {
    for (const rule of sheet.conditionalFormats) {
      for (const range of rule.ranges) {
        const startRow = Math.max(normalizedRange.startRow, range.startRow);
        const endRow = Math.min(normalizedRange.endRow, range.endRow);
        const startColumn = Math.max(normalizedRange.startColumn, range.startColumn);
        const endColumn = Math.min(normalizedRange.endColumn, range.endColumn);
        for (let row = startRow; row <= endRow; row += 1) for (let column = startColumn; column <= endColumn; column += 1) conditionalCells.add(`${row}:${column}`);
      }
    }
  }
  const validationCells = new Set<string>();
  if (params.kind === 'data-validation') {
    for (const rule of sheet.dataValidations) {
      for (const range of rule.ranges) {
        const startRow = Math.max(normalizedRange.startRow, range.startRow);
        const endRow = Math.min(normalizedRange.endRow, range.endRow);
        const startColumn = Math.max(normalizedRange.startColumn, range.startColumn);
        const endColumn = Math.min(normalizedRange.endColumn, range.endColumn);
        for (let row = startRow; row <= endRow; row += 1) for (let column = startColumn; column <= endColumn; column += 1) validationCells.add(`${row}:${column}`);
      }
    }
  }
  for (let row = normalizedRange.startRow; row <= normalizedRange.endRow; row++) {
    for (let column = normalizedRange.startColumn; column <= normalizedRange.endColumn; column++) {
      const cell = sheet.cells.get(row, column);
      let match = false;
      switch (params.kind) {
        case 'blanks':
          match = !cell || cell.value == null || cell.value === '';
          break;
        case 'constants':
          match = Boolean(cell && cell.value != null && cell.value !== '' && !cell.formula);
          break;
        case 'formulas':
          match = Boolean(cell?.formula);
          break;
        case 'comments':
        case 'comments-notes':
          match = sheet.review.getThreadsAt(row, column).length > 0 || sheet.review.hasNoteAt(row, column);
          break;
        case 'errors':
          match = Boolean(typeof cell?.value === 'string' && cell.value.startsWith('#'));
          break;
        case 'visible':
          match = !sheet.hiddenRows.has(row) && !sheet.hiddenColumns.has(column);
          break;
        case 'conditional-format':
          match = conditionalCells.has(`${row}:${column}`);
          break;
        case 'data-validation':
          match = validationCells.has(`${row}:${column}`);
          break;
        case 'objects':
          match = sheet.drawings.some((drawing) => drawing.anchor.kind !== 'absolute'
            && drawing.anchor.row !== undefined && drawing.anchor.column !== undefined
            && drawing.anchor.row >= row && drawing.anchor.row <= row
            && drawing.anchor.column >= column && drawing.anchor.column <= column);
          break;
      }
      if (match) {
        hits.push({ sheetId: params.sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column });
      }
    }
  }
  return hits;
}

function applyPasteCell(
  spec: PasteSpecialSpec,
  transfer: ClipboardTransfer,
  source: CellData,
  target: CellData | undefined,
  rowDelta: number,
  colDelta: number,
  sourceAddress: string,
): CellData | undefined {
  const destination = target ? structuredClone(target) : { value: null };
  const sourceIsBlank = source.value === null || source.value === undefined;
  if (spec.skipBlanks && sourceIsBlank && !source.formula) return undefined;
  if (spec.link) {
    return { value: null, formula: `=${sourceAddress}` };
  }
  if (spec.content === 'none' && spec.operation === 'none') {
    if (spec.formatting === 'none') return undefined;
    return {
      ...destination,
      style: spec.formatting === 'all-except-borders' && source.style
        ? { ...structuredClone(source.style), borders: destination.style?.borders }
        : source.style ? structuredClone(source.style) : destination.style,
      numberFormat: spec.formatting === 'number-format' || spec.formatting === 'source-formatting' || spec.formatting === 'all' ? source.numberFormat : destination.numberFormat,
    };
  }
  const sourceFormula = source.formula
    ? transfer === 'move' ? source.formula : shiftFormula(source.formula, rowDelta, colDelta)
    : undefined;

  if (spec.operation !== 'none') {
    if (sourceFormula || target?.formula) throw new Error('Paste arithmetic cannot operate on formula cells');
    const sourceValue = source.value;
    const targetValue = target?.value;
    if (sourceIsBlank && spec.skipBlanks) return undefined;
    if (typeof sourceValue !== 'number' || (targetValue !== null && targetValue !== undefined && typeof targetValue !== 'number')) {
      throw new Error(`Paste arithmetic ${spec.operation} requires numeric source and target values`);
    }
    const left = typeof targetValue === 'number' ? targetValue : 0;
    const right = sourceValue;
    if (spec.operation === 'divide' && right === 0) throw new Error('Paste arithmetic divide cannot use zero');
    const value = spec.operation === 'add' ? left + right
      : spec.operation === 'subtract' ? left - right
        : spec.operation === 'multiply' ? left * right
          : left / right;
    return { ...clearFormulaProvenance(destination), value, formula: undefined };
  }

  if (spec.content === 'values') {
    // Values means values only: no formula, style, number format or cached
    // display metadata may leak into the destination.
    const next: CellData = { value: source.value ?? null };
    if (spec.formatting === 'number-format') next.numberFormat = source.numberFormat;
    if (spec.formatting === 'source-formatting' || spec.formatting === 'all') next.style = source.style ? structuredClone(source.style) : undefined;
    if (spec.formatting === 'all-except-borders' && source.style) next.style = { ...structuredClone(source.style), borders: destination.style?.borders };
    return next;
  }
  if (spec.content === 'formulas') {
    if (!sourceFormula) return { ...clearFormulaProvenance(destination), value: source.value ?? null, formula: undefined };
    const next = clearFormulaProvenance(destination);
    return {
      ...next,
      value: null,
      formula: sourceFormula,
      formulaValue: undefined,
      ...(spec.formatting === 'none' ? { style: destination.style, numberFormat: destination.numberFormat } : {}),
      ...(spec.formatting === 'number-format' ? { numberFormat: source.numberFormat } : {}),
      ...(spec.formatting === 'source-formatting' || spec.formatting === 'all' ? { style: source.style ? structuredClone(source.style) : undefined } : {}),
      ...(spec.formatting === 'all-except-borders' && source.style ? { style: { ...structuredClone(source.style), borders: destination.style?.borders } } : {}),
    };
  }
  if (spec.formatting === 'none') {
    return { ...clearFormulaProvenance(destination), value: source.value ?? null, formula: sourceFormula };
  }
  if (spec.formatting === 'number-format') {
    return {
      ...clearFormulaProvenance(destination),
      value: source.value ?? null,
      formula: sourceFormula,
      numberFormat: source.numberFormat,
    };
  }
  const next = clearFormulaProvenance(source);
  if (sourceFormula) next.formula = sourceFormula;
  if (spec.formatting === 'all-except-borders' && next.style) {
    next.style = { ...next.style, borders: destination.style?.borders };
  }
  return next;
}

function rangeContains(range: RangeRef, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId && left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
}

function assertPastePreconditions(workbook: WorkbookModel, params: PasteRangeParams): RangeRef {
  const sheet = workbook.getSheet(params.sheetId);
  if (!Number.isInteger(params.targetOrigin.row) || !Number.isInteger(params.targetOrigin.column)
    || params.targetOrigin.row < 0 || params.targetOrigin.column < 0) throw new Error('Paste target origin is invalid');
  const sourceRange = params.clipboard.range;
  const sourceSheet = workbook.getSheet(sourceRange.sheetId);
  const sourceRows = params.clipboard.sourceExtent.rows;
  const sourceColumns = params.clipboard.sourceExtent.columns;
  if (sourceRows === 0 || sourceColumns === 0) throw new Error('Clipboard payload contains no cells');
  if (!Number.isSafeInteger(sourceRows) || !Number.isSafeInteger(sourceColumns) || sourceRows < 0 || sourceColumns < 0) throw new Error('Clipboard source extent is invalid');
  const targetRange: RangeRef = {
    sheetId: params.sheetId,
    startRow: params.targetOrigin.row,
    endRow: params.targetOrigin.row + Math.max(0, (params.spec.transpose ? sourceColumns : sourceRows) - 1),
    startColumn: params.targetOrigin.column,
    endColumn: params.targetOrigin.column + Math.max(0, (params.spec.transpose ? sourceRows : sourceColumns) - 1),
  };
  if (targetRange.endRow < targetRange.startRow || targetRange.endColumn < targetRange.startColumn) throw new Error('Paste target extent is invalid');
  if (targetRange.endRow > 1048575 || targetRange.endColumn > 16383) throw new Error('Paste exceeds canonical worksheet limits');
  if (sourceRange.sheetId.length === 0 || sourceRange.startRow < 0 || sourceRange.startColumn < 0) throw new Error('Clipboard source range is invalid');
  if (sourceRange.endRow < sourceRange.startRow || sourceRange.endColumn < sourceRange.startColumn) throw new Error('Clipboard source range is invalid');
  if (params.spec.formatting === 'source-theme' && !params.clipboard.rangeMetadata.sourceWorkbookThemeRef) throw new Error('Paste source theme is unavailable for this clipboard payload');
  if (!isPasteSpecialSpecSupported(params.spec, params.clipboard)) throw new Error('Paste Special option is not supported by the canonical workbook model');
  if (params.spec.metadata.validation && params.clipboard.rangeMetadata.validations.length === 0 && params.spec.content === 'all') {
    // Empty metadata is a valid no-op; the branch is intentionally explicit so
    // malformed hosts cannot omit the metadata envelope.
    if (!params.clipboard.rangeMetadata) throw new Error('Clipboard range metadata is required');
  }
  const protectedRange = sheet.protectionRules.find((rule) => rule.locked && rule.range && rangesIntersect(rule.range, targetRange));
  if (protectedRange) throw new Error(`Paste target is protected by ${protectedRange.id}`);
  if (params.transfer === 'move' && sourceRange.sheetId === params.sheetId && rangesIntersect(sourceRange, targetRange)) {
    throw new Error('Cut source and target ranges may not overlap');
  }
  return targetRange;
}

function keyFor(row: number, column: number): string {
  return `${row}:${column}`;
}

function snapshotCells(sheet: WorksheetModel, ranges: RangeRef[]): CellSnapshot[] {
  const output: CellSnapshot[] = [];
  const seen = new Set<string>();
  for (const range of ranges) {
    sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (value, row, column) => {
      const key = keyFor(row, column);
      if (seen.has(key)) return;
      seen.add(key);
      output.push({ row, column, value: structuredClone(value) });
    });
  }
  return output;
}

function snapshotMetadata(sheet: WorksheetModel, ranges: RangeRef[], include: PasteSpecialSpec['metadata']): Pick<PasteSnapshot, 'notes' | 'hyperlinks' | 'commentCells' | 'comments'> {
  const contains = (row: number, column: number) => ranges.some((range) => rangeContains(range, row, column));
  const notes = include.commentsNotes ? sheet.review.noteEntries().filter((entry) => contains(entry.row, entry.column)).map((entry) => ({ key: entry.key, value: entry.note })) : undefined;
  const hyperlinks = include.hyperlinks ? [...sheet.hyperlinks.entries()].filter(([key]) => {
    const [row, column] = key.split(':').map(Number);
    return Number.isInteger(row) && Number.isInteger(column) && contains(row, column);
  }).map(([key, value]) => ({ key, value: structuredClone(value) })) : undefined;
  const commentCells = include.commentsNotes ? sheet.review.threadEntries().filter((thread) => contains(thread.row, thread.column)).map((thread) => keyFor(thread.row, thread.column)) : undefined;
  const comments = include.commentsNotes ? sheet.review.threadEntries().filter((thread) => contains(thread.row, thread.column)) : undefined;
  return { notes, hyperlinks, commentCells, comments };
}

function applyPasteSnapshot(sheet: WorksheetModel, snapshot: PasteSnapshot): void {
  for (const range of [...(snapshot.clearRanges ?? []), ...(snapshot.clearMetadataRanges ?? [])]) {
    if (range.sheetId !== sheet.id) continue;
    sheet.rowCount = Math.max(sheet.rowCount, range.endRow + 1);
    sheet.columnCount = Math.max(sheet.columnCount, range.endColumn + 1);
  }
  for (const range of snapshot.clearRanges ?? []) {
    if (range.sheetId !== sheet.id) continue;
    sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (_value, row, column) => sheet.cells.delete(row, column));
  }
  for (const range of snapshot.clearMetadataRanges ?? []) {
    if (range.sheetId !== sheet.id) continue;
    for (const entry of sheet.review.noteEntries()) if (rangeContains(range, entry.row, entry.column)) sheet.review.removeNote(entry.row, entry.column);
    for (const key of [...sheet.hyperlinks.keys()]) {
      const [row, column] = key.split(':').map(Number);
      if (Number.isInteger(row) && Number.isInteger(column) && rangeContains(range, row, column)) sheet.hyperlinks.delete(key);
    }
    for (const thread of sheet.review.threadEntries()) if (rangeContains(range, thread.row, thread.column)) sheet.review.removeThread(thread.id);
  }
  for (const cell of snapshot.cells) {
    if (cell.value) sheet.cells.set(cell.row, cell.column, structuredClone(cell.value));
    else if (!(snapshot.clearRanges ?? []).some((range) => rangeContains(range, cell.row, cell.column))) sheet.cells.delete(cell.row, cell.column);
  }
  if (snapshot.notes) {
    for (const entry of snapshot.notes) {
      const [row, column] = entry.key.split(':').map(Number);
      if (entry.value) sheet.review.setNote(row!, column!, entry.value);
      else sheet.review.removeNote(row!, column!);
    }
  }
  if (snapshot.hyperlinks) {
    for (const entry of snapshot.hyperlinks) {
      if (entry.value) sheet.hyperlinks.set(entry.key, structuredClone(entry.value));
      else sheet.hyperlinks.delete(entry.key);
    }
  }
  if (snapshot.comments || snapshot.commentCells) {
    const covered = new Set(snapshot.commentCells ?? snapshot.comments?.map((entry) => keyFor(entry.row, entry.column)) ?? []);
    for (const thread of sheet.review.threadEntries()) if (covered.has(keyFor(thread.row, thread.column))) sheet.review.removeThread(thread.id);
    for (const thread of snapshot.comments ?? []) sheet.review.addThread(thread);
  }
  if (snapshot.validations) {
    sheet.dataValidations.length = 0;
    sheet.dataValidations.push(...structuredClone(snapshot.validations));
  }
  if (snapshot.conditionalFormats) {
    sheet.conditionalFormats.length = 0;
    sheet.conditionalFormats.push(...structuredClone(snapshot.conditionalFormats));
  }
  if (snapshot.columnWidths) {
    for (const entry of snapshot.columnWidths) {
      if (entry.widthPx === undefined) delete sheet.columnWidthsPx[entry.column];
      else sheet.columnWidthsPx[entry.column] = entry.widthPx;
    }
  }
}

function applyPasteMetadataPlan(workbook: WorkbookModel, params: PasteRangeParams, targetRange: RangeRef, after: PasteSnapshot): void {
  const source = params.clipboard.range;
  const sourceSheet = workbook.getSheet(source.sheetId);
  const targetSheet = workbook.getSheet(params.sheetId);
  const metadata = params.clipboard.rangeMetadata;
  if (params.spec.metadata.commentsNotes) {
    const notes = after.notes ?? [];
    for (const entry of metadata.notes) {
      const row = targetRange.startRow + entry.rowOffset;
      const column = targetRange.startColumn + entry.columnOffset;
      notes.push({ key: keyFor(row, column), value: structuredClone(entry.value) });
    }
    after.notes = notes;
    const comments = after.comments ?? [];
    for (const entry of metadata.comments) {
      comments.push({ ...structuredClone(entry.value), sheetId: params.sheetId, row: targetRange.startRow + entry.rowOffset, column: targetRange.startColumn + entry.columnOffset });
    }
    after.comments = comments;
  }
  if (params.spec.metadata.hyperlinks) {
    const hyperlinks = after.hyperlinks ?? [];
    for (const entry of metadata.hyperlinks) {
      hyperlinks.push({ key: keyFor(targetRange.startRow + entry.rowOffset, targetRange.startColumn + entry.columnOffset), value: structuredClone(entry.value) });
    }
    after.hyperlinks = hyperlinks;
  }
  if (params.spec.metadata.validation) {
    let targetRules = sheetRuleRegistry.cropRules(after.validations ?? [], targetRange);
    if (params.transfer === 'move' && source.sheetId === params.sheetId) {
      targetRules = sheetRuleRegistry.cropRules(targetRules, source);
    }
    const sourceRules = sheetRuleRegistry.cloneRulesForPaste(metadata.validations, {
      source,
      target: targetRange,
      transpose: params.spec.transpose,
      id: (rule) => `${rule.id}@paste:${targetRange.startRow}:${targetRange.startColumn}`,
    });
    after.validations = [...targetRules, ...sourceRules];
  }
  if (params.spec.metadata.conditionalFormats) {
    let targetRules = sheetRuleRegistry.cropRules(after.conditionalFormats ?? [], targetRange);
    if (params.transfer === 'move' && source.sheetId === params.sheetId) {
      targetRules = sheetRuleRegistry.cropRules(targetRules, source);
    }
    const sourceRules = sheetRuleRegistry.cloneRulesForPaste(metadata.conditionalFormats, {
      source,
      target: targetRange,
      transpose: params.spec.transpose,
      id: (rule) => `${rule.id}@paste:${targetRange.startRow}:${targetRange.startColumn}`,
    });
    after.conditionalFormats = [...targetRules, ...sourceRules];
  }
  if (params.spec.metadata.columnWidths) {
    after.columnWidths = metadata.columnWidths.map((entry) => ({
      column: targetRange.startColumn + entry.offset,
      widthPx: entry.widthPx,
    }));
    if (params.transfer === 'move') {
      after.columnWidths.push(...metadata.columnWidths.map((entry) => ({ column: source.startColumn + entry.offset, widthPx: undefined })));
    }
  }
  if (params.transfer === 'move' && source.sheetId === params.sheetId) {
    if (after.notes) for (const entry of metadata.notes) after.notes.push({ key: keyFor(source.startRow + entry.rowOffset, source.startColumn + entry.columnOffset) });
    if (after.hyperlinks) for (const entry of metadata.hyperlinks) after.hyperlinks.push({ key: keyFor(source.startRow + entry.rowOffset, source.startColumn + entry.columnOffset) });
    if (after.comments) after.comments = after.comments.filter((entry) => !rangeContains(source, entry.row, entry.column));
  }
  // Keep the source read in the planner so malformed cross-sheet references
  // fail before the mutation is registered.
  if (!sourceSheet || !targetSheet) throw new Error('Clipboard source or paste target sheet is unavailable');
}

export function registerEditingCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<PasteMutationParams>({
    id: 'range.paste',
    handler: (item, context) => {
    if (!isPasteMutation(item.params)) throw new Error('Invalid range.paste mutation payload');
    const params = item.params;
    const targetSheet = context.workbook.getSheet(params.sheetId);
    applyPasteSnapshot(targetSheet, params.snapshot);
    if (params.sourceRange && params.sourceRange.sheetId !== params.sheetId) {
      const sourceSheet = context.workbook.getSheet(params.sourceRange.sheetId);
      if (params.sourceSnapshot) applyPasteSnapshot(sourceSheet, params.sourceSnapshot);
    }
    },
    metadata: {
      schema: { name: 'PasteMutation', validate: isPasteMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: pasteAffectedRanges, mode: 'declared' },
      inverseIds: ['range.paste'],
    },
  });

  runtime.registry.registerCommand<PasteRangeParams>({
    id: 'sheet.range.paste',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (!isPasteSpecialSpec(params.spec)) throw new Error('Paste Special requires a canonical specification');
      const clipboard = parseClipboardPayload(params.clipboard, params.inputContext);
      const sourceRange = params.clipboard.range;
      const transfer = params.transfer;
      if (transfer !== 'copy' && transfer !== 'move' || params.clipboard.transfer !== transfer) {
        throw new Error('Paste transfer must match the canonical clipboard transfer');
      }
      if (transfer === 'move' && (!sourceRange || sourceRange.sheetId.length === 0)) {
        throw new Error('Move clipboard payload must include a source range');
      }
      const canonicalParams = { ...params, clipboard };
      const targetRange = assertPastePreconditions(context.workbook, canonicalParams);
      const sourceRow = sourceRange?.startRow ?? 0;
      const sourceColumn = sourceRange?.startColumn ?? 0;
      const sourceSheetName = context.workbook.getSheet(sourceRange.sheetId).name.replaceAll("'", "''");
      const rowCount = clipboard.sourceExtent.rows;
      const columnCount = clipboard.sourceExtent.columns;
      const targetRows = params.spec.transpose ? columnCount : rowCount;
      const targetColumns = params.spec.transpose ? rowCount : columnCount;
      const targetCellRange: RangeRef = {
        sheetId: params.sheetId,
        startRow: params.targetOrigin.row,
        endRow: params.targetOrigin.row + Math.max(0, targetRows - 1),
        startColumn: params.targetOrigin.column,
        endColumn: params.targetOrigin.column + Math.max(0, targetColumns - 1),
      };
      const affectedRanges: RangeRef[] = [structuredClone(targetCellRange)];
      if (transfer === 'move' && sourceRange) affectedRanges.push(structuredClone(sourceRange));
      const touchedRanges = transfer === 'move' && sourceRange && sourceRange.sheetId === params.sheetId ? [targetCellRange, sourceRange] : [targetCellRange];
      const clearsCells = params.spec.content !== 'none' && !params.spec.skipBlanks ? [structuredClone(targetCellRange)] : [];
      if (transfer === 'move' && sourceRange) clearsCells.push(structuredClone(sourceRange));
      const clearsMetadata = Object.values(params.spec.metadata).some(Boolean)
        ? [structuredClone(targetCellRange), ...(transfer === 'move' && sourceRange ? [structuredClone(sourceRange)] : [])]
        : [];
      const sparseWidths = (targetSheet: WorksheetModel, ranges: RangeRef[]) => Object.entries(targetSheet.columnWidthsPx)
        .map(([column, widthPx]) => ({ column: Number(column), widthPx }))
        .filter((entry) => ranges.some((range) => entry.column >= range.startColumn && entry.column <= range.endColumn));
      const before: PasteSnapshot = {
        clearRanges: clearsCells.filter((range) => range.sheetId === params.sheetId),
        clearMetadataRanges: clearsMetadata.filter((range) => range.sheetId === params.sheetId),
        cells: snapshotCells(sheet, touchedRanges),
        ...snapshotMetadata(sheet, touchedRanges, params.spec.metadata),
        ...(params.spec.metadata.validation ? { validations: structuredClone(sheet.dataValidations) } : {}),
        ...(params.spec.metadata.conditionalFormats ? { conditionalFormats: structuredClone(sheet.conditionalFormats) } : {}),
        ...(params.spec.metadata.columnWidths ? { columnWidths: sparseWidths(sheet, touchedRanges) } : {}),
      };
      const sourceSheet = transfer === 'move' ? context.workbook.getSheet(sourceRange!.sheetId) : undefined;
      const sourceBefore = transfer === 'move' && sourceRange && sourceRange.sheetId !== params.sheetId && sourceSheet
        ? {
          clearRanges: [structuredClone(sourceRange)],
          clearMetadataRanges: clearsMetadata.filter((range) => range.sheetId === sourceRange.sheetId),
          cells: snapshotCells(sourceSheet, [sourceRange]),
          ...snapshotMetadata(sourceSheet, [sourceRange], params.spec.metadata),
          ...(params.spec.metadata.validation ? { validations: structuredClone(sourceSheet.dataValidations) } : {}),
          ...(params.spec.metadata.conditionalFormats ? { conditionalFormats: structuredClone(sourceSheet.conditionalFormats) } : {}),
          ...(params.spec.metadata.columnWidths ? { columnWidths: sparseWidths(sourceSheet, [sourceRange]) } : {}),
        }
        : undefined;
      const after: PasteSnapshot = {
        clearRanges: structuredClone(before.clearRanges ?? []),
        clearMetadataRanges: structuredClone(before.clearMetadataRanges ?? []),
        cells: structuredClone(before.cells),
        ...(before.notes ? { notes: structuredClone(before.notes) } : {}),
        ...(before.hyperlinks ? { hyperlinks: structuredClone(before.hyperlinks) } : {}),
        ...(before.commentCells ? { commentCells: structuredClone(before.commentCells) } : {}),
        ...(before.comments ? { comments: structuredClone(before.comments) } : {}),
        ...(before.validations ? { validations: structuredClone(before.validations) } : {}),
        ...(before.conditionalFormats ? { conditionalFormats: structuredClone(before.conditionalFormats) } : {}),
        ...(before.columnWidths ? { columnWidths: structuredClone(before.columnWidths) } : {}),
      };
      const inRanges = (row: number, column: number, ranges: RangeRef[]) => ranges.some((range) => rangeContains(range, row, column));
      after.cells = after.cells.filter((entry) => !inRanges(entry.row, entry.column, after.clearRanges ?? []));
      if (after.notes) after.notes = after.notes.filter((entry) => {
        const [row, column] = entry.key.split(':').map(Number);
        return !inRanges(row, column, after.clearMetadataRanges ?? []);
      });
      if (after.hyperlinks) after.hyperlinks = after.hyperlinks.filter((entry) => {
        const [row, column] = entry.key.split(':').map(Number);
        return !inRanges(row, column, after.clearMetadataRanges ?? []);
      });
      if (after.commentCells) after.commentCells = after.commentCells.filter((key) => {
        const [row, column] = key.split(':').map(Number);
        return !inRanges(row, column, after.clearMetadataRanges ?? []);
      });
      if (after.comments) after.comments = after.comments.filter((entry) => !inRanges(entry.row, entry.column, after.clearMetadataRanges ?? []));
      const sourceAfter = sourceBefore ? {
        clearRanges: structuredClone(sourceBefore.clearRanges ?? []),
        clearMetadataRanges: structuredClone(sourceBefore.clearMetadataRanges ?? []),
        cells: [],
        ...(sourceBefore.notes ? { notes: [] } : {}),
        ...(sourceBefore.hyperlinks ? { hyperlinks: [] } : {}),
        ...(sourceBefore.commentCells ? { commentCells: [], comments: [] } : {}),
        ...(sourceBefore.validations ? { validations: sheetRuleRegistry.cropRules(sourceBefore.validations, sourceRange!) } : {}),
        ...(sourceBefore.conditionalFormats ? { conditionalFormats: sheetRuleRegistry.cropRules(sourceBefore.conditionalFormats, sourceRange!) } : {}),
        ...(sourceBefore.columnWidths ? { columnWidths: sourceBefore.columnWidths.map((entry) => ({ column: entry.column, widthPx: undefined })) } : {}),
      } : undefined;
      const afterCells = new Map(after.cells.map((entry) => [keyFor(entry.row, entry.column), entry]));
      const setAfterCell = (row: number, column: number, value: CellData | undefined) => {
        afterCells.set(keyFor(row, column), { row, column, ...(value ? { value: structuredClone(value) } : {}) });
      };
      for (const occupied of clipboard.occupiedCells) {
          const source = structuredClone(occupied.value);
          const rowOffset = occupied.rowOffset;
          const columnOffset = occupied.columnOffset;
          const row = params.targetOrigin.row + (params.spec.transpose ? columnOffset : rowOffset);
          const column = params.targetOrigin.column + (params.spec.transpose ? rowOffset : columnOffset);
          const sourceAddress = `'${sourceSheetName}'!${columnLabel(sourceColumn + columnOffset)}${sourceRow + rowOffset + 1}`;
          let next = applyPasteCell(
            params.spec,
            transfer,
            source,
            sheet.cells.get(row, column),
            row - sourceRow,
            column - sourceColumn,
            sourceAddress,
          );
          if (params.spec.metadata.hyperlinks && (source.hyperlink || source.hyperlinkDetail)) {
            next ??= structuredClone(sheet.cells.get(row, column) ?? { value: null });
            if (source.hyperlink) next.hyperlink = source.hyperlink;
            if (source.hyperlinkDetail) next.hyperlinkDetail = structuredClone(source.hyperlinkDetail);
          }
          if (next !== undefined) setAfterCell(row, column, next);
      }
      after.cells = [...afterCells.values()];
      applyPasteMetadataPlan(context.workbook, canonicalParams, targetRange, after);
      context.applyMutation({
        id: 'range.paste',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: {
          ...canonicalParams,
          sourceExtent: { rows: rowCount, columns: columnCount },
          transfer,
          sourceRange: transfer === 'move' ? structuredClone(sourceRange) : undefined,
          clearSource: transfer === 'move',
          snapshot: after,
          ...(sourceAfter ? { sourceSnapshot: sourceAfter } : {}),
        },
        affectedRanges,
        inverse: [{
          id: 'range.paste',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { ...canonicalParams, sourceExtent: { rows: rowCount, columns: columnCount }, sourceRange: transfer === 'move' ? structuredClone(sourceRange) : undefined, clearSource: transfer === 'move', snapshot: before, ...(sourceBefore ? { sourceSnapshot: sourceBefore } : {}) },
          affectedRanges,
        }],
        apply: () => {
          applyPasteSnapshot(sheet, after);
          if (transfer === 'move' && sourceRange && sourceRange.sheetId !== params.sheetId) {
            const sourceSheet = context.workbook.getSheet(sourceRange.sheetId);
            if (sourceAfter) applyPasteSnapshot(sourceSheet, sourceAfter);
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<MultiRangeStyleParams>({
    id: 'sheet.style.setMulti',
    execute: (params, context) => {
      const ranges = normalizeRanges(params.ranges);
      if (ranges.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      let lastResult = { operationId: context.operationId, mutationCount: 0, affectedRanges: [] as RangeRef[] };
      for (const range of ranges) {
        lastResult = runtime.execute('sheet.style.set', { sheetId: params.sheetId, range, style: params.style });
      }
      return lastResult;
    },
  });

  runtime.registry.registerCommand<FormatCellsParams>({
    id: 'sheet.format.set',
    execute: (params, context) => {
      const style: Partial<CellStyle> = { ...params.style };
      if (params.numberFormat !== undefined) style.numberFormat = params.numberFormat;
      const results = [];
      if (Object.keys(style).length > 0) {
        results.push(runtime.execute('sheet.style.setMulti', {
          sheetId: params.sheetId,
          ranges: params.ranges,
          style,
        }));
      }
      if (params.border) {
        results.push(runtime.execute('sheet.borders.set', {
          sheetId: params.sheetId,
          ranges: params.ranges,
          placement: params.border.placement,
          line: params.border.line,
        }));
      }
      return {
        operationId: context.operationId,
        mutationCount: results.reduce((count, result) => count + result.mutationCount, 0),
        affectedRanges: params.ranges,
      };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; ranges: RangeRef[]; numberFormat: string }>({
    id: 'sheet.numberFormat.apply',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges = normalizeRanges(params.ranges);
      for (const range of affectedRanges) {
        forEachCell(sheet, range, (row, column, cell) => {
          previous.push({ row, column, value: cell ? structuredClone(cell) : undefined });
        });
      }
      context.applyMutation({
        id: 'style.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row: item.row, column: item.column, previous: item.value },
          affectedRanges: [{ sheetId: params.sheetId, startRow: item.row, endRow: item.row, startColumn: item.column, endColumn: item.column }],
        })),
        apply: () => {
          for (const range of affectedRanges) {
            forEachCell(sheet, range, (row, column, cell) => {
              let next = cell ? { ...cell } : { value: null as CellData['value'] };
              next.numberFormat = params.numberFormat;
              if (next.style) next.style = { ...next.style, numberFormat: params.numberFormat };
              else next.style = { numberFormat: params.numberFormat };
              next.displayValue = formatValue(next.value, params.numberFormat);
              sheet.cells.set(row, column, next);
            });
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; preset: FreezePreset }>({
    id: 'sheet.freeze.preset',
    execute: (params, context) => {
      const pane: WorksheetPane = params.preset === 'none'
        ? { kind: 'none' }
        : { kind: 'frozen', xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0, state: 'frozen' };
      if (params.preset === 'firstRow' || params.preset === 'both') {
        if (pane.kind === 'frozen') { pane.ySplit = 1; pane.startRow = 1; }
      }
      if (params.preset === 'firstColumn' || params.preset === 'both') {
        if (pane.kind === 'frozen') { pane.xSplit = 1; pane.startColumn = 1; }
      }
      return runtime.execute('sheet.freeze.set', { sheetId: params.sheetId, pane });
    },
  });

  runtime.registry.registerMutation<SheetViewParams>({
    id: 'view.set',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid view.set mutation payload');
      const params = item.params as SheetViewParams;
      const sheet = context.workbook.getSheet(params.sheetId);
      if (params.showGridlines !== undefined) sheet.showGridlines = params.showGridlines;
      if (params.showHeaders !== undefined) sheet.showHeaders = params.showHeaders;
      if (params.zoom !== undefined) sheet.zoom = params.zoom;
    },
    metadata: {
      schema: { name: 'SheetView', validate: isSheetViewMutation },
      permission: { capability: 'sheet.view.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['view.set'],
    },
  });

  runtime.registry.registerCommand<SheetViewParams>({
    id: 'sheet.view.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = { showGridlines: sheet.showGridlines, showHeaders: sheet.showHeaders, zoom: sheet.zoom };
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'view.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{
          id: 'view.set',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, ...previous },
          affectedRanges,
        }],
        apply: () => {
          if (params.showGridlines !== undefined) sheet.showGridlines = params.showGridlines;
          if (params.showHeaders !== undefined) sheet.showHeaders = params.showHeaders;
          if (params.zoom !== undefined) sheet.zoom = Math.max(25, Math.min(400, params.zoom));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  const validateCellShiftEnvelope = (params: CellShiftParams, context: { workbook: WorkbookModel }): void => {
    const plan = planCellShift(context.workbook, params);
    if (JSON.stringify(plan.band) !== JSON.stringify(params.affectedBand)) throw new Error('Cell shift affected band is not canonical');
  };
  const cellShiftMutationHandler = (operation: CellShiftParams['operation'], id: 'cells.inserted' | 'cells.deleted') => (item: { params: unknown }, context: { workbook: WorkbookModel }) => {
      if (!isCellShiftMutation(item.params) || item.params.operation !== operation) throw new Error(`Invalid ${id} mutation payload`);
      validateCellShiftEnvelope(item.params, context);
      StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: item.params.sheetId, sourceRange: item.params.range, operation: item.params.operation, axis: item.params.axis });
    };
  runtime.registry.registerMutation<CellShiftParams>({ id: 'cells.inserted', handler: cellShiftMutationHandler('insert', 'cells.inserted'), metadata: { schema: { name: 'CellShiftInsert', validate: (value: unknown): value is CellShiftParams => isCellShiftMutation(value) && value.operation === 'insert' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.affectedBand)], mode: 'exact' }, inverseIds: ['cells.inserted.restore'] } });
  runtime.registry.registerMutation<CellShiftParams>({ id: 'cells.deleted', handler: cellShiftMutationHandler('delete', 'cells.deleted'), metadata: { schema: { name: 'CellShiftDelete', validate: (value: unknown): value is CellShiftParams => isCellShiftMutation(value) && value.operation === 'delete' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.affectedBand)], mode: 'exact' }, inverseIds: ['cells.deleted.restore'] } });
  const cellShiftRestoreMutationHandler = (operation: CellShiftParams['operation'], id: 'cells.inserted.restore' | 'cells.deleted.restore') => (item: { params: unknown }, context: { workbook: WorkbookModel }) => {
      if (!isCellShiftRestoreMutation(item.params) || item.params.spec.operation !== operation) throw new Error(`Invalid ${id} mutation payload`);
      validateCellShiftEnvelope(item.params.spec, context);
      const plan = planCellShift(context.workbook, item.params.spec);
      const sheet = context.workbook.getSheet(item.params.spec.sheetId);
      StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: item.params.spec.sheetId, sourceRange: item.params.spec.range, operation: operation === 'insert' ? 'delete' : 'insert', axis: item.params.spec.axis });
      for (let row = plan.band.startRow; row <= plan.band.endRow; row += 1) for (let column = plan.band.startColumn; column <= plan.band.endColumn; column += 1) sheet.cells.delete(row, column);
      for (const entry of item.params.cells) sheet.cells.set(entry.row, entry.column, structuredClone(entry.cell));
    };
  runtime.registry.registerMutation<CellShiftRestoreParams>({ id: 'cells.inserted.restore', handler: cellShiftRestoreMutationHandler('insert', 'cells.inserted.restore'), metadata: { schema: { name: 'CellShiftInsertRestore', validate: (value: unknown): value is CellShiftRestoreParams => isCellShiftRestoreMutation(value) && value.spec.operation === 'insert' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.spec.affectedBand)], mode: 'exact' }, inverseIds: ['cells.inserted'] } });
  runtime.registry.registerMutation<CellShiftRestoreParams>({ id: 'cells.deleted.restore', handler: cellShiftRestoreMutationHandler('delete', 'cells.deleted.restore'), metadata: { schema: { name: 'CellShiftDeleteRestore', validate: (value: unknown): value is CellShiftRestoreParams => isCellShiftRestoreMutation(value) && value.spec.operation === 'delete' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.spec.affectedBand)], mode: 'exact' }, inverseIds: ['cells.deleted'] } });
  const createCellShiftParams = (params: Omit<CellShiftParams, 'affectedBand'>, context: { workbook: WorkbookModel }) => {
    const plan = planCellShift(context.workbook, params);
    const canonicalParams: CellShiftParams = { ...params, affectedBand: plan.band };
    const sheet = context.workbook.getSheet(params.sheetId);
    const snapshot: Array<{ row: number; column: number; cell: CellData }> = [];
    forEachCell(sheet, plan.band, (row, column, cell) => { if (cell) snapshot.push({ row, column, cell: structuredClone(cell) }); });
    const affectedRanges: RangeRef[] = [structuredClone(plan.band)];
    return { canonicalParams, snapshot, affectedRanges };
  };
  runtime.registry.registerCommand<Omit<CellShiftParams, 'affectedBand'>>({ id: 'sheet.cells.insert', execute: (params, context) => { const { canonicalParams, snapshot, affectedRanges } = createCellShiftParams(params, context); context.applyMutation({ id: 'cells.inserted', unitId: context.workbook.unitId, sheetId: params.sheetId, params: canonicalParams, affectedRanges, inverse: [{ id: 'cells.inserted.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { spec: canonicalParams, cells: snapshot }, affectedRanges }], apply: () => StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: params.sheetId, sourceRange: params.range, operation: 'insert', axis: params.axis }) }); return { operationId: context.operationId, mutationCount: 1, affectedRanges }; } });
  runtime.registry.registerCommand<Omit<CellShiftParams, 'affectedBand'>>({ id: 'sheet.cells.delete', execute: (params, context) => { const { canonicalParams, snapshot, affectedRanges } = createCellShiftParams(params, context); context.applyMutation({ id: 'cells.deleted', unitId: context.workbook.unitId, sheetId: params.sheetId, params: canonicalParams, affectedRanges, inverse: [{ id: 'cells.deleted.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { spec: canonicalParams, cells: snapshot }, affectedRanges }], apply: () => StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: params.sheetId, sourceRange: params.range, operation: 'delete', axis: params.axis }) }); return { operationId: context.operationId, mutationCount: 1, affectedRanges }; } });

  runtime.registry.registerMutation<{ sourceSheetId: string; newId: string; newName: string }>({
    id: 'sheet.duplicated',
    handler: (item, context) => {
      if (!isSheetDuplicateMutation(item.params)) throw new Error('Invalid sheet.duplicated mutation payload');
      const params = item.params;
      context.workbook.duplicateSheet(params.sourceSheetId, params.newId, params.newName);
    },
    metadata: {
      schema: { name: 'DuplicateSheet', validate: isSheetDuplicateMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.remove'],
    },
  });

  runtime.registry.registerCommand<{ sourceSheetId: string; newId: string; newName: string }>({
    id: 'sheet.duplicate',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.duplicated',
        unitId: context.workbook.unitId,
        sheetId: params.newId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.remove', unitId: context.workbook.unitId, sheetId: params.newId, params: { id: params.newId }, affectedRanges }],
        apply: () => context.workbook.duplicateSheet(params.sourceSheetId, params.newId, params.newName),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'sheet.hidden',
    handler: (item, context) => {
      if (!isSheetIdMutation(item.params)) throw new Error('Invalid sheet.hidden mutation payload');
      context.workbook.getSheet(item.params.sheetId).hidden = true;
    },
    metadata: {
      schema: { name: 'SheetHidden', validate: isSheetIdMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.unhidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'sheet.unhidden',
    handler: (item, context) => {
      if (!isSheetIdMutation(item.params)) throw new Error('Invalid sheet.unhidden mutation payload');
      context.workbook.getSheet(item.params.sheetId).hidden = false;
    },
    metadata: {
      schema: { name: 'SheetUnhidden', validate: isSheetIdMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.hidden'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.hide',
    execute: (params, context) => {
      const visible = context.workbook.getVisibleSheets();
      if (visible.length <= 1 && !context.workbook.getSheet(params.sheetId).hidden) {
        throw new Error('Cannot hide the only visible worksheet');
      }
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => { context.workbook.getSheet(params.sheetId).hidden = true; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.unhide',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.unhidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.hidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => { context.workbook.getSheet(params.sheetId).hidden = false; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string; toIndex: number }>({
    id: 'sheet.reordered',
    handler: (item, context) => {
      if (!isSheetReorderedMutation(item.params)) throw new Error('Invalid sheet.reordered mutation payload');
      const params = item.params;
      context.workbook.reorderSheet(params.sheetId, params.toIndex);
    },
    metadata: {
      schema: { name: 'ReorderSheet', validate: isSheetReorderedMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.reordered'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; toIndex: number }>({
    id: 'sheet.reorder',
    execute: (params, context) => {
      const previous = [...context.workbook.sheetOrder];
      const fromIndex = previous.indexOf(params.sheetId);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.reordered',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.reordered', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, toIndex: fromIndex }, affectedRanges }],
        apply: () => context.workbook.reorderSheet(params.sheetId, params.toIndex),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string; color?: string }>({
    id: 'sheet.tabColor',
    handler: (item, context) => {
      if (!isTabColorMutation(item.params)) throw new Error('Invalid sheet.tabColor mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).tabColor = params.color;
    },
    metadata: {
      schema: { name: 'TabColor', validate: isTabColorMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.tabColor'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; color?: string }>({
    id: 'sheet.tabColor.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.tabColor;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.tabColor',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.tabColor', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, color: previous }, affectedRanges }],
        apply: () => { sheet.tabColor = params.color; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<GoToParams>({
    id: 'navigation.goto',
    execute: (params, context) => {
      const target = resolveGoToRange(context.workbook, params);
      if (!target) throw new Error(`Invalid reference: ${params.reference}`);
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [target],
      };
    },
  });

  runtime.registry.registerCommand<GoToSpecialParams>({
    id: 'navigation.gotoSpecial',
    execute: (params, context) => {
      const ranges = resolveGoToSpecial(context.workbook, params);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: ranges };
    },
  });

}

export function buildClipboardFromRange(workbook: WorkbookModel, range: RangeRef): ClipboardPayload {
  return copyRangeToClipboardData(workbook, range);
}

export function restoreCellFromMutation(
  workbook: WorkbookModel,
  item: MutationInfo<{ row: number; column: number; previous?: CellData }>,
): void {
  const sheet = workbook.getSheet(item.sheetId);
  const { row, column, previous } = item.params;
  if (previous) sheet.cells.set(row, column, previous);
  else sheet.cells.delete(row, column);
}
