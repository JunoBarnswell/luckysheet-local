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
  WorkbookTheme,
} from '@react-sheets/core-model';
import { cellKey, clearFormulaProvenance, columnLabel, MAX_SHEET_COLUMN_COUNT, MAX_SHEET_ROW_COUNT, planCellShift, sheetRuleRegistry, type CellShiftSpec } from '@react-sheets/core-model';
import { StructuralTransform } from '@react-sheets/core-model';
import { parseFormula } from '@react-sheets/formula-engine';
import { formatValue } from '@react-sheets/number-format';
import type { CommandContext, CommandResult, CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import {
  DEFAULT_PASTE_SPECIAL_SPEC,
  copyRangeToClipboardData,
  parseClipboardPayload,
  shiftFormula,
  type ClipboardTransfer,
  type ClipboardPayload,
  type PasteSpecialSpec,
  isPasteSpecialSpecSupported,
} from '../clipboard';
import { isCellInputInterpretationContext, type CellInputInterpretationContext } from '../text-input';

export type FreezePreset = 'none' | 'firstRow' | 'firstColumn' | 'both';
export type GoToSpecialKind =
  | 'blanks'
  | 'constants'
  | 'constant-numbers'
  | 'constant-text'
  | 'constant-logical'
  | 'constant-errors'
  | 'formulas'
  | 'formula-numbers'
  | 'formula-text'
  | 'formula-logical'
  | 'formula-errors'
  | 'comments'
  | 'notes'
  | 'comments-notes'
  | 'visible'
  | 'errors'
  | 'conditional-format'
  | 'conditional-format-all'
  | 'conditional-format-same'
  | 'data-validation'
  | 'data-validation-all'
  | 'data-validation-same'
  | 'current-region'
  | 'current-array'
  | 'row-differences'
  | 'column-differences'
  | 'precedents'
  | 'dependents'
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
  /** Required only when the payload still contains host text/HTML representations. */
  inputContext?: CellInputInterpretationContext;
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

type PasteMutationParams = Omit<PasteRangeParams, 'inputContext'> & {
  sourceExtent: { rows: number; columns: number };
  sourceRange?: RangeRef;
  clearSource?: boolean;
  snapshot: PasteSnapshot;
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
  workbookTheme?: WorkbookTheme;
}

function isPasteMutation(value: unknown): value is PasteMutationParams {
  if (!isRecord(value) || !isRecord(value.clipboard) || !isRecord(value.clipboard.sourceExtent)) return false;
  const clipboard = value.clipboard;
  const sourceExtent = clipboard.sourceExtent as Record<string, unknown>;
  const declaredExtent = value.sourceExtent;
  const clipboardRange = clipboard.range;
  if (!isRecord(declaredExtent) || !isRange(clipboardRange)) return false;
  const clipboardRangeWithinWorksheet = clipboardRange.sheetId.trim().length > 0
    && clipboardRange.startRow < MAX_SHEET_ROW_COUNT && clipboardRange.endRow < MAX_SHEET_ROW_COUNT
    && clipboardRange.startColumn < MAX_SHEET_COLUMN_COUNT && clipboardRange.endColumn < MAX_SHEET_COLUMN_COUNT;
  const extentMatchesClipboard = Number(sourceExtent.rows) === Number(declaredExtent.rows)
    && Number(sourceExtent.columns) === Number(declaredExtent.columns)
    && clipboardRange.endRow - clipboardRange.startRow + 1 === Number(sourceExtent.rows)
    && clipboardRange.endColumn - clipboardRange.startColumn + 1 === Number(sourceExtent.columns);
  const sourceMatchesClipboard = isRange(value.sourceRange)
    && value.sourceRange.sheetId === clipboardRange.sheetId
    && value.sourceRange.startRow === clipboardRange.startRow
    && value.sourceRange.endRow === clipboardRange.endRow
    && value.sourceRange.startColumn === clipboardRange.startColumn
    && value.sourceRange.endColumn === clipboardRange.endColumn;
  const valid = typeof value.sheetId === 'string'
    && isRecord(value.targetOrigin) && Number.isSafeInteger(value.targetOrigin.row) && Number(value.targetOrigin.row) >= 0
    && Number.isSafeInteger(value.targetOrigin.column) && Number(value.targetOrigin.column) >= 0
    && Number(value.targetOrigin.row) < MAX_SHEET_ROW_COUNT && Number(value.targetOrigin.column) < MAX_SHEET_COLUMN_COUNT
    && (value.transfer === 'copy' || value.transfer === 'move')
    && clipboard.transfer === value.transfer
    && clipboard.schema === 'SparseClipboardPayload'
    && clipboardRangeWithinWorksheet
    && Number.isInteger(sourceExtent.rows) && Number(sourceExtent.rows) > 0
    && Number.isInteger(sourceExtent.columns) && Number(sourceExtent.columns) > 0
    && Array.isArray(clipboard.occupiedCells)
    && clipboard.occupiedCells.every((cell) => isRecord(cell)
      && Number.isInteger(cell.rowOffset) && Number(cell.rowOffset) >= 0 && Number(cell.rowOffset) < Number(sourceExtent.rows)
      && Number.isInteger(cell.columnOffset) && Number(cell.columnOffset) >= 0 && Number(cell.columnOffset) < Number(sourceExtent.columns)
      && isCellData(cell.value))
    && isRecord(clipboard.rangeMetadata)
    && Array.isArray(clipboard.rangeMetadata.columnWidths)
    && Array.isArray(clipboard.rangeMetadata.validations)
    && Array.isArray(clipboard.rangeMetadata.conditionalFormats)
    && Array.isArray(clipboard.rangeMetadata.notes)
    && Array.isArray(clipboard.rangeMetadata.comments)
    && Array.isArray(clipboard.rangeMetadata.hyperlinks)
    && extentMatchesClipboard
    && isPasteSpecialSpec(value.spec)
    && isPasteSpecialSpecSupported(value.spec, clipboard as unknown as ClipboardPayload)
    && isRecord(value.sourceExtent) && Number.isInteger(value.sourceExtent.rows) && Number.isInteger(value.sourceExtent.columns)
    && isPasteSnapshot(value.snapshot)
    && (value.transfer === 'move'
      ? sourceMatchesClipboard && value.clearSource === true
      : value.sourceRange === undefined && value.clearSource === false);
  if (!valid) return false;
  const pasteParams = value as unknown as PasteMutationParams;
  const cellRanges = pasteCellRanges(pasteParams);
  const mappedWidthColumns = pasteTargetWidthColumns(pasteParams);
  if (!mappedWidthColumns) return false;
  return cellRanges.every((range) => range.startRow >= 0 && range.endRow < MAX_SHEET_ROW_COUNT
      && range.startColumn >= 0 && range.endColumn < MAX_SHEET_COLUMN_COUNT)
    && isPasteSnapshotWithinRanges(pasteParams.snapshot, cellRanges, mappedWidthColumns)
    && isPasteSnapshotConsistentWithSpec(pasteParams)
    && isPasteRuleCollectionForSheet(pasteParams.snapshot.validations, pasteParams.sheetId, 'validation')
    && isPasteRuleCollectionForSheet(pasteParams.snapshot.conditionalFormats, pasteParams.sheetId, 'conditional-format');
}

function pasteCellRanges(value: PasteMutationParams): RangeRef[] {
  const rowCount = value.spec.transpose ? value.sourceExtent.columns : value.sourceExtent.rows;
  const columnCount = value.spec.transpose ? value.sourceExtent.rows : value.sourceExtent.columns;
  const ranges = [{ sheetId: value.sheetId, startRow: value.targetOrigin.row, endRow: value.targetOrigin.row + Math.max(0, rowCount - 1), startColumn: value.targetOrigin.column, endColumn: value.targetOrigin.column + Math.max(0, columnCount - 1) }];
  if (value.clearSource && value.sourceRange) ranges.push(structuredClone(value.sourceRange));
  return ranges;
}

function pasteTargetWidthColumns(value: PasteMutationParams): Set<number> | undefined {
  if (!value.spec.metadata.columnWidths) return new Set();
  const columns = new Set<number>();
  const offsets = new Set<number>();
  for (const entry of value.clipboard.rangeMetadata.columnWidths) {
    if (!isRecord(entry) || !Number.isSafeInteger(entry.offset) || Number(entry.offset) < 0
      || Number(entry.offset) >= value.sourceExtent.columns
      || typeof entry.widthPx !== 'number' || !Number.isFinite(entry.widthPx) || entry.widthPx <= 0
      || offsets.has(Number(entry.offset))) return undefined;
    offsets.add(Number(entry.offset));
    const column = value.targetOrigin.column + Number(entry.offset);
    if (!Number.isSafeInteger(column) || column < 0 || column >= MAX_SHEET_COLUMN_COUNT) return undefined;
    columns.add(column);
  }
  return columns;
}

function isPasteSnapshotConsistentWithSpec(value: PasteMutationParams): boolean {
  const [targetRange] = pasteCellRanges(value);
  if (!targetRange || (value.transfer === 'move' && value.sourceRange?.sheetId !== value.sheetId)) return false;
  if (value.transfer === 'move' && value.sourceRange && rangesIntersect(targetRange, value.sourceRange)) return false;
  const expectedClearRanges = value.spec.content !== 'none' && !value.spec.skipBlanks ? [targetRange] : [];
  if (value.transfer === 'move' && value.sourceRange) expectedClearRanges.push(value.sourceRange);
  const hasMetadata = Object.values(value.spec.metadata).some(Boolean);
  const expectedMetadataRanges = hasMetadata
    ? [targetRange, ...(value.transfer === 'move' && value.sourceRange ? [value.sourceRange] : [])]
    : [];
  const sameRanges = (actual: readonly RangeRef[] | undefined, expected: readonly RangeRef[]) => Array.isArray(actual)
    && actual.length === expected.length
    && actual.every((range, index) => {
      const other = expected[index];
      return other !== undefined && range.sheetId === other.sheetId
        && range.startRow === other.startRow && range.endRow === other.endRow
        && range.startColumn === other.startColumn && range.endColumn === other.endColumn;
    });
  const snapshot = value.snapshot;
  return sameRanges(snapshot.clearRanges, expectedClearRanges)
    && sameRanges(snapshot.clearMetadataRanges, expectedMetadataRanges)
    && (snapshot.notes !== undefined) === value.spec.metadata.commentsNotes
    && (snapshot.comments !== undefined) === value.spec.metadata.commentsNotes
    && (snapshot.commentCells !== undefined) === value.spec.metadata.commentsNotes
    && (snapshot.hyperlinks !== undefined) === value.spec.metadata.hyperlinks
    && (snapshot.validations !== undefined) === value.spec.metadata.validation
    && (snapshot.conditionalFormats !== undefined) === value.spec.metadata.conditionalFormats
    && (snapshot.columnWidths !== undefined) === value.spec.metadata.columnWidths
    && (snapshot.workbookTheme !== undefined) === (value.spec.formatting === 'source-theme');
}

function pasteAffectedRanges(value: PasteMutationParams): RangeRef[] {
  const cellRanges = pasteCellRanges(value);
  if (value.snapshot.validations !== undefined
    || value.snapshot.conditionalFormats !== undefined
    || value.snapshot.workbookTheme !== undefined) return [{
    sheetId: value.sheetId,
    startRow: 0,
    endRow: MAX_SHEET_ROW_COUNT - 1,
    startColumn: 0,
    endColumn: MAX_SHEET_COLUMN_COUNT - 1,
  }];
  const ranges = [...cellRanges];
  for (const entry of value.snapshot.columnWidths ?? []) ranges.push({
    sheetId: value.sheetId,
    startRow: 0,
    endRow: MAX_SHEET_ROW_COUNT - 1,
    startColumn: entry.column,
    endColumn: entry.column,
  });
  return ranges;
}

function isPasteSnapshotWithinRanges(snapshot: unknown, ranges: readonly RangeRef[], mappedWidthColumns: ReadonlySet<number>): boolean {
  if (!isRecord(snapshot)) return false;
  const containsPoint = (row: unknown, column: unknown) => typeof row === 'number' && Number.isSafeInteger(row)
    && row >= 0 && row < MAX_SHEET_ROW_COUNT
    && typeof column === 'number' && Number.isSafeInteger(column)
    && column >= 0 && column < MAX_SHEET_COLUMN_COUNT
    && ranges.some((range) => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn);
  const containsRange = (candidate: unknown) => isRange(candidate)
    && ranges.some((range) => candidate.sheetId === range.sheetId
      && candidate.startRow >= range.startRow && candidate.endRow <= range.endRow
      && candidate.startColumn >= range.startColumn && candidate.endColumn <= range.endColumn);
  const keyWithinRanges = (key: unknown) => {
    if (typeof key !== 'string' || !/^\d+:\d+$/.test(key)) return false;
    const [row, column] = key.split(':').map(Number);
    return key === `${row}:${column}` && containsPoint(row, column);
  };
  const hasOnlyScopedPoints = (entries: unknown, point: (entry: unknown) => boolean) => entries === undefined
    || (Array.isArray(entries) && entries.every(point));

  return (snapshot.clearRanges === undefined || (Array.isArray(snapshot.clearRanges) && snapshot.clearRanges.every(containsRange)))
    && (snapshot.clearMetadataRanges === undefined || (Array.isArray(snapshot.clearMetadataRanges) && snapshot.clearMetadataRanges.every(containsRange)))
    && Array.isArray(snapshot.cells) && snapshot.cells.every((entry) => isRecord(entry) && containsPoint(entry.row, entry.column))
    && hasOnlyScopedPoints(snapshot.notes, (entry) => isRecord(entry) && keyWithinRanges(entry.key))
    && hasOnlyScopedPoints(snapshot.hyperlinks, (entry) => isRecord(entry) && keyWithinRanges(entry.key))
    && hasOnlyScopedPoints(snapshot.commentCells, keyWithinRanges)
    && hasOnlyScopedPoints(snapshot.comments, (entry) => isRecord(entry)
      && entry.sheetId === ranges[0]?.sheetId && containsPoint(entry.row, entry.column))
    && hasOnlyScopedPoints(snapshot.columnWidths, (entry) => {
      if (!isRecord(entry) || typeof entry.column !== 'number' || !Number.isSafeInteger(entry.column)) return false;
      const column = entry.column;
      return column >= 0 && column < MAX_SHEET_COLUMN_COUNT
        && (mappedWidthColumns.has(column) || ranges.some((range) => range.sheetId === ranges[0]?.sheetId
          && column >= range.startColumn && column <= range.endColumn));
    });
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

function isDefaultMoveSpec(spec: PasteSpecialSpec): boolean {
  return spec.content === DEFAULT_PASTE_SPECIAL_SPEC.content
    && spec.formatting === DEFAULT_PASTE_SPECIAL_SPEC.formatting
    && spec.metadata.commentsNotes === DEFAULT_PASTE_SPECIAL_SPEC.metadata.commentsNotes
    && spec.metadata.validation === DEFAULT_PASTE_SPECIAL_SPEC.metadata.validation
    && spec.metadata.columnWidths === DEFAULT_PASTE_SPECIAL_SPEC.metadata.columnWidths
    && spec.metadata.conditionalFormats === DEFAULT_PASTE_SPECIAL_SPEC.metadata.conditionalFormats
    && spec.metadata.hyperlinks === DEFAULT_PASTE_SPECIAL_SPEC.metadata.hyperlinks
    && spec.operation === DEFAULT_PASTE_SPECIAL_SPEC.operation
    && spec.skipBlanks === DEFAULT_PASTE_SPECIAL_SPEC.skipBlanks
    && spec.transpose === DEFAULT_PASTE_SPECIAL_SPEC.transpose
    && spec.link === DEFAULT_PASTE_SPECIAL_SPEC.link;
}

function isPasteSnapshot(value: unknown): value is PasteSnapshot {
  if (!isRecord(value) || !Array.isArray(value.cells)) return false;
  if (value.clearRanges !== undefined) {
    if (!Array.isArray(value.clearRanges) || !value.clearRanges.every(isRange)) return false;
  }
  if (value.clearMetadataRanges !== undefined) {
    if (!Array.isArray(value.clearMetadataRanges) || !value.clearMetadataRanges.every(isRange)) return false;
  }
  for (const entry of value.cells) {
    if (!isRecord(entry) || !Number.isInteger(entry.row) || !Number.isInteger(entry.column)) return false;
    if (entry.value !== undefined && !isCellData(entry.value)) return false;
  }
  if (value.notes !== undefined) {
    if (!Array.isArray(value.notes)) return false;
    for (const entry of value.notes) {
      if (!isRecord(entry) || typeof entry.key !== 'string') return false;
      if (entry.value !== undefined && !isCellNoteSnapshot(entry.value)) return false;
    }
  }
  if (value.hyperlinks !== undefined) {
    if (!Array.isArray(value.hyperlinks)) return false;
    for (const entry of value.hyperlinks) {
      if (!isRecord(entry) || typeof entry.key !== 'string') return false;
      if (entry.value !== undefined && !isCellHyperlinkSnapshot(entry.value)) return false;
    }
  }
  if (value.commentCells !== undefined) {
    if (!Array.isArray(value.commentCells) || !value.commentCells.every((key) => typeof key === 'string')) return false;
  }
  if (value.comments !== undefined) {
    if (!Array.isArray(value.comments) || !value.comments.every(isCommentThreadSnapshot)) return false;
  }
  if (value.validations !== undefined && !Array.isArray(value.validations)) return false;
  if (value.conditionalFormats !== undefined && !Array.isArray(value.conditionalFormats)) return false;
  if (value.columnWidths !== undefined) {
    if (!Array.isArray(value.columnWidths)) return false;
    for (const entry of value.columnWidths) {
      if (!isRecord(entry) || !Number.isSafeInteger(entry.column)
        || Number(entry.column) < 0 || Number(entry.column) >= MAX_SHEET_COLUMN_COUNT) return false;
      if (entry.widthPx !== undefined && (typeof entry.widthPx !== 'number' || !Number.isFinite(entry.widthPx) || entry.widthPx <= 0)) return false;
    }
  }
  return isPasteWorkbookTheme(value.workbookTheme);
}

function isPasteWorkbookTheme(value: unknown): value is WorkbookTheme | undefined {
  if (value === undefined) return true;
  if (!isRecord(value) || typeof value.id !== 'string' || value.id.trim().length === 0 || !isRecord(value.colors)) return false;
  return Object.values(value.colors).every((color) => typeof color === 'string' && /^#[0-9a-f]{6}$/i.test(color));
}

function isCellNoteSnapshot(value: unknown): value is CellNote {
  return isRecord(value) && typeof value.id === 'string' && value.id.trim().length > 0
    && typeof value.author === 'string' && typeof value.text === 'string'
    && typeof value.createdAt === 'string' && typeof value.visible === 'boolean';
}

function isCellHyperlinkSnapshot(value: unknown): value is CellHyperlink {
  if (!isRecord(value) || typeof value.id !== 'string' || !value.id.trim() || !isRecord(value.target)) return false;
  const target = value.target;
  if (value.tooltip !== undefined && typeof value.tooltip !== 'string') return false;
  switch (target.kind) {
    case 'url': return typeof target.url === 'string' && target.url.length > 0;
    case 'email': return typeof target.address === 'string' && target.address.length > 0
      && (target.subject === undefined || typeof target.subject === 'string');
    case 'sheet': return typeof target.sheetId === 'string' && target.sheetId.length > 0
      && (target.address === undefined || typeof target.address === 'string')
      && (target.row === undefined || (Number.isSafeInteger(target.row) && Number(target.row) >= 0))
      && (target.column === undefined || (Number.isSafeInteger(target.column) && Number(target.column) >= 0));
    case 'name': return typeof target.name === 'string' && target.name.length > 0;
    default: return false;
  }
}

function isCommentReplySnapshot(value: unknown): boolean {
  return isRecord(value) && typeof value.id === 'string' && value.id.length > 0
    && typeof value.author === 'string' && typeof value.text === 'string' && typeof value.createdAt === 'string'
    && (value.mentions === undefined || (Array.isArray(value.mentions) && value.mentions.every((mention) => typeof mention === 'string')));
}

function isCommentThreadSnapshot(value: unknown): value is CommentThread {
  return isRecord(value) && typeof value.id === 'string' && value.id.length > 0
    && typeof value.sheetId === 'string' && Number.isSafeInteger(value.row) && Number(value.row) >= 0
    && Number.isSafeInteger(value.column) && Number(value.column) >= 0
    && typeof value.author === 'string' && typeof value.text === 'string' && typeof value.createdAt === 'string'
    && Array.isArray(value.replies) && value.replies.every(isCommentReplySnapshot)
    && (value.mentions === undefined || (Array.isArray(value.mentions) && value.mentions.every((mention) => typeof mention === 'string')))
    && (value.resolved === undefined || typeof value.resolved === 'boolean')
    && (value.resolvedAt === undefined || typeof value.resolvedAt === 'string');
}

function isPasteRuleCollectionForSheet(
  value: unknown,
  sheetId: string,
  kind: 'validation' | 'conditional-format',
): boolean {
  if (value === undefined) return true;
  if (!Array.isArray(value)) return false;
  const ruleTypes = kind === 'validation'
    ? ['list', 'whole', 'decimal', 'date', 'time', 'checkbox', 'textLength', 'custom']
    : ['highlight', 'dataBar', 'colorScale', 'iconSet', 'topBottom'];
  const operators = kind === 'validation'
    ? ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan']
    : ['greaterThan', 'lessThan', 'between', 'equal', 'notEqual', 'containsText', 'notContainsText', 'duplicate', 'unique', 'formula', 'top', 'bottom'];
  return value.every((candidate) => {
    if (!isRecord(candidate) || candidate.sheetId !== sheetId || !Array.isArray(candidate.ranges)
      || !candidate.ranges.every((range) => isRange(range) && range.sheetId === sheetId
        && range.endRow < MAX_SHEET_ROW_COUNT && range.endColumn < MAX_SHEET_COLUMN_COUNT)) return false;
    if (typeof candidate.id !== 'string' || candidate.id.trim().length === 0
      || !ruleTypes.includes(String(candidate.type))
      || (candidate.operator !== undefined && !operators.includes(String(candidate.operator)))) return false;
    if (kind === 'validation') {
      if ((candidate.formula1 !== undefined && typeof candidate.formula1 !== 'string')
        || (candidate.formula2 !== undefined && typeof candidate.formula2 !== 'string')) return false;
      const listSource = candidate.listSource;
      if (listSource !== undefined) {
        if (!isRecord(listSource)) return false;
        if (listSource.kind === 'values' && (!Array.isArray(listSource.values) || !listSource.values.every((entry) => typeof entry === 'string'))) return false;
        if (listSource.kind === 'range') {
          const range = listSource.range;
          if (!isRange(range) || range.sheetId !== sheetId
            || range.endRow >= MAX_SHEET_ROW_COUNT || range.endColumn >= MAX_SHEET_COLUMN_COUNT) return false;
        }
        if (listSource.kind === 'formula' && typeof listSource.formula !== 'string') return false;
        if (listSource.kind !== 'values' && listSource.kind !== 'range' && listSource.kind !== 'formula') return false;
      }
    } else if ((candidate.value1 !== undefined && typeof candidate.value1 !== 'string'
      && !(typeof candidate.value1 === 'number' && Number.isFinite(candidate.value1)))
      || (candidate.value2 !== undefined && typeof candidate.value2 !== 'string'
        && !(typeof candidate.value2 === 'number' && Number.isFinite(candidate.value2)))) return false;
    if (kind === 'conditional-format' && candidate.topBottom !== undefined) {
      const topBottom = candidate.topBottom;
      if (!isRecord(topBottom)
        || (topBottom.direction !== 'top' && topBottom.direction !== 'bottom')
        || (topBottom.percent !== undefined && typeof topBottom.percent !== 'boolean')) return false;
    }
    try {
      const normalized = kind === 'validation'
        ? sheetRuleRegistry.normalizeDataValidation(candidate as unknown as DataValidationRule, (formula) => { parseFormula(formula); })
        : sheetRuleRegistry.normalizeConditionalFormat(candidate as unknown as ConditionalFormatRule, 1, (formula) => { parseFormula(formula); });
      const formulaAnchor = normalized.formulaAnchor;
      return normalized.ranges.length > 0 && normalized.ranges.every((range) => range.sheetId === sheetId
        && range.endRow < MAX_SHEET_ROW_COUNT && range.endColumn < MAX_SHEET_COLUMN_COUNT)
        && formulaAnchor !== undefined && formulaAnchor.sheetId === sheetId
        && formulaAnchor.row < MAX_SHEET_ROW_COUNT
        && formulaAnchor.column < MAX_SHEET_COLUMN_COUNT;
    } catch {
      return false;
    }
  });
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

function rangeContainsAddress(range: RangeRef, address: { row: number; column: number }): boolean {
  return range.startRow <= address.row && range.endRow >= address.row && range.startColumn <= address.column && range.endColumn >= address.column;
}

function parseA1RangeInSheet(value: string, sheetId: string): RangeRef | undefined {
  const [first, second = first] = value.split(':');
  const parseCell = (input: string | undefined): { row: number; column: number } | undefined => {
    const match = input?.trim().match(/^\$?([A-Z]+)\$?(\d+)$/i);
    if (!match) return undefined;
    let column = 0;
    for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
    const row = Number(match[2]) - 1;
    return Number.isSafeInteger(row) && row >= 0 ? { row, column: column - 1 } : undefined;
  };
  const start = parseCell(first);
  const end = parseCell(second);
  if (!start || !end) return undefined;
  return { sheetId, startRow: Math.min(start.row, end.row), endRow: Math.max(start.row, end.row), startColumn: Math.min(start.column, end.column), endColumn: Math.max(start.column, end.column) };
}

function isErrorCell(cell: CellData | undefined): boolean {
  const value = cell?.formulaValue;
  return Boolean((typeof cell?.value === 'string' && cell.value.startsWith('#')) || (value && typeof value === 'object' && 'kind' in value && (value as { kind?: string }).kind === 'error'));
}

function cellComparisonKey(cell: CellData | undefined): string {
  return JSON.stringify(cell?.formulaValue ?? cell?.value ?? null);
}

function resolveDimensionDifferences(sheet: WorksheetModel, range: RangeRef, axis: 'row' | 'column'): RangeRef[] {
  const hits: RangeRef[] = [];
  if (axis === 'row') {
    const baseline = Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, offset) => cellComparisonKey(sheet.cells.get(range.startRow, range.startColumn + offset)));
    for (let row = range.startRow + 1; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (cellComparisonKey(sheet.cells.get(row, column)) !== baseline[column - range.startColumn]) hits.push({ sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column });
    }
  } else {
    const baseline = Array.from({ length: range.endRow - range.startRow + 1 }, (_, offset) => cellComparisonKey(sheet.cells.get(range.startRow + offset, range.startColumn)));
    for (let column = range.startColumn + 1; column <= range.endColumn; column += 1) for (let row = range.startRow; row <= range.endRow; row += 1) {
      if (cellComparisonKey(sheet.cells.get(row, column)) !== baseline[row - range.startRow]) hits.push({ sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column });
    }
  }
  return hits;
}

interface FormulaReference {
  sheetId: string;
  row: number;
  column: number;
}

function formulaReferences(workbook: WorkbookModel, owner: WorksheetModel, formula: string): FormulaReference[] {
  const references: FormulaReference[] = [];
  const pattern = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?\$?([A-Z]{1,3})\$?(\d+)/g;
  for (const match of formula.matchAll(pattern)) {
    let column = 0;
    for (const character of match[3]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
    const row = Number(match[4]) - 1;
    if (!Number.isSafeInteger(row) || row < 0) continue;
    const referencedSheet = match[1] || match[2] ? workbook.getSheetByName((match[1] || match[2]!).trim()) : owner;
    if (!referencedSheet) continue;
    references.push({ sheetId: referencedSheet.id, row, column: column - 1 });
  }
  return references;
}

function resolveFormulaPrecedents(workbook: WorkbookModel, sheet: WorksheetModel, range: RangeRef): RangeRef[] {
  const hits: RangeRef[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const formula = sheet.cells.get(row, column)?.formula;
    if (!formula) continue;
    for (const reference of formulaReferences(workbook, sheet, formula)) hits.push({ sheetId: reference.sheetId, startRow: reference.row, endRow: reference.row, startColumn: reference.column, endColumn: reference.column });
  }
  return hits;
}

function resolveFormulaDependents(workbook: WorkbookModel, sheet: WorksheetModel, range: RangeRef): RangeRef[] {
  const targets = new Set<string>();
  for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) targets.add(`${sheet.id}:${row}:${column}`);
  const hits: RangeRef[] = [];
  for (const candidate of workbook.getSheets()) candidate.cells.forEach((cell, row, column) => {
    if (!cell.formula) return;
    if (formulaReferences(workbook, candidate, cell.formula).some((reference) => targets.has(`${reference.sheetId}:${reference.row}:${reference.column}`))) hits.push({ sheetId: candidate.id, startRow: row, endRow: row, startColumn: column, endColumn: column });
  });
  return hits;
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
  const anchor = { row: normalizedRange.startRow, column: normalizedRange.startColumn };
  if (params.kind === 'current-array') {
    const spill = sheet.spillRanges.find((candidate) => rangeContainsAddress(candidate.range, anchor));
    if (spill) return [structuredClone(spill.range)];
    const cell = sheet.cells.get(anchor.row, anchor.column);
    if (cell?.formulaMetadata?.kind === 'array' && cell.formulaMetadata.range) {
      const arrayRange = parseA1RangeInSheet(cell.formulaMetadata.range, params.sheetId);
      if (arrayRange) return [arrayRange];
    }
    return [];
  }
  if (params.kind === 'row-differences' || params.kind === 'column-differences') {
    return resolveDimensionDifferences(sheet, normalizedRange, params.kind === 'row-differences' ? 'row' : 'column');
  }
  if (params.kind === 'precedents') return resolveFormulaPrecedents(workbook, sheet, normalizedRange);
  if (params.kind === 'dependents') return resolveFormulaDependents(workbook, sheet, normalizedRange);
  if (params.kind === 'current-region') {
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
  if (params.kind === 'conditional-format' || params.kind === 'conditional-format-all' || params.kind === 'conditional-format-same') {
    const firstType = sheet.conditionalFormats[0]?.type;
    for (const rule of sheet.conditionalFormats) {
      if (params.kind === 'conditional-format-same' && firstType && rule.type !== firstType) continue;
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
  if (params.kind === 'data-validation' || params.kind === 'data-validation-all' || params.kind === 'data-validation-same') {
    const firstType = sheet.dataValidations[0]?.type;
    for (const rule of sheet.dataValidations) {
      if (params.kind === 'data-validation-same' && firstType && rule.type !== firstType) continue;
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
        case 'constant-numbers':
          match = Boolean(cell && typeof cell.value === 'number' && !cell.formula);
          break;
        case 'constant-text':
          match = Boolean(cell && typeof cell.value === 'string' && cell.value !== '' && !cell.formula);
          break;
        case 'constant-logical':
          match = Boolean(cell && typeof cell.value === 'boolean' && !cell.formula);
          break;
        case 'constant-errors':
          match = isErrorCell(cell) && !cell?.formula;
          break;
        case 'formulas':
          match = Boolean(cell?.formula);
          break;
        case 'formula-numbers':
          match = Boolean(cell?.formula && typeof cell.formulaValue === 'number');
          break;
        case 'formula-text':
          match = Boolean(cell?.formula && typeof cell.formulaValue === 'string');
          break;
        case 'formula-logical':
          match = Boolean(cell?.formula && typeof cell.formulaValue === 'boolean');
          break;
        case 'formula-errors':
          match = Boolean(cell?.formula && isErrorCell(cell));
          break;
        case 'comments':
          match = sheet.review.getThreadsAt(row, column).length > 0;
          break;
        case 'notes':
          match = sheet.review.hasNoteAt(row, column);
          break;
        case 'comments-notes':
          match = sheet.review.getThreadsAt(row, column).length > 0 || sheet.review.hasNoteAt(row, column);
          break;
        case 'errors':
          match = isErrorCell(cell);
          break;
        case 'visible':
          match = !sheet.hiddenRows.has(row) && !sheet.hiddenColumns.has(column);
          break;
        case 'conditional-format':
        case 'conditional-format-all':
        case 'conditional-format-same':
          match = conditionalCells.has(`${row}:${column}`);
          break;
        case 'data-validation':
        case 'data-validation-all':
        case 'data-validation-same':
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
  if (transfer === 'copy' && next.presentation?.kind === 'barcode' && next.presentation.source.kind === 'formula') {
    next.presentation = {
      ...next.presentation,
      source: {
        ...next.presentation.source,
        formula: shiftFormula(next.presentation.source.formula, rowDelta, colDelta),
      },
    };
  }
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
  if (sourceRange.endRow < sourceRange.startRow || sourceRange.endColumn < sourceRange.startColumn
    || sourceRange.endRow >= MAX_SHEET_ROW_COUNT || sourceRange.endColumn >= MAX_SHEET_COLUMN_COUNT) throw new Error('Clipboard source range is invalid');
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

function coordinatesFromKey(key: string): { row: number; column: number } {
  if (!/^\d+:\d+$/.test(key)) throw new Error(`Invalid cell metadata key: ${key}`);
  const [rowText, columnText] = key.split(':');
  const row = Number(rowText);
  const column = Number(columnText);
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column)) throw new Error(`Invalid cell metadata key: ${key}`);
  return { row, column };
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
    const { row, column } = coordinatesFromKey(key);
    return Number.isInteger(row) && Number.isInteger(column) && contains(row, column);
  }).map(([key, value]) => ({ key, value: structuredClone(value) })) : undefined;
  const commentCells = include.commentsNotes ? sheet.review.threadEntries().filter((thread) => contains(thread.row, thread.column)).map((thread) => keyFor(thread.row, thread.column)) : undefined;
  const comments = include.commentsNotes ? sheet.review.threadEntries().filter((thread) => contains(thread.row, thread.column)) : undefined;
  return { notes, hyperlinks, commentCells, comments };
}

function applyPasteSnapshot(workbook: WorkbookModel, sheet: WorksheetModel, snapshot: PasteSnapshot): void {
  if (snapshot.workbookTheme) workbook.setTheme(snapshot.workbookTheme);
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
    if (snapshot.notes !== undefined) {
      for (const entry of sheet.review.noteEntries()) if (rangeContains(range, entry.row, entry.column)) sheet.review.removeNote(entry.row, entry.column);
    }
    if (snapshot.hyperlinks !== undefined) {
      for (const key of [...sheet.hyperlinks.keys()]) {
        const { row, column } = coordinatesFromKey(key);
        if (Number.isInteger(row) && Number.isInteger(column) && rangeContains(range, row, column)) sheet.hyperlinks.delete(key);
      }
    }
    if (snapshot.comments !== undefined || snapshot.commentCells !== undefined) {
      for (const thread of sheet.review.threadEntries()) if (rangeContains(range, thread.row, thread.column)) sheet.review.removeThread(thread.id);
    }
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
      notes.push({
        key: keyFor(row, column),
        value: { ...structuredClone(entry.value), id: `${entry.value.id}@paste:${row}:${column}` },
      });
    }
    after.notes = notes;
    const comments = after.comments ?? [];
    for (const entry of metadata.comments) {
      const row = targetRange.startRow + entry.rowOffset;
      const column = targetRange.startColumn + entry.columnOffset;
      comments.push({
        ...structuredClone(entry.value),
        id: `${entry.value.id}@paste:${row}:${column}`,
        sheetId: params.sheetId,
        row,
        column,
      });
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
    const targetWidths = metadata.columnWidths.map((entry) => ({
      column: targetRange.startColumn + entry.offset,
      widthPx: entry.widthPx,
    }));
    after.columnWidths = targetWidths;
    if (params.transfer === 'move') {
      const targetColumns = new Set(targetWidths.map((entry) => entry.column));
      after.columnWidths.push(...metadata.columnWidths
        .map((entry) => ({ column: source.startColumn + entry.offset, widthPx: undefined }))
        .filter((entry) => !targetColumns.has(entry.column)));
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

export interface RangeMoveMutationParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetOrigin: { row: number; column: number };
}

function isRangeMoveMutation(value: unknown): value is RangeMoveMutationParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || !isRange(value.sourceRange)
    || value.sourceRange.sheetId !== value.sheetId || !isRecord(value.targetOrigin)) return false;
  const { row, column } = value.targetOrigin;
  const height = value.sourceRange.endRow - value.sourceRange.startRow + 1;
  const width = value.sourceRange.endColumn - value.sourceRange.startColumn + 1;
  return typeof row === 'number' && Number.isSafeInteger(row) && row >= 0
    && typeof column === 'number' && Number.isSafeInteger(column) && column >= 0
    && Number.isSafeInteger(row + height - 1)
    && Number.isSafeInteger(column + width - 1);
}

function rangeMoveAffectedRanges(params: RangeMoveMutationParams): [RangeRef, RangeRef] {
  const { sourceRange, targetOrigin } = params;
  const rowDelta = targetOrigin.row - sourceRange.startRow;
  const columnDelta = targetOrigin.column - sourceRange.startColumn;
  return [
    structuredClone(sourceRange),
    {
      sheetId: params.sheetId,
      startRow: sourceRange.startRow + rowDelta,
      endRow: sourceRange.endRow + rowDelta,
      startColumn: sourceRange.startColumn + columnDelta,
      endColumn: sourceRange.endColumn + columnDelta,
    },
  ];
}

export function applyRangeMoveMutation(context: CommandContext, params: RangeMoveMutationParams): CommandResult {
  if (!isRangeMoveMutation(params)) throw new Error('Invalid range.move mutation payload');
  const [sourceRange, targetRange] = rangeMoveAffectedRanges(params);
  const sheet = context.workbook.getSheet(params.sheetId);
  const overwritten = sheet.cells.getRegion(
    targetRange.startRow,
    targetRange.endRow,
    targetRange.startColumn,
    targetRange.endColumn,
  );
  const inverse: MutationInfo[] = [{
    id: 'range.move',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: {
      sheetId: params.sheetId,
      sourceRange: structuredClone(targetRange),
      targetOrigin: { row: sourceRange.startRow, column: sourceRange.startColumn },
    } satisfies RangeMoveMutationParams,
    affectedRanges: [structuredClone(targetRange), structuredClone(sourceRange)],
  }, ...overwritten.map(({ row, column, cell }) => ({
    id: 'cell.restore',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: { sheetId: params.sheetId, row, column, previous: structuredClone(cell) },
    affectedRanges: [{ sheetId: params.sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }],
  }))];
  context.applyMutation({
    id: 'range.move',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: structuredClone(params),
    affectedRanges: [structuredClone(sourceRange), structuredClone(targetRange)],
    inverse,
    apply: () => StructuralTransform.apply(context.workbook, {
      kind: 'move-range',
      sheetId: params.sheetId,
      sourceRange,
      targetOrigin: structuredClone(params.targetOrigin),
    }, context.structuralReferenceOwners),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges: [sourceRange, targetRange] };
}

export function registerEditingCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<RangeMoveMutationParams>({
    id: 'range.move',
    handler: (item, context) => {
      if (!isRangeMoveMutation(item.params)) throw new Error('Invalid range.move mutation payload');
      return StructuralTransform.apply(context.workbook, {
        kind: 'move-range',
        sheetId: item.params.sheetId,
        sourceRange: item.params.sourceRange,
        targetOrigin: item.params.targetOrigin,
      }, context.structuralReferenceOwners);
    },
    metadata: {
      schema: { name: 'RangeMove', validate: isRangeMoveMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: rangeMoveAffectedRanges, mode: 'exact' },
      historyRebase: { kind: 'invalidate', reason: 'range moves have no canonical history transform' },
      inversePolicy: { allowedMutationIds: ['range.move', 'cell.restore'], minCount: 1 },
    },
  });

  runtime.registry.registerMutation<PasteMutationParams>({
    id: 'range.paste',
    handler: (item, context) => {
    if (!isPasteMutation(item.params)) throw new Error('Invalid range.paste mutation payload');
    const params = item.params;
    if (item.sheetId !== params.sheetId) throw new Error('Invalid range.paste mutation target sheet');
    if (params.clearSource && params.sourceRange?.sheetId !== params.sheetId) {
      throw new Error('UNSUPPORTED_FEATURE: cross-sheet cut/paste requires a canonical structural move patch');
    }
    context.workbook.getSheet(params.clipboard.range.sheetId);
    const targetSheet = context.workbook.getSheet(params.sheetId);
    applyPasteSnapshot(context.workbook, targetSheet, params.snapshot);
    },
    metadata: {
      schema: { name: 'PasteMutation', validate: isPasteMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: pasteAffectedRanges, mode: 'exact' },
      historyRebase: {
        kind: 'invalidate',
        reason: 'paste moves or changes workbook theme without a canonical history transform',
        when: (mutation) => isPasteMutation(mutation.params)
          && (mutation.params.transfer === 'move'
            || mutation.params.clearSource === true
            || mutation.params.snapshot.workbookTheme !== undefined),
      },
      inverseIds: ['range.paste'],
    },
  });

  runtime.registry.registerCommand<PasteRangeParams>({
    id: 'sheet.range.paste',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (!isPasteSpecialSpec(params.spec)) throw new Error('Paste Special requires a canonical specification');
      const requiresInterpretation = Boolean(params.clipboard.representations?.length || params.clipboard.html !== undefined || params.clipboard.text !== undefined);
      if (requiresInterpretation && !isCellInputInterpretationContext(params.inputContext)) {
        throw new Error('External clipboard text requires a cell input interpretation context');
      }
      const clipboard = requiresInterpretation
        ? parseClipboardPayload(params.clipboard, params.inputContext!)
        : structuredClone(params.clipboard);
      const sourceRange = params.clipboard.range;
      const transfer = params.transfer;
      if (transfer !== 'copy' && transfer !== 'move' || params.clipboard.transfer !== transfer) {
        throw new Error('Paste transfer must match the canonical clipboard transfer');
      }
      if (transfer === 'move' && (!sourceRange || sourceRange.sheetId.length === 0)) {
        throw new Error('Move clipboard payload must include a source range');
      }
      const canonicalParams: Omit<PasteRangeParams, 'inputContext'> = {
        sheetId: params.sheetId,
        targetOrigin: structuredClone(params.targetOrigin),
        clipboard,
        transfer: params.transfer,
        spec: structuredClone(params.spec),
      };
      const targetRange = assertPastePreconditions(context.workbook, canonicalParams);
      if (transfer === 'move' && sourceRange.sheetId !== params.sheetId) {
        throw new Error('UNSUPPORTED_FEATURE: cross-sheet cut/paste requires a canonical structural move patch');
      }
      if (transfer === 'move' && sourceRange?.sheetId === params.sheetId && isDefaultMoveSpec(params.spec)) {
        return applyRangeMoveMutation(context, {
          sheetId: params.sheetId,
          sourceRange,
          targetOrigin: params.targetOrigin,
        });
      }
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
      const touchedRanges = transfer === 'move' && sourceRange && sourceRange.sheetId === params.sheetId ? [targetCellRange, sourceRange] : [targetCellRange];
      const clearsCells = params.spec.content !== 'none' && !params.spec.skipBlanks ? [structuredClone(targetCellRange)] : [];
      if (transfer === 'move' && sourceRange) clearsCells.push(structuredClone(sourceRange));
      const clearsMetadata = Object.values(params.spec.metadata).some(Boolean)
        ? [structuredClone(targetCellRange), ...(transfer === 'move' && sourceRange ? [structuredClone(sourceRange)] : [])]
        : [];
      const sparseWidths = (targetSheet: WorksheetModel, ranges: RangeRef[]) => {
        const columns = new Set<number>();
        for (const range of ranges) {
          for (let column = range.startColumn; column <= range.endColumn; column += 1) columns.add(column);
        }
        return [...columns].sort((left, right) => left - right).map((column) => ({
          column,
          widthPx: targetSheet.columnWidthsPx[column],
        }));
      };
      const before: PasteSnapshot = {
        clearRanges: clearsCells.filter((range) => range.sheetId === params.sheetId),
        clearMetadataRanges: clearsMetadata.filter((range) => range.sheetId === params.sheetId),
        cells: snapshotCells(sheet, touchedRanges),
        ...snapshotMetadata(sheet, touchedRanges, params.spec.metadata),
        ...(params.spec.metadata.validation ? { validations: structuredClone(sheet.dataValidations) } : {}),
        ...(params.spec.metadata.conditionalFormats ? { conditionalFormats: structuredClone(sheet.conditionalFormats) } : {}),
        ...(params.spec.metadata.columnWidths ? { columnWidths: sparseWidths(sheet, touchedRanges) } : {}),
        ...(params.spec.formatting === 'source-theme' ? { workbookTheme: structuredClone(context.workbook.theme) } : {}),
      };
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
        ...(params.spec.formatting === 'source-theme' ? { workbookTheme: structuredClone(clipboard.rangeMetadata.sourceWorkbookThemeRef!) } : {}),
      };
      const inRanges = (row: number, column: number, ranges: RangeRef[]) => ranges.some((range) => rangeContains(range, row, column));
      after.cells = after.cells.filter((entry) => !inRanges(entry.row, entry.column, after.clearRanges ?? []));
      if (after.notes) after.notes = after.notes.filter((entry) => {
        const { row, column } = coordinatesFromKey(entry.key);
        return !inRanges(row, column, after.clearMetadataRanges ?? []);
      });
      if (after.hyperlinks) after.hyperlinks = after.hyperlinks.filter((entry) => {
        const { row, column } = coordinatesFromKey(entry.key);
        return !inRanges(row, column, after.clearMetadataRanges ?? []);
      });
      if (after.commentCells) after.commentCells = after.commentCells.filter((key) => {
        const { row, column } = coordinatesFromKey(key);
        return !inRanges(row, column, after.clearMetadataRanges ?? []);
      });
      if (after.comments) after.comments = after.comments.filter((entry) => !inRanges(entry.row, entry.column, after.clearMetadataRanges ?? []));
      const afterCells = new Map(after.cells.map((entry) => [keyFor(entry.row, entry.column), entry]));
      const beforeCellKeys = new Set(before.cells.map((entry) => keyFor(entry.row, entry.column)));
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
          if (next !== undefined) {
            const key = keyFor(row, column);
            if (!beforeCellKeys.has(key) && !inRanges(row, column, before.clearRanges ?? [])) {
              before.cells.push({ row, column });
              beforeCellKeys.add(key);
            }
            setAfterCell(row, column, next);
          }
      }
      after.cells = [...afterCells.values()];
      applyPasteMetadataPlan(context.workbook, canonicalParams, targetRange, after);
      const mutationParams: PasteMutationParams = {
        ...canonicalParams,
        sourceExtent: { rows: rowCount, columns: columnCount },
        transfer,
        sourceRange: transfer === 'move' ? structuredClone(sourceRange) : undefined,
        clearSource: transfer === 'move',
        snapshot: after,
      };
      const affectedRanges = pasteAffectedRanges(mutationParams);
      context.applyMutation({
        id: 'range.paste',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: mutationParams,
        affectedRanges,
        inverse: [{
          id: 'range.paste',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { ...canonicalParams, sourceExtent: { rows: rowCount, columns: columnCount }, sourceRange: transfer === 'move' ? structuredClone(sourceRange) : undefined, clearSource: transfer === 'move', snapshot: before },
          affectedRanges,
        }],
        apply: () => applyPasteSnapshot(context.workbook, sheet, after),
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
  const cellShiftMutationHandler = (operation: CellShiftParams['operation'], id: 'cells.inserted' | 'cells.deleted') => (item: { params: unknown }, context: CommandContext) => {
      if (!isCellShiftMutation(item.params) || item.params.operation !== operation) throw new Error(`Invalid ${id} mutation payload`);
      validateCellShiftEnvelope(item.params, context);
      return StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: item.params.sheetId, sourceRange: item.params.range, operation: item.params.operation, axis: item.params.axis }, context.structuralReferenceOwners);
    };
  runtime.registry.registerMutation<CellShiftParams>({ id: 'cells.inserted', handler: cellShiftMutationHandler('insert', 'cells.inserted'), metadata: { schema: { name: 'CellShiftInsert', validate: (value: unknown): value is CellShiftParams => isCellShiftMutation(value) && value.operation === 'insert' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.affectedBand)], mode: 'exact' }, historyRebase: { kind: 'invalidate', reason: 'cell shifts have no canonical history transform' }, inverseIds: ['cells.inserted.restore'] } });
  runtime.registry.registerMutation<CellShiftParams>({ id: 'cells.deleted', handler: cellShiftMutationHandler('delete', 'cells.deleted'), metadata: { schema: { name: 'CellShiftDelete', validate: (value: unknown): value is CellShiftParams => isCellShiftMutation(value) && value.operation === 'delete' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.affectedBand)], mode: 'exact' }, historyRebase: { kind: 'invalidate', reason: 'cell shifts have no canonical history transform' }, inverseIds: ['cells.deleted.restore'] } });
  const cellShiftRestoreMutationHandler = (operation: CellShiftParams['operation'], id: 'cells.inserted.restore' | 'cells.deleted.restore') => (item: { params: unknown }, context: CommandContext) => {
      if (!isCellShiftRestoreMutation(item.params) || item.params.spec.operation !== operation) throw new Error(`Invalid ${id} mutation payload`);
      validateCellShiftEnvelope(item.params.spec, context);
      const plan = planCellShift(context.workbook, item.params.spec);
      const sheet = context.workbook.getSheet(item.params.spec.sheetId);
      const effect = StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: item.params.spec.sheetId, sourceRange: item.params.spec.range, operation: operation === 'insert' ? 'delete' : 'insert', axis: item.params.spec.axis }, context.structuralReferenceOwners);
      for (let row = plan.band.startRow; row <= plan.band.endRow; row += 1) for (let column = plan.band.startColumn; column <= plan.band.endColumn; column += 1) sheet.cells.delete(row, column);
      for (const entry of item.params.cells) sheet.cells.set(entry.row, entry.column, structuredClone(entry.cell));
      return effect;
    };
  runtime.registry.registerMutation<CellShiftRestoreParams>({ id: 'cells.inserted.restore', handler: cellShiftRestoreMutationHandler('insert', 'cells.inserted.restore'), metadata: { schema: { name: 'CellShiftInsertRestore', validate: (value: unknown): value is CellShiftRestoreParams => isCellShiftRestoreMutation(value) && value.spec.operation === 'insert' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.spec.affectedBand)], mode: 'exact' }, historyRebase: { kind: 'invalidate', reason: 'cell-shift restoration has no canonical history transform' }, inverseIds: ['cells.inserted'] } });
  runtime.registry.registerMutation<CellShiftRestoreParams>({ id: 'cells.deleted.restore', handler: cellShiftRestoreMutationHandler('delete', 'cells.deleted.restore'), metadata: { schema: { name: 'CellShiftDeleteRestore', validate: (value: unknown): value is CellShiftRestoreParams => isCellShiftRestoreMutation(value) && value.spec.operation === 'delete' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.spec.affectedBand)], mode: 'exact' }, historyRebase: { kind: 'invalidate', reason: 'cell-shift restoration has no canonical history transform' }, inverseIds: ['cells.deleted'] } });
  const createCellShiftParams = (params: Omit<CellShiftParams, 'affectedBand'>, context: { workbook: WorkbookModel }) => {
    const plan = planCellShift(context.workbook, params);
    const canonicalParams: CellShiftParams = { ...params, affectedBand: plan.band };
    const sheet = context.workbook.getSheet(params.sheetId);
    const snapshot: Array<{ row: number; column: number; cell: CellData }> = [];
    forEachCell(sheet, plan.band, (row, column, cell) => { if (cell) snapshot.push({ row, column, cell: structuredClone(cell) }); });
    const affectedRanges: RangeRef[] = [structuredClone(plan.band)];
    return { canonicalParams, snapshot, affectedRanges };
  };
  runtime.registry.registerCommand<Omit<CellShiftParams, 'affectedBand'>>({ id: 'sheet.cells.insert', execute: (params, context) => { const { canonicalParams, snapshot, affectedRanges } = createCellShiftParams(params, context); context.applyMutation({ id: 'cells.inserted', unitId: context.workbook.unitId, sheetId: params.sheetId, params: canonicalParams, affectedRanges, inverse: [{ id: 'cells.inserted.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { spec: canonicalParams, cells: snapshot }, affectedRanges }], apply: () => StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: params.sheetId, sourceRange: params.range, operation: 'insert', axis: params.axis }, context.structuralReferenceOwners) }); return { operationId: context.operationId, mutationCount: 1, affectedRanges }; } });
  runtime.registry.registerCommand<Omit<CellShiftParams, 'affectedBand'>>({ id: 'sheet.cells.delete', execute: (params, context) => { const { canonicalParams, snapshot, affectedRanges } = createCellShiftParams(params, context); context.applyMutation({ id: 'cells.deleted', unitId: context.workbook.unitId, sheetId: params.sheetId, params: canonicalParams, affectedRanges, inverse: [{ id: 'cells.deleted.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { spec: canonicalParams, cells: snapshot }, affectedRanges }], apply: () => StructuralTransform.apply(context.workbook, { kind: 'cell-shift', sheetId: params.sheetId, sourceRange: params.range, operation: 'delete', axis: params.axis }, context.structuralReferenceOwners) }); return { operationId: context.operationId, mutationCount: 1, affectedRanges }; } });

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
