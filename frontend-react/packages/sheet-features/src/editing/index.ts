import type {
  CellData,
  CellStyle,
  WorksheetPane,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
  BorderLine,
  BorderPlacement,
} from '@react-sheets/core-model';
import { type CellShiftSpec } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import {
  copyRangeToClipboardData,
  parseClipboardPayload,
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

function cellShiftBand(workbook: WorkbookModel, params: CellShiftSpec): RangeRef {
  const sheet = workbook.getSheet(params.sheetId);
  const range = normalizeRanges([{ ...params.range, sheetId: params.sheetId }])[0]!;
  if (range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) throw new Error('Cell shift range exceeds worksheet bounds');
  return params.axis === 'row'
    ? { ...range, endRow: sheet.rowCount - 1 }
    : { ...range, endColumn: sheet.columnCount - 1 };
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
};

function isPasteMutation(value: unknown): value is PasteMutationParams {
  if (!isRecord(value) || !isRecord(value.clipboard) || !isRecord(value.clipboard.sourceExtent)) return false;
  const clipboard = value.clipboard;
  const sourceExtent = clipboard.sourceExtent as Record<string, unknown>;
  return typeof value.sheetId === 'string'
    && isRecord(value.targetOrigin) && Number.isInteger(value.targetOrigin.row) && Number.isInteger(value.targetOrigin.column)
    && (value.transfer === 'copy' || value.transfer === 'move')
    && clipboard.transfer === value.transfer
    && clipboard.schema === 'SparseClipboardPayload'
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
    && isPasteSpecialSpec(value.spec)
    && isPasteSpecialSpecSupported(value.spec, clipboard as unknown as ClipboardPayload)
    && isRecord(value.sourceExtent) && Number.isInteger(value.sourceExtent.rows) && Number.isInteger(value.sourceExtent.columns)
    && (value.transfer === 'move'
      ? isRange(value.sourceRange) && value.clearSource === true
      : value.sourceRange === undefined && value.clearSource === false);
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


function isCellShiftMutation(value: unknown): value is CellShiftParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range)
    && isRange(value.affectedBand)
    && value.range.sheetId === value.sheetId
    && value.affectedBand.sheetId === value.sheetId
    && (value.operation === 'insert' || value.operation === 'delete')
    && (value.axis === 'row' || value.axis === 'column');
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

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId && left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
}

function assertPastePreconditions(workbook: WorkbookModel, params: PasteRangeParams): RangeRef {
  const sheet = workbook.getSheet(params.sheetId);
  if (!Number.isInteger(params.targetOrigin.row) || !Number.isInteger(params.targetOrigin.column)
    || params.targetOrigin.row < 0 || params.targetOrigin.column < 0) throw new Error('Paste target origin is invalid');
  const sourceRange = params.clipboard.range;
  workbook.getSheet(sourceRange.sheetId);
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

export function registerEditingCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<PasteMutationParams>({
    id: 'range.paste',
    metadata: {
      schema: { name: 'PasteMutation', validate: isPasteMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: pasteAffectedRanges, mode: 'declared' },
    },
  });

  runtime.registry.registerCommand<PasteRangeParams>({
    id: 'sheet.range.paste',
    execute: (params, context) => {
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
      const rowCount = clipboard.sourceExtent.rows;
      const columnCount = clipboard.sourceExtent.columns;
      const affectedRanges: RangeRef[] = [structuredClone(targetRange)];
      if (transfer === 'move' && sourceRange) affectedRanges.push(structuredClone(sourceRange));
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
        },
        affectedRanges,
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
        lastResult = context.executeCommand('sheet.style.set', { sheetId: params.sheetId, range, style: params.style });
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
        results.push(context.executeCommand('sheet.style.setMulti', {
          sheetId: params.sheetId,
          ranges: params.ranges,
          style,
        }));
      }
      if (params.border) {
        results.push(context.executeCommand('sheet.borders.set', {
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
      const affectedRanges = normalizeRanges(params.ranges);
      context.applyMutation({
        id: 'style.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
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
      return context.executeCommand('sheet.freeze.set', { sheetId: params.sheetId, pane });
    },
  });

  runtime.registry.registerMutation<SheetViewParams>({
    id: 'view.set',
    metadata: {
      schema: { name: 'SheetView', validate: isSheetViewMutation },
      permission: { capability: 'sheet.view.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
    },
  });

  runtime.registry.registerCommand<SheetViewParams>({
    id: 'sheet.view.set',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'view.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<CellShiftParams>({ id: 'cells.inserted', metadata: { schema: { name: 'CellShiftInsert', validate: (value: unknown): value is CellShiftParams => isCellShiftMutation(value) && value.operation === 'insert' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.affectedBand)], mode: 'exact' } } });
  runtime.registry.registerMutation<CellShiftParams>({ id: 'cells.deleted', metadata: { schema: { name: 'CellShiftDelete', validate: (value: unknown): value is CellShiftParams => isCellShiftMutation(value) && value.operation === 'delete' }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [structuredClone(params.affectedBand)], mode: 'exact' } } });
  const createCellShiftParams = (params: Omit<CellShiftParams, 'affectedBand'>, context: { workbook: WorkbookModel }) => {
    const band = cellShiftBand(context.workbook, params);
    const canonicalParams: CellShiftParams = { ...params, affectedBand: band };
    const affectedRanges: RangeRef[] = [structuredClone(band)];
    return { canonicalParams, affectedRanges };
  };
  runtime.registry.registerCommand<Omit<CellShiftParams, 'affectedBand'>>({ id: 'sheet.cells.insert', execute: (params, context) => { const { canonicalParams, affectedRanges } = createCellShiftParams(params, context); context.applyMutation({ id: 'cells.inserted', unitId: context.workbook.unitId, sheetId: params.sheetId, params: canonicalParams, affectedRanges }); return { operationId: context.operationId, mutationCount: 1, affectedRanges }; } });
  runtime.registry.registerCommand<Omit<CellShiftParams, 'affectedBand'>>({ id: 'sheet.cells.delete', execute: (params, context) => { const { canonicalParams, affectedRanges } = createCellShiftParams(params, context); context.applyMutation({ id: 'cells.deleted', unitId: context.workbook.unitId, sheetId: params.sheetId, params: canonicalParams, affectedRanges }); return { operationId: context.operationId, mutationCount: 1, affectedRanges }; } });

  runtime.registry.registerMutation<{ sourceSheetId: string; newId: string; newName: string }>({
    id: 'sheet.duplicated',
    metadata: {
      schema: { name: 'DuplicateSheet', validate: isSheetDuplicateMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
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
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'sheet.hidden',
    metadata: {
      schema: { name: 'SheetHidden', validate: isSheetIdMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
    },
  });
  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'sheet.unhidden',
    metadata: {
      schema: { name: 'SheetUnhidden', validate: isSheetIdMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
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
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string; toIndex: number }>({
    id: 'sheet.reordered',
    metadata: {
      schema: { name: 'ReorderSheet', validate: isSheetReorderedMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; toIndex: number }>({
    id: 'sheet.reorder',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.reordered',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string; color?: string }>({
    id: 'sheet.tabColor',
    metadata: {
      schema: { name: 'TabColor', validate: isTabColorMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; color?: string }>({
    id: 'sheet.tabColor.set',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.tabColor',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
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
