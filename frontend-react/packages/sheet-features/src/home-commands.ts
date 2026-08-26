import type {
  CellData,
  CellStyle,
  ConditionalFormatRule,
  DataSourceManifest,
  DrawingObject,
  AutoFilterModel,
  RangeRef,
  SheetDataRegion,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { protectionResolver } from '@react-sheets/core-model';
import type { CommandContext, CommandResult, CommandRuntime } from '@react-sheets/command-runtime';
import { normalizeAutoFilterModel, validateDataInput, type DataSortParams } from './data-features';
import { resolveActiveAutoFilter, resolveFilterOwner, validateFilterOwnership } from './sheet-table-features';
import { copyRangeToClipboardData, createPasteSpecialSpec, shiftFormula, type ClipboardPayload } from './clipboard';
import { resolveGoToRange, resolveGoToSpecial, type GoToSpecialKind, type GoToSpecialParams } from './editing';
import { isFormulaError, isSpillChild, type FormulaError, type ScalarValue } from '@react-sheets/formula-engine';
import { parseReplacementValue, replacementCell, replaceFindText } from './find-replace';
import { isCellInputInterpretationContext, type CellInputInterpretationContext } from './text-input';
import { planFill, validateFillPlan, type FillDirection, type FillMode, type FillPlanParams, type FillWrite } from './fill-series';

/**
 * The Home tab owns high-level semantic commands.  Low-level mutations remain
 * registered by the sheet feature modules; commands in this file deliberately
 * compose those mutations through CommandRuntime so a user gesture still has
 * one history entry and one collaboration operation.
 */

export interface HomeRangeParams {
  sheetId: string;
  range: RangeRef;
}

export interface HomeFillParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetRange: RangeRef;
  direction: FillDirection;
  mode: FillMode;
  dateSystem?: '1900' | '1904';
}

interface FillMutationWrite {
  row: number;
  column: number;
  before?: CellData;
  after?: CellData;
}

interface FillMutationParams extends HomeFillParams {
  writes: FillMutationWrite[];
}

export interface AutoSumParams {
  sheetId: string;
  range: RangeRef;
  target?: { row: number; column: number };
  functionName?: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';
  /** Generate one formula for each selected column when target is omitted. */
  byColumn?: boolean;
}

/**
 * Application-owned data-source loading is asynchronous. The loader must
 * resolve all blocks first, then submit this complete payload to the command
 * runtime; the mutation itself is synchronous, replayable and undoable.
 */
export interface DataRegionMaterializeParams {
  sheetId: string;
  region: SheetDataRegion;
  regionIndex: number;
  manifest: DataSourceManifest;
  willRemoveSource: boolean;
  range: RangeRef;
  previousCells: Array<{ row: number; column: number; cell: CellData }>;
  materializedCells: Array<{ row: number; column: number; cell: CellData }>;
  materializedCellCount?: number;
}

export interface ReplaceRangeParams {
  sheetId: string;
  find: string;
  replace: string;
  inputContext: CellInputInterpretationContext;
  range?: RangeRef;
  matchCase?: boolean;
  entireCell?: boolean;
  scope?: 'sheet' | 'workbook';
  searchIn?: 'values' | 'formulas' | 'both';
}

/**
 * The only replacement value contract.  A replacement is parsed once before
 * any cell mutation, so `0` cannot be confused with a failed/falsy parse and
 * every caller observes the same typed result.
 */
export { parseReplacementValue, replacementCell, type ReplacementValue } from './find-replace';

export interface FilterToggleParams {
  sheetId: string;
  range?: RangeRef;
  autoFilter?: AutoFilterModel;
}

export interface FilterCriteriaParams {
  sheetId: string;
  range?: RangeRef;
}

export interface FilterSortParams {
  sheetId: string;
  column: number;
  ascending: boolean;
}

export interface SortReapplyParams {
  sheetId: string;
}

export interface FormatPainterParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetRange: RangeRef;
  /** Host-only state; it never changes the model. */
  continuous?: boolean;
}

export interface RangeMoveParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetOrigin: { row: number; column: number };
  inputContext: CellInputInterpretationContext;
}

interface FormatPainterCell {
  style?: CellStyle;
  numberFormat?: string;
}

interface FormatPainterMutationParams {
  sheetId: string;
  targetRange: RangeRef;
  styles: FormatPainterCell[][];
}

export interface HomeStylePreset {
  id: string;
  name: string;
  style: Partial<CellStyle>;
}

/** Static, serializable gallery. Applying a preset writes its styleId and
 * canonical style to cells; no independent per-workbook style table is made. */
export const CELL_STYLE_PRESETS: readonly HomeStylePreset[] = [
  { id: 'normal', name: 'Normal', style: {} },
  { id: 'good', name: 'Good', style: { background: '#e2f0d9', textColor: '#006100' } },
  { id: 'bad', name: 'Bad', style: { background: '#ffc7ce', textColor: '#9c0006' } },
  { id: 'neutral', name: 'Neutral', style: { background: '#ffeb9c', textColor: '#9c6500' } },
  { id: 'title', name: 'Title', style: { bold: true, fontSizePx: 24, horizontalAlignment: 'center' } },
  { id: 'heading1', name: 'Heading 1', style: { bold: true, fontSizePx: 18.6666666667 } },
  { id: 'heading2', name: 'Heading 2', style: { bold: true, fontSizePx: 16 } },
  { id: 'total', name: 'Total', style: { bold: true, borders: { top: { style: 'double', color: '#334155' } } } },
] as const;

type AppliedSortState = DataSortParams & { revision: number };
type SheetWithHomeState = WorksheetModel & {
  appliedSortState?: AppliedSortState;
  sortRevision?: number;
};
type DrawingWithHomeState = DrawingObject & { visible?: boolean };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isFiniteInt(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && isFiniteInt(value.startRow) && isFiniteInt(value.endRow)
    && isFiniteInt(value.startColumn) && isFiniteInt(value.endColumn)
    && value.endRow >= value.startRow && value.endColumn >= value.startColumn;
}

function normalizeRange(range: RangeRef, sheetId: string): RangeRef {
  if (range.sheetId !== sheetId) throw new Error('Range must target the command sheet');
  return {
    sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function rangeEquals(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow === right.startRow && left.endRow === right.endRow
    && left.startColumn === right.startColumn && left.endColumn === right.endColumn;
}

function cellRange(sheetId: string, row: number, column: number): RangeRef {
  return { sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column };
}

function scriptEntryIntent(sheetId: string, row: number, column: number, candidate: CellData) {
  return {
    kind: 'script' as const,
    target: { sheetId, row, column },
    candidate: structuredClone(candidate),
    validationDecision: { status: 'not-applicable' as const },
  };
}

function rangeEntryIntent(
  kind: 'script' | 'formula-result',
  sheetId: string,
  startRow: number,
  startColumn: number,
  values: CellData[][],
) {
  return {
    kind,
    target: {
      sheetId,
      startRow,
      endRow: startRow + Math.max(0, values.length - 1),
      startColumn,
      endColumn: startColumn + Math.max(0, Math.max(1, ...values.map((row) => row.length)) - 1),
    },
    candidate: structuredClone(values),
    validationDecision: { status: 'not-applicable' as const },
  };
}

function rangeAffected(range: RangeRef): RangeRef[] {
  return [structuredClone(range)];
}

function homeResult(context: CommandContext, affectedRanges: RangeRef[], mutationCount = 0): CommandResult {
  return { operationId: context.operationId, mutationCount, affectedRanges };
}

function requireSheetId(value: unknown): asserts value is { sheetId: string } {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || value.sheetId.trim() === '') {
    throw new Error('Home command requires a sheetId');
  }
}

function requireRange(value: unknown, sheetId: string, field = 'range'): RangeRef {
  if (!isRecord(value) || !isRange(value[field])) throw new Error(`Home command requires a valid ${field}`);
  return normalizeRange(value[field] as RangeRef, sheetId);
}

function columnLabel(column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

type ResolvedAutoSumValue = ScalarValue | FormulaError;

function normalizeAutoSumValue(value: unknown): ResolvedAutoSumValue {
  if (Array.isArray(value)) throw new Error('AutoSum cannot use an unresolved array result');
  if (value === undefined || value === null) return null;
  if (isFormulaError(value)) return value;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('AutoSum cannot use a non-finite numeric result');
    return value;
  }
  if (typeof value === 'string' || typeof value === 'boolean') return value;
  throw new Error(`AutoSum resolved value has unsupported type: ${typeof value}`);
}

function resolveAutoSumValue(sheet: WorksheetModel, row: number, column: number, context: CommandContext): ResolvedAutoSumValue {
  const resolved = context.resolveCellValue?.(sheet, row, column);
  if (resolved !== undefined) return normalizeAutoSumValue(resolved);
  const cell = sheet.cells.get(row, column);
  if (cell?.formula !== undefined) {
    if (cell.formulaValue === undefined) throw new Error(`AutoSum formula result unavailable at ${sheet.id}!${row}:${column}`);
    return normalizeAutoSumValue(cell.formulaValue);
  }
  return normalizeAutoSumValue(cell?.value ?? null);
}

function isNumericResolvedValue(value: ResolvedAutoSumValue): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function contiguousNumericAbove(sheet: WorksheetModel, row: number, column: number, context: CommandContext): { startRow: number; endRow: number } | undefined {
  let endRow = row - 1;
  while (endRow >= 0 && !isNumericResolvedValue(resolveAutoSumValue(sheet, endRow, column, context))) endRow -= 1;
  if (endRow < 0) return undefined;
  let startRow = endRow;
  while (startRow > 0 && isNumericResolvedValue(resolveAutoSumValue(sheet, startRow - 1, column, context))) startRow -= 1;
  return { startRow, endRow };
}

function contiguousNumericLeft(sheet: WorksheetModel, row: number, column: number, context: CommandContext): { startColumn: number; endColumn: number } | undefined {
  let endColumn = column - 1;
  while (endColumn >= 0 && !isNumericResolvedValue(resolveAutoSumValue(sheet, row, endColumn, context))) endColumn -= 1;
  if (endColumn < 0) return undefined;
  let startColumn = endColumn;
  while (startColumn > 0 && isNumericResolvedValue(resolveAutoSumValue(sheet, row, startColumn - 1, context))) startColumn -= 1;
  return { startColumn, endColumn };
}

function hasNumericResolvedValue(sheet: WorksheetModel, range: RangeRef, context: CommandContext): boolean {
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (isNumericResolvedValue(resolveAutoSumValue(sheet, row, column, context))) return true;
    }
  }
  return false;
}

function formulaForAutoSum(sheet: WorksheetModel, targetRow: number, targetColumn: number, functionName: AutoSumParams['functionName'], selectedRange: RangeRef, context: CommandContext): string {
  const name = functionName ?? 'SUM';
  const above = contiguousNumericAbove(sheet, targetRow, targetColumn, context);
  if (above) return `=${name}(${columnLabel(targetColumn)}${above.startRow + 1}:${columnLabel(targetColumn)}${above.endRow + 1})`;
  const left = contiguousNumericLeft(sheet, targetRow, targetColumn, context);
  if (left) return `=${name}(${columnLabel(left.startColumn)}${targetRow + 1}:${columnLabel(left.endColumn)}${targetRow + 1})`;
  if (!hasNumericResolvedValue(sheet, selectedRange, context)) throw new Error('AutoSum source contains no numeric result');
  if (targetRow >= selectedRange.startRow && targetRow <= selectedRange.endRow
    && targetColumn >= selectedRange.startColumn && targetColumn <= selectedRange.endColumn) {
    throw new Error('AutoSum source range would include its target');
  }
  return `=${name}(${columnLabel(selectedRange.startColumn)}${selectedRange.startRow + 1}:${columnLabel(selectedRange.endColumn)}${selectedRange.endRow + 1})`;
}

function assertSafeAutoSumTarget(sheet: WorksheetModel, target: { row: number; column: number }): void {
  const targetRange = cellRange(sheet.id, target.row, target.column);
  const cell = sheet.cells.get(target.row, target.column);
  if (cell && (cell.formula !== undefined || cell.formulaValue !== undefined
    || (cell.value !== undefined && cell.value !== null))) {
    throw new Error(`AutoSum target ${sheet.id}!${target.row}:${target.column} is not blank`);
  }
  if (sheet.merges.some((merge) => rangesIntersect(merge.range, targetRange))) {
    throw new Error(`AutoSum target ${sheet.id}!${target.row}:${target.column} intersects a merged range`);
  }
  if (sheet.spillRanges.some((spill) => isSpillChild(spill, target.row, target.column))) {
    throw new Error(`AutoSum target ${sheet.id}!${target.row}:${target.column} is a formula spill child`);
  }
  const decision = protectionResolver.resolve({
    sheetId: sheet.id,
    rules: sheet.protectionRules,
    ranges: [targetRange],
    action: 'edit-cell',
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    readCellStyle: (row, column) => sheet.cells.get(row, column)?.style,
  });
  if (!decision.allowed) throw new Error(decision.reason ?? `AutoSum target ${sheet.id}!${target.row}:${target.column} is protected`);
}

function isValidAutoSumParams(value: unknown): value is AutoSumParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || !isRange(value.range)) return false;
  if (value.target !== undefined && (!isRecord(value.target) || !isFiniteInt(value.target.row) || !isFiniteInt(value.target.column))) return false;
  return value.functionName === undefined || ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].includes(String(value.functionName));
}

function isValidFillParams(value: unknown): value is HomeFillParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || !isRange(value.sourceRange) || !isRange(value.targetRange)) return false;
  return ['down', 'up', 'right', 'left'].includes(String(value.direction))
    && (value.mode === 'copy' || value.mode === 'series')
    && (value.dateSystem === undefined || value.dateSystem === '1900' || value.dateSystem === '1904');
}

function isValidReplaceParams(value: unknown): value is ReplaceRangeParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || typeof value.find !== 'string' || typeof value.replace !== 'string') return false;
  if (!isCellInputInterpretationContext(value.inputContext)) return false;
  return (value.range === undefined || isRange(value.range))
    && (value.matchCase === undefined || typeof value.matchCase === 'boolean')
    && (value.entireCell === undefined || typeof value.entireCell === 'boolean')
    && (value.scope === undefined || value.scope === 'sheet' || value.scope === 'workbook')
    && (value.searchIn === undefined || ['values', 'formulas', 'both'].includes(String(value.searchIn)));
}

function isValidFilterToggleParams(value: unknown): value is FilterToggleParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string') return false;
  return (value.range === undefined || isRange(value.range)) && (value.autoFilter === undefined || isRecord(value.autoFilter));
}

function isValidFilterCriteriaParams(value: unknown): value is FilterCriteriaParams {
  return isRecord(value) && typeof value.sheetId === 'string' && (value.range === undefined || isRange(value.range));
}

function isValidFilterSortParams(value: unknown): value is FilterSortParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && Number.isSafeInteger(value.column) && typeof value.ascending === 'boolean';
}

function isValidSortReapplyParams(value: unknown): value is SortReapplyParams {
  return isRecord(value) && typeof value.sheetId === 'string' && value.sheetId.length > 0;
}

function isValidFormatPainterParams(value: unknown): value is FormatPainterParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.sourceRange) && isRange(value.targetRange)
    && (value.continuous === undefined || typeof value.continuous === 'boolean');
}

function isValidRangeMoveParams(value: unknown): value is RangeMoveParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.sourceRange)
    && isRecord(value.targetOrigin) && isFiniteInt(value.targetOrigin.row) && isFiniteInt(value.targetOrigin.column)
    && isCellInputInterpretationContext(value.inputContext);
}

function isCellSnapshot(value: unknown): value is { row: number; column: number; cell: CellData } {
  return isRecord(value) && isFiniteInt(value.row) && isFiniteInt(value.column) && isRecord(value.cell) && 'value' in value.cell;
}

function isDataRegionMaterializeParams(value: unknown): value is DataRegionMaterializeParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || !isRecord(value.region)
    || typeof value.region.id !== 'string' || typeof value.region.sourceId !== 'string' || !isRange(value.region.range)
    || value.region.range.sheetId !== value.sheetId || !isRange(value.range) || value.range.sheetId !== value.sheetId
    || !Number.isSafeInteger(value.regionIndex) || Number(value.regionIndex) < 0
    || !isRecord(value.manifest) || value.manifest.schema !== 'DataSourceManifest' || value.manifest.version !== 1
    || typeof value.manifest.id !== 'string' || !Array.isArray(value.manifest.fields) || !Array.isArray(value.manifest.blocks)
    || typeof value.willRemoveSource !== 'boolean' || !Array.isArray(value.previousCells) || !value.previousCells.every(isCellSnapshot)
    || !Array.isArray(value.materializedCells) || !value.materializedCells.every(isCellSnapshot)) return false;
  return value.materializedCellCount === undefined || Number.isSafeInteger(value.materializedCellCount);
}

function dataRegionMaterializeAffected(params: DataRegionMaterializeParams): RangeRef[] {
  return [structuredClone(params.range)];
}

function applyDataRegionMaterialization(params: DataRegionMaterializeParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  const index = sheet.dataRegions.findIndex((region) => region.id === params.region.id);
  if (index < 0) throw new Error(`Unknown data region: ${params.region.id}`);
  const currentRegion = sheet.dataRegions[index]!;
  if (!rangeEquals(currentRegion.range, params.region.range)
    || currentRegion.sourceId !== params.region.sourceId
    || currentRegion.headerRow !== params.region.headerRow) {
    throw new Error(`Data region ${params.region.id} changed before materialization commit`);
  }
  const currentManifest = context.workbook.dataModel.sources.get(params.manifest.id);
  if (!currentManifest || currentManifest.revision !== params.manifest.revision) {
    throw new Error(`Data source ${params.manifest.id} changed before materialization commit`);
  }
  for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
    for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  for (const entry of params.materializedCells) sheet.cells.set(entry.row, entry.column, structuredClone(entry.cell));
  sheet.dataRegions.splice(index, 1);
  if (params.willRemoveSource) context.workbook.dataModel.sources.delete(params.manifest.id);
}

function restoreDataRegionMaterialization(params: DataRegionMaterializeParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  if (sheet.dataRegions.some((region) => region.id === params.region.id)) throw new Error(`Data region already exists: ${params.region.id}`);
  if (params.willRemoveSource) context.workbook.dataModel.sources.set(params.manifest.id, structuredClone(params.manifest));
  for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
    for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  for (const entry of params.previousCells) sheet.cells.set(entry.row, entry.column, structuredClone(entry.cell));
  sheet.dataRegions.splice(Math.min(params.regionIndex, sheet.dataRegions.length), 0, structuredClone(params.region));
}

function isValidFormatPainterMutation(value: unknown): value is FormatPainterMutationParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.targetRange)
    && Array.isArray(value.styles) && value.styles.every((row) => Array.isArray(row) && row.every((entry) => isRecord(entry)));
}

function formatPainterAffected(params: FormatPainterMutationParams): RangeRef[] {
  return rangeAffected(params.targetRange);
}

function applyFormatPainterMutation(params: FormatPainterMutationParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  const target = normalizeRange(params.targetRange, params.sheetId);
  for (let row = target.startRow; row <= target.endRow; row += 1) {
    for (let column = target.startColumn; column <= target.endColumn; column += 1) {
      const rowOffset = (row - target.startRow) % Math.max(1, params.styles.length);
      const sourceRow = params.styles[rowOffset] ?? [];
      const source = sourceRow[(column - target.startColumn) % Math.max(1, sourceRow.length)] ?? {};
      const current = sheet.cells.get(row, column) ?? { value: null };
      const next = { ...current };
      if (source.style && Object.keys(source.style).length > 0) next.style = structuredClone(source.style);
      else delete next.style;
      if (source.numberFormat !== undefined) next.numberFormat = source.numberFormat;
      else delete next.numberFormat;
      delete next.displayValue;
      sheet.cells.set(row, column, next);
    }
  }
}

function replaceText(original: string, params: ReplaceRangeParams): string | undefined {
  return replaceFindText(original, {
    query: params.find,
    matchCase: params.matchCase,
    entireCell: params.entireCell,
    wildcard: false,
  }, params.replace);
}

interface ReplaceCandidate {
  readonly text: string;
  readonly formula: boolean;
}

/**
 * Resolve the one value that Find/Replace is allowed to inspect. Formula
 * results come from the host resolver when available; displayValue is never a
 * source of search semantics. Formula text remains a separate candidate when
 * the caller explicitly searches formulas.
 */
function replaceCandidate(
  sheet: WorksheetModel,
  row: number,
  column: number,
  cell: CellData,
  searchIn: NonNullable<ReplaceRangeParams['searchIn']>,
  context: CommandContext,
): ReplaceCandidate | undefined {
  if (searchIn === 'formulas' || (searchIn === 'both' && cell.formula !== undefined)) {
    return cell.formula === undefined ? undefined : { text: cell.formula, formula: true };
  }
  const resolved = context.resolveCellValue?.(sheet, row, column);
  const value = resolved === undefined ? (cell.formulaValue ?? cell.value) : resolved;
  if (isFormulaError(value)) return { text: value.code, formula: false };
  if (value === null || value === undefined) return { text: '', formula: false };
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return { text: String(value), formula: false };
  }
  throw new Error(`Replace source at ${sheet.id}!${row}:${column} has an unsupported resolved value`);
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId && left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
}

function isFillMutationWrite(value: unknown): value is FillMutationWrite {
  if (!isRecord(value) || !isFiniteInt(value.row) || !isFiniteInt(value.column)) return false;
  const validCell = (cell: unknown): cell is CellData | undefined => cell === undefined
    || (isRecord(cell) && 'value' in cell);
  return validCell(value.before) && validCell(value.after);
}

function isValidFillMutationParams(value: unknown): value is FillMutationParams {
  if (!isRecord(value) || !isValidFillParams(value)) return false;
  const candidate = value as HomeFillParams & { writes?: unknown };
  return Array.isArray(candidate.writes)
    && candidate.writes.length > 0
    && candidate.writes.every(isFillMutationWrite);
}

function fillAffectedRanges(params: FillMutationParams): RangeRef[] {
  return [structuredClone(params.targetRange)];
}

function cellsEqual(left: CellData | undefined, right: CellData | undefined): boolean {
  if (left === undefined || right === undefined) return left === right;
  return JSON.stringify(left) === JSON.stringify(right);
}

function assertFillProtection(sheet: WorksheetModel, range: RangeRef): void {
  const decision = protectionResolver.resolve({
    sheetId: sheet.id,
    rules: sheet.protectionRules,
    ranges: [range],
    action: 'edit-cell',
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    readCellStyle: (row, column) => sheet.cells.get(row, column)?.style,
  });
  if (!decision.allowed) throw new Error(decision.reason ?? `Fill target ${sheet.id}!${range.startRow}:${range.startColumn}-${range.endRow}:${range.endColumn} is protected`);
}

function fillWritesEqual(expected: readonly FillWrite[], actual: readonly FillMutationWrite[]): boolean {
  if (expected.length !== actual.length) return false;
  return expected.every((entry, index) => {
    const candidate = actual[index];
    return candidate !== undefined && entry.row === candidate.row && entry.column === candidate.column
      && cellsEqual(entry.cell, candidate.after);
  });
}

function applyFillMutation(params: FillMutationParams, context: CommandContext, canonical: boolean): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  const target = normalizeRange(params.targetRange, params.sheetId);
  assertNoDataRegionIntersection(sheet, target, 'Fill');
  assertFillProtection(sheet, target);
  const planParams: FillPlanParams = {
    sheetId: params.sheetId,
    sourceRange: params.sourceRange,
    targetRange: params.targetRange,
    direction: params.direction,
    mode: params.mode,
    dateSystem: params.dateSystem,
  };
  const validated = validateFillPlan(sheet, planParams);
  if (canonical) {
    const plan = planFill(sheet, validated);
    if (!fillWritesEqual(plan.writes, params.writes)) throw new Error('Fill mutation does not match the canonical plan');
  }
  for (const write of params.writes) {
    if (write.row < target.startRow || write.row > target.endRow || write.column < target.startColumn || write.column > target.endColumn) {
      throw new Error('Fill mutation contains a write outside its target range');
    }
    const current = sheet.cells.get(write.row, write.column);
    if (!cellsEqual(current, write.before)) throw new Error(`Fill target changed at ${params.sheetId}!${write.row}:${write.column}`);
  }
  for (const write of params.writes) {
    if (write.after === undefined) sheet.cells.delete(write.row, write.column);
    else sheet.cells.set(write.row, write.column, structuredClone(write.after));
  }
}

/**
 * The sheet-features package intentionally has no dependency on the
 * application-owned block query/overlay service. Until that canonical service
 * is injected into CommandRuntime, Home commands fail closed for a
 * data-region intersection instead of treating CellMatrix as the block value.
 */
function assertNoDataRegionIntersection(sheet: WorksheetModel, range: RangeRef, operation: string): void {
  const region = sheet.dataRegions.find((candidate) => rangesIntersect(candidate.range, range));
  if (region) throw new Error(`${operation} does not support data-region ${region.id} without the canonical resolved-cell transaction`);
}

function cellsInRange(sheet: WorksheetModel, range: RangeRef): Array<{ row: number; column: number; cell?: CellData }> {
  const values: Array<{ row: number; column: number; cell?: CellData }> = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      values.push({ row, column, cell: sheet.cells.get(row, column) });
    }
  }
  return values;
}

function rangeArea(range: RangeRef): number {
  return (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1);
}

function rangeAreaOfCells(sheet: WorksheetModel, range: RangeRef): Array<{ row: number; column: number; cell?: CellData }> {
  return cellsInRange(sheet, range);
}

function setAppliedSortState(sheet: WorksheetModel, state: AppliedSortState | undefined): void {
  const target = sheet as SheetWithHomeState;
  if (state === undefined) delete target.appliedSortState;
  else target.appliedSortState = structuredClone(state);
}

function getAppliedSortState(sheet: WorksheetModel): AppliedSortState | undefined {
  const state = (sheet as SheetWithHomeState).appliedSortState;
  return state ? structuredClone(state) : undefined;
}

function hasFilterCriteria(filter: AutoFilterModel | undefined): boolean {
  return Boolean(filter && Object.values(filter.columns).some((entry) => Boolean(entry.criterion)));
}

function buildFilterFromParams(params: FilterToggleParams, sheet: WorksheetModel): AutoFilterModel {
  if (params.autoFilter) return normalizeAutoFilterModel(params.autoFilter);
  if (!params.range) throw new Error('Filter range is required when creating a filter');
  const range = normalizeRange(params.range, params.sheetId);
  const columns: AutoFilterModel['columns'] = {};
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    columns[column] = { column, showButton: true, hiddenButton: false };
  }
  return { sheetId: params.sheetId, range, columns };
}

function findPreset(id: string | HomeStylePreset): HomeStylePreset {
  if (typeof id !== 'string') return id;
  const preset = CELL_STYLE_PRESETS.find((entry) => entry.id === id || entry.name.toLocaleLowerCase() === id.toLocaleLowerCase());
  if (!preset) throw new Error(`Unknown cell style preset: ${id}`);
  return preset;
}

function isPreset(value: unknown): value is string | HomeStylePreset {
  return typeof value === 'string' || (isRecord(value) && typeof value.id === 'string' && isRecord(value.style));
}

function stylePresetMutationRanges(params: { sheetId: string; ranges: RangeRef[] }): RangeRef[] {
  return params.ranges.map((range) => structuredClone(range));
}

interface StylePresetMutationParams {
  sheetId: string;
  ranges: RangeRef[];
  styleId: string;
  style: Partial<CellStyle>;
}

function isStylePresetMutation(value: unknown): value is StylePresetMutationParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.styleId === 'string'
    && isRecord(value.style) && Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange);
}

function applyStylePresetMutation(params: StylePresetMutationParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  for (const range of params.ranges) {
    const normalized = normalizeRange(range, params.sheetId);
    for (let row = normalized.startRow; row <= normalized.endRow; row += 1) {
      for (let column = normalized.startColumn; column <= normalized.endColumn; column += 1) {
        const current = sheet.cells.get(row, column) ?? { value: null };
        const next = { ...current, styleId: params.styleId, style: structuredClone(params.style) };
        if (params.style.numberFormat !== undefined) next.numberFormat = params.style.numberFormat;
        delete next.displayValue;
        sheet.cells.set(row, column, next);
      }
    }
  }
}

interface ConditionalFormatReorderParams {
  sheetId: string;
  ruleIds: string[];
  ranges?: RangeRef[];
}

function isConditionalFormatReorder(value: unknown): value is ConditionalFormatReorderParams {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.ruleIds)
    && value.ruleIds.every((entry) => typeof entry === 'string')
    && (value.ranges === undefined || (Array.isArray(value.ranges) && value.ranges.every(isRange)));
}

function allConditionalRanges(sheet: WorksheetModel): RangeRef[] {
  return sheet.conditionalFormats.flatMap((rule) => rule.ranges.map((range) => structuredClone(range)));
}

interface TableStyleParams {
  sheetId: string;
  tableId: string;
  styleName?: string;
  showBandedRows?: boolean;
  showBandedColumns?: boolean;
}

interface TableStyleMutationParams extends TableStyleParams {
  range: RangeRef;
  clearStyleName?: boolean;
}

function isTableStyleParams(value: unknown): value is TableStyleParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.tableId === 'string'
    && (value.styleName === undefined || typeof value.styleName === 'string')
    && (value.showBandedRows === undefined || typeof value.showBandedRows === 'boolean')
    && (value.showBandedColumns === undefined || typeof value.showBandedColumns === 'boolean');
}

function isTableStyleMutationParams(value: unknown): value is TableStyleMutationParams {
  return isTableStyleParams(value) && isRecord(value) && isRange(value.range)
    && (value.clearStyleName === undefined || typeof value.clearStyleName === 'boolean');
}

interface DrawingVisibilityParams {
  sheetId: string;
  drawingId: string;
  visible: boolean;
}

function isDrawingVisibilityParams(value: unknown): value is DrawingVisibilityParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.drawingId === 'string' && typeof value.visible === 'boolean';
}

interface DrawingRenameParams {
  sheetId: string;
  drawingId: string;
  name: string;
}

interface DrawingRenameMutationParams extends DrawingRenameParams {}

function isDrawingRenameParams(value: unknown): value is DrawingRenameParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.drawingId === 'string'
    && typeof value.name === 'string' && value.name.trim().length > 0;
}

function isDrawingRenameMutation(value: unknown): value is DrawingRenameMutationParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.drawingId === 'string' && typeof value.name === 'string';
}

function drawingAffected(sheetId: string): RangeRef[] {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function tableStyleAffected(params: TableStyleMutationParams): RangeRef[] {
  return [structuredClone(params.range)];
}

export function registerHomeCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<FillMutationParams>({
    id: 'fill.applied',
    handler: (item, context) => {
      if (!isValidFillMutationParams(item.params)) throw new Error('Invalid fill.applied mutation payload');
      applyFillMutation(item.params, context, true);
    },
    metadata: {
      schema: { name: 'FillApplied', validate: isValidFillMutationParams },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: fillAffectedRanges, mode: 'exact' },
      inverseIds: ['fill.restored'],
    },
  });
  runtime.registry.registerMutation<FillMutationParams>({
    id: 'fill.restored',
    handler: (item, context) => {
      if (!isValidFillMutationParams(item.params)) throw new Error('Invalid fill.restored mutation payload');
      applyFillMutation(item.params, context, false);
    },
    metadata: {
      schema: { name: 'FillRestored', validate: isValidFillMutationParams },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: fillAffectedRanges, mode: 'exact' },
      inverseIds: ['fill.applied'],
    },
  });
  runtime.registry.registerMutation<DataRegionMaterializeParams>({
    id: 'dataRegion.materialize.commit',
    handler: (item, context) => {
      if (!isDataRegionMaterializeParams(item.params)) throw new Error('Invalid dataRegion.materialized mutation payload');
      applyDataRegionMaterialization(item.params, context);
    },
    metadata: {
      schema: { name: 'DataRegionMaterializeCommit', validate: isDataRegionMaterializeParams },
      permission: { capability: 'sheet.data-region.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: dataRegionMaterializeAffected, mode: 'exact' },
      inverseIds: ['dataRegion.materialize.restore'],
    },
  });
  runtime.registry.registerMutation<DataRegionMaterializeParams>({
    id: 'dataRegion.materialize.restore',
    handler: (item, context) => {
      if (!isDataRegionMaterializeParams(item.params)) throw new Error('Invalid dataRegion.restored mutation payload');
      restoreDataRegionMaterialization(item.params, context);
    },
    metadata: {
      schema: { name: 'DataRegionMaterializeRestore', validate: isDataRegionMaterializeParams },
      permission: { capability: 'sheet.data-region.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: dataRegionMaterializeAffected, mode: 'exact' },
      inverseIds: ['dataRegion.materialize.commit'],
    },
  });
  runtime.registry.registerCommand<DataRegionMaterializeParams>({
    id: 'dataRegion.materialize.commit',
    execute: (params, context) => {
      if (!isDataRegionMaterializeParams(params)) throw new Error('Invalid data-region materialize transaction');
      const affectedRanges = dataRegionMaterializeAffected(params);
      context.applyMutation({
        id: 'dataRegion.materialize.commit',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'dataRegion.materialize.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => applyDataRegionMaterialization(params, context),
      });
      return homeResult(context, affectedRanges, 1);
    },
  });

  runtime.registry.registerMutation<FormatPainterMutationParams>({
    id: 'format.painter.applied',
    handler: (item, context) => {
      if (!isValidFormatPainterMutation(item.params)) throw new Error('Invalid format.painter.applied mutation payload');
      applyFormatPainterMutation(item.params, context);
    },
    metadata: {
      schema: { name: 'FormatPainterApplied', validate: isValidFormatPainterMutation },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: formatPainterAffected, mode: 'exact' },
      inverseIds: ['cell.restore'],
    },
  });
  runtime.registry.registerCommand<FormatPainterParams>({
    id: 'format.painter.apply',
    execute: (params, context) => {
      if (!isValidFormatPainterParams(params)) throw new Error('Invalid format painter parameters');
      const source = normalizeRange(params.sourceRange, params.sheetId);
      const target = normalizeRange(params.targetRange, params.sheetId);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertNoDataRegionIntersection(sheet, source, 'Format painter');
      assertNoDataRegionIntersection(sheet, target, 'Format painter');
      const styles: FormatPainterCell[][] = [];
      for (let row = source.startRow; row <= source.endRow; row += 1) {
        const line: FormatPainterCell[] = [];
        for (let column = source.startColumn; column <= source.endColumn; column += 1) {
          const cell = sheet.cells.get(row, column);
          line.push({
            style: cell?.style ? structuredClone(cell.style) : undefined,
            numberFormat: cell?.numberFormat ?? cell?.style?.numberFormat,
          });
        }
        styles.push(line);
      }
      const previous = cellsInRange(sheet, target);
      const affectedRanges = rangeAffected(target);
      const inverse = previous.map(({ row, column, cell }) => ({
        id: 'cell.restore' as const,
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, row, column, previous: cell ? structuredClone(cell) : undefined },
        affectedRanges: [cellRange(params.sheetId, row, column)],
      }));
      context.applyMutation({
        id: 'format.painter.applied',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, targetRange: target, styles },
        affectedRanges,
        inverse,
        apply: () => applyFormatPainterMutation({ sheetId: params.sheetId, targetRange: target, styles }, context),
      });
      return homeResult(context, affectedRanges, 1);
    },
  });

  runtime.registry.registerCommand<RangeMoveParams>({
    id: 'sheet.range.move',
    execute: (params, context) => {
      if (!isValidRangeMoveParams(params)) throw new Error('Invalid range move parameters');
      const sourceRange = normalizeRange(params.sourceRange, params.sheetId);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertNoDataRegionIntersection(sheet, sourceRange, 'Range move');
      const targetRange: RangeRef = {
        sheetId: params.sheetId,
        startRow: params.targetOrigin.row,
        endRow: params.targetOrigin.row + sourceRange.endRow - sourceRange.startRow,
        startColumn: params.targetOrigin.column,
        endColumn: params.targetOrigin.column + sourceRange.endColumn - sourceRange.startColumn,
      };
      assertNoDataRegionIntersection(sheet, targetRange, 'Range move');
      const clipboard = copyRangeToClipboardData(context.workbook, sourceRange);
      clipboard.transfer = 'move';
      const targetEndRow = params.targetOrigin.row + sourceRange.endRow - sourceRange.startRow;
      const targetEndColumn = params.targetOrigin.column + sourceRange.endColumn - sourceRange.startColumn;
      if (targetEndRow >= sheet.rowCount || targetEndColumn >= sheet.columnCount) throw new Error('Range move exceeds worksheet bounds');
      return runtime.execute('sheet.range.paste', {
        sheetId: params.sheetId,
        targetOrigin: params.targetOrigin,
        clipboard,
        inputContext: params.inputContext,
        entryIntent: {
          kind: 'paste',
          target: targetRange,
          validationDecision: { status: 'not-applicable' },
        },
        transfer: 'move',
        spec: createPasteSpecialSpec(),
      });
    },
  });

  const mergeRange = (params: HomeRangeParams & { confirmDataLoss?: boolean }, context: CommandContext, center: boolean): CommandResult => {
    requireSheetId(params);
    const range = requireRange(params, params.sheetId);
    const sheet = context.workbook.getSheet(params.sheetId);
    assertNoDataRegionIntersection(sheet, range, 'Merge');
    const anchor = structuredClone(sheet.cells.get(range.startRow, range.startColumn));
    const hasNonAnchorContent = rangeAreaOfCells(sheet, range).some(({ row, column, cell }) =>
      (row !== range.startRow || column !== range.startColumn) && Boolean(cell && (cell.value !== null && cell.value !== undefined || cell.formula)));
    if (hasNonAnchorContent && params.confirmDataLoss === false) throw new Error('Merge would discard non-anchor cell contents');
    const result = runtime.execute('sheet.merge.set', { sheetId: params.sheetId, range });
    if (rangeArea(range) > 1) {
      runtime.execute('sheet.range.clear', { sheetId: params.sheetId, range, family: 'contents' });
      if (anchor) runtime.execute('sheet.cell.set', { sheetId: params.sheetId, row: range.startRow, column: range.startColumn, value: anchor, entryIntent: scriptEntryIntent(params.sheetId, range.startRow, range.startColumn, anchor) });
    }
    if (center) runtime.execute('sheet.style.set', { sheetId: params.sheetId, range, style: { horizontalAlignment: 'center' } });
    return result;
  };

  runtime.registry.registerCommand<HomeRangeParams & { confirmDataLoss?: boolean }>({
    id: 'sheet.merge.center',
    execute: (params, context) => mergeRange(params, context, true),
  });
  runtime.registry.registerCommand<HomeRangeParams & { confirmDataLoss?: boolean }>({
    id: 'sheet.merge.cells',
    execute: (params, context) => mergeRange(params, context, false),
  });
  runtime.registry.registerCommand<HomeRangeParams & { confirmDataLoss?: boolean }>({
    id: 'sheet.merge.across',
    execute: (params, context) => {
      requireSheetId(params);
      const range = requireRange(params, params.sheetId);
      let result = homeResult(context, []);
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        result = runtime.execute('sheet.merge.cells', {
          sheetId: params.sheetId,
          range: { ...range, startRow: row, endRow: row },
          confirmDataLoss: params.confirmDataLoss,
        });
      }
      return result;
    },
  });
  runtime.registry.registerCommand<HomeRangeParams>({
    id: 'sheet.merge.unmerge',
    execute: (params, context) => {
      requireSheetId(params);
      const range = requireRange(params, params.sheetId);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertNoDataRegionIntersection(sheet, range, 'Unmerge');
      const spans = sheet.merges.filter((merge) =>
        merge.range.startRow >= range.startRow && merge.range.endRow <= range.endRow
        && merge.range.startColumn >= range.startColumn && merge.range.endColumn <= range.endColumn);
      let mutationCount = 0;
      const affectedRanges: RangeRef[] = [];
      for (const span of spans) {
        const result = runtime.execute('sheet.merge.remove', { sheetId: params.sheetId, range: span.range });
        mutationCount += result.mutationCount;
        affectedRanges.push(...result.affectedRanges);
      }
      return homeResult(context, affectedRanges, mutationCount);
    },
  });
  runtime.registry.registerCommand<AutoSumParams>({
    id: 'formula.autosum',
    execute: (params, context) => {
      if (!isValidAutoSumParams(params)) throw new Error('Invalid AutoSum parameters');
      const range = normalizeRange(params.range, params.sheetId);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertNoDataRegionIntersection(sheet, range, 'AutoSum');
      const targets: Array<{ row: number; column: number }> = [];
      if (params.target) {
        targets.push({ row: params.target.row, column: params.target.column });
      } else if (params.byColumn || range.startColumn !== range.endColumn) {
        const targetRow = range.endRow + 1;
        for (let column = range.startColumn; column <= range.endColumn; column += 1) targets.push({ row: targetRow, column });
      } else {
        targets.push({ row: range.endRow + 1, column: range.startColumn });
      }
      if (targets.some(({ row, column }) => row >= sheet.rowCount || column >= sheet.columnCount)) throw new Error('AutoSum target is outside worksheet bounds');
      if (targets.some((target) => target.row >= range.startRow && target.row <= range.endRow
        && target.column >= range.startColumn && target.column <= range.endColumn)) {
        throw new Error('AutoSum source range would include its target');
      }
      if (!hasNumericResolvedValue(sheet, range, context)) throw new Error('AutoSum source contains no numeric result');
      for (const target of targets) {
        assertNoDataRegionIntersection(sheet, cellRange(params.sheetId, target.row, target.column), 'AutoSum');
        assertSafeAutoSumTarget(sheet, target);
      }
      const values: CellData[][] = [];
      const minRow = Math.min(...targets.map((target) => target.row));
      const maxRow = Math.max(...targets.map((target) => target.row));
      const minColumn = Math.min(...targets.map((target) => target.column));
      const maxColumn = Math.max(...targets.map((target) => target.column));
      for (let row = minRow; row <= maxRow; row += 1) {
        const line: CellData[] = [];
        for (let column = minColumn; column <= maxColumn; column += 1) {
          const target = targets.find((entry) => entry.row === row && entry.column === column);
          if (!target) {
            line.push(structuredClone(sheet.cells.get(row, column) ?? { value: null }));
            continue;
          }
          const previous = sheet.cells.get(row, column);
          line.push({
            ...(previous?.style ? { style: structuredClone(previous.style) } : {}),
            ...(previous?.numberFormat ? { numberFormat: previous.numberFormat } : {}),
            value: null,
            formula: formulaForAutoSum(sheet, row, column, params.functionName, range, context),
          });
        }
        values.push(line);
      }
      const result = runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: minRow,
        startColumn: minColumn,
        values,
        entryIntent: rangeEntryIntent('formula-result', params.sheetId, minRow, minColumn, values),
      });
      return result;
    },
  });

  runtime.registry.registerCommand<HomeFillParams>({
    id: 'sheet.range.fill',
    execute: (params, context) => {
      if (!isValidFillParams(params)) throw new Error('Invalid fill parameters');
      const source = normalizeRange(params.sourceRange, params.sheetId);
      const target = normalizeRange(params.targetRange, params.sheetId);
      if (!rangesIntersect(source, target)) throw new Error('Fill source and target must overlap');
      const sheet = context.workbook.getSheet(params.sheetId);
      assertNoDataRegionIntersection(sheet, source, 'Fill');
      assertNoDataRegionIntersection(sheet, target, 'Fill');
      assertFillProtection(sheet, target);
      const plan = planFill(sheet, {
        sheetId: params.sheetId,
        sourceRange: source,
        targetRange: target,
        direction: params.direction,
        mode: params.mode,
        dateSystem: params.dateSystem,
      });
      if (plan.writes.length === 0) return homeResult(context, rangeAffected(target));
      const writes: FillMutationWrite[] = plan.writes.map((write) => ({
        row: write.row,
        column: write.column,
        before: sheet.cells.get(write.row, write.column) ? structuredClone(sheet.cells.get(write.row, write.column)) : undefined,
        after: write.cell ? structuredClone(write.cell) : undefined,
      }));
      const inverseWrites: FillMutationWrite[] = writes.map((write) => ({
        row: write.row,
        column: write.column,
        before: write.after ? structuredClone(write.after) : undefined,
        after: write.before ? structuredClone(write.before) : undefined,
      }));
      const mutationParams: FillMutationParams = {
        sheetId: params.sheetId,
        sourceRange: source,
        targetRange: target,
        direction: params.direction,
        mode: params.mode,
        dateSystem: params.dateSystem,
        writes,
      };
      const inverseParams: FillMutationParams = { ...mutationParams, writes: inverseWrites };
      const affectedRanges = rangeAffected(target);
      context.applyMutation({
        id: 'fill.applied',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: mutationParams,
        affectedRanges,
        inverse: [{
          id: 'fill.restored',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: inverseParams,
          affectedRanges,
        }],
        apply: () => applyFillMutation(mutationParams, context, true),
      });
      return homeResult(context, affectedRanges, 1);
    },
  });

  runtime.registry.registerCommand<ReplaceRangeParams>({
    id: 'sheet.range.replace',
    execute: (params, context) => {
      if (!isValidReplaceParams(params)) throw new Error('Invalid replace parameters');
      if (!params.find) return homeResult(context, []);
      const sheets = params.scope === 'workbook' ? context.workbook.getSheets() : [context.workbook.getSheet(params.sheetId)];
      const defaultRange = params.range ? normalizeRange(params.range, params.sheetId) : undefined;
      const patches: Array<{ sheetId: string; row: number; column: number; next: CellData; previous?: CellData }> = [];
      const searchIn = params.searchIn ?? 'values';
      for (const sheet of sheets) {
        const range = defaultRange && sheet.id === params.sheetId
          ? defaultRange
          : { sheetId: sheet.id, startRow: 0, endRow: sheet.rowCount - 1, startColumn: 0, endColumn: sheet.columnCount - 1 };
        assertNoDataRegionIntersection(sheet, range, 'Replace');
        for (const { row, column, cell } of cellsInRange(sheet, range)) {
          if (!cell) continue;
          const candidate = replaceCandidate(sheet, row, column, cell, searchIn, context);
          if (!candidate) continue;
          const replaced = replaceText(candidate.text, params);
          if (replaced === undefined) continue;
          const replacement = parseReplacementValue(replaced, {
            ...params.inputContext,
            currentNumberFormat: cell.numberFormat ?? cell.style?.numberFormat,
            currentCellType: cell.editor?.kind,
          });
          if (replacement.kind === 'empty') throw new Error('Replacement text must not be empty');
          if (candidate.formula && replacement.kind !== 'formula') {
            throw new Error(`Formula replacement at ${sheet.id}!${row}:${column} must produce a formula`);
          }
          const next = replacementCell(cell, replacement);
          patches.push({ sheetId: sheet.id, row, column, next, previous: structuredClone(cell) });
        }
      }
      if (patches.length === 0) return homeResult(context, []);
      // Every nested range.set is part of this root command transaction, so
      // Replace All is one Undo even when matches span multiple sheets.
      for (const patch of patches) {
        const targetSheet = context.workbook.getSheet(patch.sheetId);
        const validation = patch.next.formula
          ? { valid: true, blocking: false, ruleId: undefined, alertStyle: undefined }
          : validateDataInput(targetSheet, patch.row, patch.column, patch.next.value);
        if (validation.blocking) throw new Error(validation.message ?? 'Find/Replace value failed data validation');
        if (!validation.valid) throw new Error('CELL_ENTRY_CONFIRMATION_REQUIRED: Find/Replace requires explicit validation confirmation');
        runtime.execute('sheet.cell.set', {
          sheetId: patch.sheetId,
          row: patch.row,
          column: patch.column,
          value: patch.next,
          entryIntent: {
            kind: 'direct-entry',
            target: { sheetId: patch.sheetId, row: patch.row, column: patch.column },
            candidate: structuredClone(patch.next),
            validationDecision: { status: 'accepted', ...(validation.ruleId ? { ruleId: validation.ruleId } : {}), ...(validation.alertStyle ? { alertStyle: validation.alertStyle } : {}) },
          },
        });
      }
      const affectedRanges = patches.map((patch) => cellRange(patch.sheetId, patch.row, patch.column));
      return homeResult(context, affectedRanges, patches.length);
    },
  });

  runtime.registry.registerCommand<FilterToggleParams>({
    id: 'sheet.autoFilter.toggle',
    execute: (params, context) => {
      if (!isValidFilterToggleParams(params)) throw new Error('Invalid filter toggle parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      if (params.range) assertNoDataRegionIntersection(sheet, normalizeRange(params.range, params.sheetId), 'Filter');
      const owner = resolveFilterOwner(sheet);
      if (owner?.kind === 'table') {
        const active = resolveActiveAutoFilter(sheet);
        if (active && (!params.range || rangeEquals(active.range, normalizeRange(params.range, params.sheetId)))) {
          return runtime.execute('sheetTable.autoFilter.set', { sheetId: params.sheetId, tableId: owner.tableId });
        }
        throw new Error('Use the Table AutoFilter owner for this range');
      }
      const current = sheet.autoFilter ? normalizeAutoFilterModel(sheet.autoFilter) : undefined;
      if (current) assertNoDataRegionIntersection(sheet, current.range, 'Filter');
      if (!current) {
        const next = buildFilterFromParams(params, sheet);
        return runtime.execute('sheet.autoFilter.set', { sheetId: params.sheetId, autoFilter: next });
      }
      const requestedRange = params.range ? normalizeRange(params.range, params.sheetId) : current.range;
      if (rangeEquals(current.range, requestedRange)) return runtime.execute('sheet.autoFilter.remove', { sheetId: params.sheetId });
      const next = params.autoFilter ? buildFilterFromParams(params, sheet) : { ...current, range: requestedRange };
      validateFilterOwnership(sheet, next, { kind: 'worksheet' });
      return runtime.execute('sheet.autoFilter.set', { sheetId: params.sheetId, autoFilter: next });
    },
  });

  runtime.registry.registerCommand<FilterCriteriaParams>({
    id: 'sheet.autoFilter.clearCriteria',
    execute: (params, context) => {
      if (!isValidFilterCriteriaParams(params)) throw new Error('Invalid filter clearCriteria parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      const currentSource = resolveActiveAutoFilter(sheet);
      const owner = resolveFilterOwner(sheet);
      if (!currentSource) return homeResult(context, []);
      const current = normalizeAutoFilterModel(currentSource);
      assertNoDataRegionIntersection(sheet, current.range, 'Filter');
      if (params.range && !rangesIntersect(current.range, normalizeRange(params.range, params.sheetId))) return homeResult(context, []);
      if (!hasFilterCriteria(current)) return homeResult(context, [current.range]);
      const cleared = { ...current, columns: Object.fromEntries(Object.entries(current.columns).map(([key, value]) => [key, { ...value, criterion: undefined }])) };
      return owner?.kind === 'table'
        ? runtime.execute('sheetTable.autoFilter.set', { sheetId: params.sheetId, tableId: owner.tableId, autoFilter: cleared })
        : runtime.execute('sheet.autoFilter.set', { sheetId: params.sheetId, autoFilter: cleared });
    },
  });

  runtime.registry.registerCommand<FilterCriteriaParams>({
    id: 'sheet.autoFilter.reapply',
    execute: (params, context) => {
      if (!isValidFilterCriteriaParams(params)) throw new Error('Invalid filter reapply parameters');
      const autoFilter = resolveActiveAutoFilter(context.workbook.getSheet(params.sheetId));
      if (autoFilter) assertNoDataRegionIntersection(context.workbook.getSheet(params.sheetId), autoFilter.range, 'Filter');
      return autoFilter ? homeResult(context, [structuredClone(autoFilter.range)]) : homeResult(context, []);
    },
  });

  runtime.registry.registerCommand<FilterSortParams>({
    id: 'sheet.autoFilter.sort',
    execute: (params, context) => {
      if (!isValidFilterSortParams(params)) throw new Error('Invalid AutoFilter sort parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      const filter = resolveActiveAutoFilter(sheet, params.column);
      const owner = resolveFilterOwner(sheet, params.column);
      if (!filter || !owner) throw new Error('No active AutoFilter in the current worksheet');
      if (params.column < filter.range.startColumn || params.column > filter.range.endColumn) throw new Error('Filter sort column is outside the filter range');
      const table = owner.kind === 'table' ? sheet.sheetTables.find((entry) => entry.id === owner.tableId) : undefined;
      const sortRange = table?.hasTotalRow ? { ...filter.range, endRow: filter.range.endRow - 1 } : filter.range;
      assertNoDataRegionIntersection(sheet, sortRange, 'Sort');
      const sortResult = runtime.execute('data.sort.rows', {
        sheetId: params.sheetId,
        range: sortRange,
        criteria: [{ column: params.column, ascending: params.ascending }],
        hasHeader: true,
      });
      const sortState = {
        ref: structuredClone(filter.range),
        conditions: [{
          ref: { ...structuredClone(sortRange), startColumn: params.column, endColumn: params.column },
          descending: !params.ascending,
        }],
      };
      const next = { ...filter, sortState };
      const filterResult = owner.kind === 'table'
        ? runtime.execute('sheetTable.autoFilter.set', { sheetId: params.sheetId, tableId: owner.tableId, autoFilter: next })
        : runtime.execute('sheet.autoFilter.set', { sheetId: params.sheetId, autoFilter: next });
      return {
        ...filterResult,
        mutationCount: sortResult.mutationCount + filterResult.mutationCount,
        affectedRanges: [...sortResult.affectedRanges, ...filterResult.affectedRanges],
      };
    },
  });

  runtime.registry.registerCommand<SortReapplyParams>({
    id: 'data.sort.reapply',
    execute: (params, context) => {
      if (!isValidSortReapplyParams(params)) throw new Error('Invalid sort reapply parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      const state = getAppliedSortState(sheet);
      if (!state) return homeResult(context, []);
      return runtime.execute('data.sort.rows', {
        sheetId: params.sheetId,
        range: state.range,
        criteria: state.criteria,
        hasHeader: state.hasHeader,
      });
    },
  });

  runtime.registry.registerCommand<HomeRangeParams & { sortColumn: number; ascending?: boolean; hasHeader?: boolean }>({
    id: 'data.sort.quick',
    execute: (params, context) => {
      requireSheetId(params);
      const range = requireRange(params, params.sheetId);
      if (!Number.isSafeInteger(params.sortColumn) || params.sortColumn < range.startColumn || params.sortColumn > range.endColumn) throw new Error('Sort column is outside range');
      const result = runtime.execute('data.sort.rows', {
        sheetId: params.sheetId,
        range,
        criteria: [{ column: params.sortColumn, ascending: params.ascending ?? true }],
        hasHeader: params.hasHeader,
      });
      const sheet = context.workbook.getSheet(params.sheetId) as SheetWithHomeState;
      sheet.appliedSortState = { sheetId: params.sheetId, range, criteria: [{ column: params.sortColumn, ascending: params.ascending ?? true }], hasHeader: params.hasHeader, revision: (sheet.sortRevision ?? 0) + 1 };
      return result;
    },
  });

  runtime.registry.registerCommand<GoToSpecialParams & { kind: GoToSpecialKind }>({
    id: 'selection.gotoSpecial',
    execute: (params, context) => {
      if (!isRecord(params) || typeof params.sheetId !== 'string' || !isRange(params.range)) throw new Error('Invalid Go To Special parameters');
      const range = normalizeRange(params.range, params.sheetId);
      assertNoDataRegionIntersection(context.workbook.getSheet(params.sheetId), range, 'Go To Special');
      const resolved = resolveGoToSpecial(context.workbook, { sheetId: params.sheetId, range, kind: params.kind });
      return homeResult(context, resolved);
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; reference: string }>({
    id: 'selection.goto',
    execute: (params, context) => {
      if (!isRecord(params) || typeof params.sheetId !== 'string' || typeof params.reference !== 'string') throw new Error('Invalid Go To parameters');
      const target = resolveGoToRange(context.workbook, params);
      return target ? homeResult(context, [target]) : homeResult(context, []);
    },
  });

  runtime.registry.registerMutation<StylePresetMutationParams>({
    id: 'style.preset.set',
    handler: (item, context) => {
      if (!isStylePresetMutation(item.params)) throw new Error('Invalid style.preset.set mutation payload');
      applyStylePresetMutation(item.params, context);
    },
    metadata: {
      schema: { name: 'CellStylePreset', validate: isStylePresetMutation },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: stylePresetMutationRanges, mode: 'exact' },
      inverseIds: ['cell.restore'],
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ranges: RangeRef[]; preset: string | HomeStylePreset }>({
    id: 'sheet.style.preset.apply',
    execute: (params, context) => {
      if (!isRecord(params) || typeof params.sheetId !== 'string' || !Array.isArray(params.ranges) || params.ranges.length === 0 || !params.ranges.every(isRange) || !isPreset(params.preset)) throw new Error('Invalid cell style preset parameters');
      const ranges = params.ranges.map((range) => normalizeRange(range, params.sheetId));
      const preset = findPreset(params.preset);
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = ranges.flatMap((range) => cellsInRange(sheet, range));
      ranges.forEach((range) => assertNoDataRegionIntersection(sheet, range, 'Cell style preset'));
      const affectedRanges = ranges.map((range) => structuredClone(range));
      const inverse = previous.map(({ row, column, cell }) => ({
        id: 'cell.restore' as const,
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, row, column, previous: cell ? structuredClone(cell) : undefined },
        affectedRanges: [cellRange(params.sheetId, row, column)],
      }));
      const mutationParams = { sheetId: params.sheetId, ranges, styleId: preset.id, style: structuredClone(preset.style) };
      context.applyMutation({
        id: 'style.preset.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: mutationParams,
        affectedRanges,
        inverse,
        apply: () => applyStylePresetMutation(mutationParams, context),
      });
      return homeResult(context, affectedRanges, 1);
    },
  });

  runtime.registry.registerMutation<TableStyleMutationParams>({
    id: 'sheetTable.style.set',
    handler: (item, context) => {
      if (!isTableStyleMutationParams(item.params)) throw new Error('Invalid table style mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const index = sheet.sheetTables.findIndex((table) => table.id === item.params.tableId);
      if (index < 0) throw new Error(`Unknown sheet table: ${item.params.tableId}`);
      const table = sheet.sheetTables[index]!;
      sheet.sheetTables[index] = {
        ...table,
        ...(item.params.clearStyleName ? { styleName: undefined } : item.params.styleName === undefined ? {} : { styleName: item.params.styleName }),
        ...(item.params.showBandedRows === undefined ? {} : { showBandedRows: item.params.showBandedRows }),
        ...(item.params.showBandedColumns === undefined ? {} : { showBandedColumns: item.params.showBandedColumns }),
      };
    },
    metadata: {
      schema: { name: 'SheetTableStyleSet', validate: isTableStyleMutationParams },
      permission: { capability: 'sheet.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: tableStyleAffected, mode: 'exact' },
      inverseIds: ['sheetTable.style.set'],
    },
  });
  runtime.registry.registerCommand<TableStyleParams>({
    id: 'sheetTable.style.set',
    execute: (params, context) => {
      if (!isTableStyleParams(params)) throw new Error('Invalid table style parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      const table = sheet.sheetTables.find((entry) => entry.id === params.tableId);
      if (!table) throw new Error(`Unknown sheet table: ${params.tableId}`);
      const previous: TableStyleMutationParams = {
        sheetId: params.sheetId,
        tableId: params.tableId,
        range: structuredClone(table.range),
        ...(table.styleName === undefined ? { clearStyleName: true } : { styleName: table.styleName }),
        showBandedRows: table.showBandedRows,
        showBandedColumns: table.showBandedColumns,
      };
      const next: TableStyleMutationParams = {
        ...previous,
        clearStyleName: false,
        ...(params.styleName === undefined ? {} : { styleName: params.styleName }),
        ...(params.showBandedRows === undefined ? {} : { showBandedRows: params.showBandedRows }),
        ...(params.showBandedColumns === undefined ? {} : { showBandedColumns: params.showBandedColumns }),
      };
      const affectedRanges = tableStyleAffected(next);
      context.applyMutation({
        id: 'sheetTable.style.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: next,
        affectedRanges,
        inverse: [{ id: 'sheetTable.style.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => {
          const index = sheet.sheetTables.findIndex((entry) => entry.id === params.tableId);
          if (index < 0) throw new Error(`Unknown sheet table: ${params.tableId}`);
          sheet.sheetTables[index] = {
            ...sheet.sheetTables[index]!,
            ...(next.clearStyleName ? { styleName: undefined } : next.styleName === undefined ? {} : { styleName: next.styleName }),
            ...(next.showBandedRows === undefined ? {} : { showBandedRows: next.showBandedRows }),
            ...(next.showBandedColumns === undefined ? {} : { showBandedColumns: next.showBandedColumns }),
          };
        },
      });
      return homeResult(context, affectedRanges, 1);
    },
  });
  runtime.registry.registerMutation<ConditionalFormatReorderParams>({
    id: 'cf.reorder',
    handler: (item, context) => {
      if (!isConditionalFormatReorder(item.params)) throw new Error('Invalid cf.reorder mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const byId = new Map(sheet.conditionalFormats.map((rule) => [rule.id, rule] as const));
      const next = item.params.ruleIds.map((id) => byId.get(id)).filter((rule): rule is ConditionalFormatRule => rule !== undefined);
      for (const rule of sheet.conditionalFormats) if (!item.params.ruleIds.includes(rule.id)) next.push(rule);
      sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...next.map((rule, index) => ({ ...structuredClone(rule), priority: index + 1 })));
    },
    metadata: {
      schema: { name: 'ConditionalFormatReorder', validate: isConditionalFormatReorder },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.ranges?.map((range) => structuredClone(range)) ?? [], mode: 'declared' },
      inverseIds: ['cf.reorder'],
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleIds: string[] }>({
    id: 'conditionalFormat.reorder',
    execute: (params, context) => {
      if (!isConditionalFormatReorder(params)) throw new Error('Invalid conditional format reorder parameters');
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousIds = sheet.conditionalFormats.map((rule) => rule.id);
      const affectedRanges = allConditionalRanges(sheet);
      context.applyMutation({
        id: 'cf.reorder',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: affectedRanges },
        affectedRanges,
        inverse: [{ id: 'cf.reorder', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, ruleIds: previousIds, ranges: affectedRanges }, affectedRanges }],
        apply: () => {
          const byId = new Map(sheet.conditionalFormats.map((rule) => [rule.id, rule] as const));
          const next = params.ruleIds.map((id) => byId.get(id)).filter((rule): rule is ConditionalFormatRule => rule !== undefined);
          for (const rule of sheet.conditionalFormats) if (!params.ruleIds.includes(rule.id)) next.push(rule);
          sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...next);
        },
      });
      return homeResult(context, affectedRanges, 1);
    },
  });

  runtime.registry.registerMutation<DrawingVisibilityParams>({
    id: 'drawing.visibility.set',
    handler: (item, context) => {
      if (!isDrawingVisibilityParams(item.params)) throw new Error('Invalid drawing.visibility.set mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const drawing = sheet.drawings.find((entry) => entry.id === item.params.drawingId) as DrawingWithHomeState | undefined;
      if (!drawing) throw new Error(`Unknown drawing: ${item.params.drawingId}`);
      drawing.visible = item.params.visible;
    },
    metadata: {
      schema: { name: 'DrawingVisibilitySet', validate: isDrawingVisibilityParams },
      permission: { capability: 'drawing.edit', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => drawingAffected(params.sheetId), mode: 'exact' },
      inverseIds: ['drawing.visibility.set'],
    },
  });
  runtime.registry.registerCommand<DrawingVisibilityParams>({
    id: 'drawing.visibility.set',
    execute: (params, context) => {
      if (!isDrawingVisibilityParams(params)) throw new Error('Invalid drawing visibility parameters');
      const drawing = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId) as DrawingWithHomeState | undefined;
      if (!drawing) throw new Error(`Unknown drawing: ${params.drawingId}`);
      const previous = drawing.visible ?? true;
      const affectedRanges = drawingAffected(params.sheetId);
      context.applyMutation({
        id: 'drawing.visibility.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'drawing.visibility.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, visible: previous }, affectedRanges }],
        apply: () => { drawing.visible = params.visible; },
      });
      return homeResult(context, affectedRanges, 1);
    },
  });

  runtime.registry.registerMutation<DrawingRenameMutationParams>({
    id: 'drawing.rename',
    handler: (item, context) => {
      if (!isDrawingRenameMutation(item.params)) throw new Error('Invalid drawing.rename mutation payload');
      const drawing = context.workbook.getSheet(item.params.sheetId).drawings.find((entry) => entry.id === item.params.drawingId);
      if (!drawing) throw new Error(`Unknown drawing: ${item.params.drawingId}`);
      if (item.params.name.trim()) drawing.name = item.params.name;
      else delete drawing.name;
    },
    metadata: {
      schema: { name: 'DrawingRename', validate: isDrawingRenameMutation },
      permission: { capability: 'drawing.edit', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => drawingAffected(params.sheetId), mode: 'exact' },
      inverseIds: ['drawing.rename'],
    },
  });
  runtime.registry.registerCommand<DrawingRenameParams>({
    id: 'drawing.rename',
    execute: (params, context) => {
      if (!isDrawingRenameParams(params)) throw new Error('Invalid drawing rename parameters');
      const drawing = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId);
      if (!drawing) throw new Error(`Unknown drawing: ${params.drawingId}`);
      const previous = drawing.name;
      const affectedRanges = drawingAffected(params.sheetId);
      context.applyMutation({
        id: 'drawing.rename',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'drawing.rename', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, name: previous ?? '' }, affectedRanges }],
        apply: () => { drawing.name = params.name; },
      });
      return homeResult(context, affectedRanges, 1);
    },
  });
}

export function buildClipboardPayload(workbook: WorkbookModel, range: RangeRef): ClipboardPayload {
  return copyRangeToClipboardData(workbook, range);
}
