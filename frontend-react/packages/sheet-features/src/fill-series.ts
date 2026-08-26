import { clearFormulaProvenance, type CellData, type RangeRef, type WorksheetModel } from '@react-sheets/core-model';
import { canonicalExcelDateFromSerial, canonicalExcelDateFromValue, canonicalExcelDateToIso, type ExcelDateSystem } from '@react-sheets/formula-engine';
import { shiftFormula } from './clipboard';

/** The only directions understood by the fill planner. */
export type FillDirection = 'down' | 'up' | 'right' | 'left';

/** Copy and numeric-series are deliberately separate semantics. */
export type FillMode = 'copy' | 'series';

export interface FillPlanParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetRange: RangeRef;
  direction: FillDirection;
  mode: FillMode;
  /** Workbook calendar used when a date-formatted seed is extended. */
  dateSystem?: ExcelDateSystem;
}

export interface FillWrite {
  readonly row: number;
  readonly column: number;
  /** `undefined` means that the target cell must be deleted. */
  readonly cell?: CellData;
}

export interface FillPlan {
  readonly sheetId: string;
  readonly sourceRange: RangeRef;
  readonly targetRange: RangeRef;
  readonly direction: FillDirection;
  readonly mode: FillMode;
  readonly dateSystem: ExcelDateSystem;
  /** Only actual changes are emitted; seed cells are never rewritten. */
  readonly writes: readonly FillWrite[];
}

/** Keep client planning bounded to the same reducer limit as the server. */
export const MAX_FILL_CELLS = 100_000;

function isFiniteInt(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function isFillDirection(value: unknown): value is FillDirection {
  return value === 'down' || value === 'up' || value === 'right' || value === 'left';
}

function isFillMode(value: unknown): value is FillMode {
  return value === 'copy' || value === 'series';
}

function isRange(value: unknown): value is RangeRef {
  if (!value || typeof value !== 'object') return false;
  const range = value as Record<string, unknown>;
  return typeof range.sheetId === 'string'
    && isFiniteInt(range.startRow) && isFiniteInt(range.endRow)
    && isFiniteInt(range.startColumn) && isFiniteInt(range.endColumn);
}

function normalizeRange(range: RangeRef, sheetId: string, name: string): RangeRef {
  if (!isRange(range) || range.sheetId !== sheetId) throw new Error(`Fill ${name} must target the command sheet`);
  return {
    sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function rangeHeight(range: RangeRef): number {
  return range.endRow - range.startRow + 1;
}

function rangeWidth(range: RangeRef): number {
  return range.endColumn - range.startColumn + 1;
}

function rangeContains(outer: RangeRef, inner: RangeRef): boolean {
  return outer.sheetId === inner.sheetId
    && outer.startRow <= inner.startRow && outer.endRow >= inner.endRow
    && outer.startColumn <= inner.startColumn && outer.endColumn >= inner.endColumn;
}

function cellsEqual(left: CellData | undefined, right: CellData | undefined): boolean {
  if (left === undefined || right === undefined) return left === right;
  return JSON.stringify(left) === JSON.stringify(right);
}

function cellIsBlank(cell: CellData | undefined): boolean {
  return cell === undefined || (cell.formula === undefined && cell.formulaValue === undefined && cell.value === null);
}

function assertBounds(sheet: WorksheetModel, range: RangeRef, name: string): void {
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) {
    throw new Error(`Fill ${name} is outside worksheet bounds`);
  }
  if (rangeHeight(range) * rangeWidth(range) > MAX_FILL_CELLS) throw new Error(`Fill ${name} is too large`);
}

/**
 * A fill is a one-axis extension. This rejects diagonal or disjoint payloads
 * before any cell is read for execution, so every caller shares one geometry.
 */
function assertGeometry(source: RangeRef, target: RangeRef, direction: FillDirection): void {
  if (!rangeContains(target, source)) throw new Error('Fill target must contain the source range');
  const sameColumns = source.startColumn === target.startColumn && source.endColumn === target.endColumn;
  const sameRows = source.startRow === target.startRow && source.endRow === target.endRow;
  if ((direction === 'down' && (!sameColumns || target.startRow !== source.startRow || target.endRow < source.endRow))
    || (direction === 'up' && (!sameColumns || target.endRow !== source.endRow || target.startRow > source.startRow))
    || (direction === 'right' && (!sameRows || target.startColumn !== source.startColumn || target.endColumn < source.endColumn))
    || (direction === 'left' && (!sameRows || target.endColumn !== source.endColumn || target.startColumn > source.startColumn))) {
    throw new Error(`Fill ${direction} requires a contiguous one-axis target extension`);
  }
}

function sourceCell(sheet: WorksheetModel, row: number, column: number): CellData | undefined {
  const cell = sheet.cells.get(row, column);
  return cell === undefined ? undefined : structuredClone(cell);
}

function copyCell(source: CellData | undefined, rowDelta: number, columnDelta: number): CellData | undefined {
  if (source === undefined) return undefined;
  const copy = clearFormulaProvenance(source);
  if (copy.formula !== undefined) {
    copy.formula = shiftFormula(copy.formula, rowDelta, columnDelta);
    copy.value = null;
    delete copy.formulaValue;
    delete copy.displayValue;
  }
  return copy;
}

function copySourceCoordinate(source: RangeRef, row: number, column: number, direction: FillDirection): { row: number; column: number } {
  const sourceHeight = rangeHeight(source);
  const sourceWidth = rangeWidth(source);
  const sourceRow = direction === 'up'
    ? source.endRow - ((source.endRow - row) % sourceHeight)
    : source.startRow + ((row - source.startRow) % sourceHeight);
  const sourceColumn = direction === 'left'
    ? source.endColumn - ((source.endColumn - column) % sourceWidth)
    : source.startColumn + ((column - source.startColumn) % sourceWidth);
  return { row: sourceRow, column: sourceColumn };
}

function planCopy(sheet: WorksheetModel, source: RangeRef, target: RangeRef, direction: FillDirection, dateSystem: ExcelDateSystem): FillPlan {
  const writes: FillWrite[] = [];
  for (let row = target.startRow; row <= target.endRow; row += 1) {
    for (let column = target.startColumn; column <= target.endColumn; column += 1) {
      const sourceCoordinate = copySourceCoordinate(source, row, column, direction);
      const current = sourceCell(sheet, row, column);
      const copied = copyCell(
        sourceCell(sheet, sourceCoordinate.row, sourceCoordinate.column),
        row - sourceCoordinate.row,
        column - sourceCoordinate.column,
      );
      if (!cellsEqual(current, copied)) writes.push({ row, column, cell: copied });
    }
  }
  return { sheetId: sheet.id, sourceRange: source, targetRange: target, direction, mode: 'copy', dateSystem, writes };
}

interface SeriesSeed {
  readonly row: number;
  readonly column: number;
  readonly value: number;
  readonly cell: CellData;
  readonly travel: number;
  readonly kind: 'number' | 'date';
}

function isDateFormat(format: string | undefined): boolean {
  if (!format) return false;
  const unquoted = format.replace(/"(?:[^"]|"")*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  return /[ydhms]/i.test(unquoted);
}

function travelCoordinate(direction: FillDirection, row: number, column: number): number {
  const coordinate = direction === 'down' || direction === 'up' ? row : column;
  return direction === 'up' || direction === 'left' ? -coordinate : coordinate;
}

function seriesTrackKey(direction: FillDirection, row: number, column: number): number {
  return direction === 'down' || direction === 'up' ? column : row;
}

function seriesSeeds(
  sheet: WorksheetModel,
  source: RangeRef,
  direction: FillDirection,
  dateSystem: ExcelDateSystem,
): Map<number, SeriesSeed[]> {
  const tracks = new Map<number, SeriesSeed[]>();
  for (let row = source.startRow; row <= source.endRow; row += 1) {
    for (let column = source.startColumn; column <= source.endColumn; column += 1) {
      const cell = sourceCell(sheet, row, column);
      if (cellIsBlank(cell)) continue;
      if (cell?.formula !== undefined || cell?.formulaValue !== undefined) {
        throw new Error('Series fill accepts only finite numeric seeds; use copy mode for formulas and text');
      }
      const date = isDateFormat(cell?.numberFormat) ? canonicalExcelDateFromValue(cell?.value, dateSystem) : null;
      const numericValue = date?.serial ?? (typeof cell?.value === 'number' && Number.isFinite(cell.value) ? cell.value : null);
      if (numericValue === null) throw new Error('Series fill accepts only finite numeric seeds or canonical date seeds; use copy mode for formulas and text');
      const key = seriesTrackKey(direction, row, column);
      if (!cell) throw new Error('Series fill seed disappeared during planning');
      const seed: SeriesSeed = { row, column, value: numericValue, cell, travel: travelCoordinate(direction, row, column), kind: date ? 'date' : 'number' };
      const entries = tracks.get(key) ?? [];
      entries.push(seed);
      tracks.set(key, entries);
    }
  }
  for (const entries of tracks.values()) {
    entries.sort((left, right) => left.travel - right.travel);
    if (entries.some((entry) => entry.kind !== entries[0]!.kind)) throw new Error('Series seeds must use one value kind per track');
  }
  return tracks;
}

function seriesStep(seeds: readonly SeriesSeed[]): number {
  if (seeds.length < 2) return 1;
  const first = seeds[0]!;
  const second = seeds[1]!;
  const distance = second.travel - first.travel;
  if (distance <= 0) throw new Error('Series seeds must have a strict axis order');
  const step = (second.value - first.value) / distance;
  if (!Number.isFinite(step)) throw new Error('Series step is not finite');
  for (let index = 2; index < seeds.length; index += 1) {
    const seed = seeds[index]!;
    const expected = first.value + step * (seed.travel - first.travel);
    const tolerance = Number.EPSILON * Math.max(1, Math.abs(expected), Math.abs(seed.value)) * 16;
    if (Math.abs(expected - seed.value) > tolerance) throw new Error('Series seeds do not define one linear progression');
  }
  return step;
}

function planSeries(sheet: WorksheetModel, source: RangeRef, target: RangeRef, direction: FillDirection, dateSystem: ExcelDateSystem): FillPlan {
  const tracks = seriesSeeds(sheet, source, direction, dateSystem);
  const writes: FillWrite[] = [];
  const seedCoordinates = new Set<string>();
  for (const entries of tracks.values()) for (const seed of entries) seedCoordinates.add(`${seed.row}:${seed.column}`);

  for (let row = target.startRow; row <= target.endRow; row += 1) {
    for (let column = target.startColumn; column <= target.endColumn; column += 1) {
      if (seedCoordinates.has(`${row}:${column}`)) continue;
      const key = seriesTrackKey(direction, row, column);
      const seeds = tracks.get(key);
      if (!seeds || seeds.length === 0) throw new Error('Series fill requires a numeric seed on every affected track');
      const first = seeds[0]!;
      const value = first.value + seriesStep(seeds) * (travelCoordinate(direction, row, column) - first.travel);
      if (!Number.isFinite(value)) throw new Error('Series fill would produce a non-finite number');
      const next = structuredClone(first.cell);
      next.value = first.kind === 'date' ? canonicalExcelDateToIso(canonicalExcelDateFromSerial(value, dateSystem)) : value;
      delete next.formula;
      delete next.formulaValue;
      delete next.displayValue;
      delete next.formulaMetadata;
      const current = sourceCell(sheet, row, column);
      if (!cellsEqual(current, next)) writes.push({ row, column, cell: next });
    }
  }
  if (writes.length === 0 && tracks.size === 0) throw new Error('Series fill requires at least one numeric seed');
  return { sheetId: sheet.id, sourceRange: source, targetRange: target, direction, mode: 'series', dateSystem, writes };
}

/** Validate and normalize the shape shared by command and replay. */
export function validateFillPlan(sheet: WorksheetModel, params: FillPlanParams): FillPlanParams {
  if (params.sheetId !== sheet.id) throw new Error('Fill sheet mismatch');
  if (!isFillDirection(params.direction) || !isFillMode(params.mode)) throw new Error('Invalid fill mode or direction');
  const source = normalizeRange(params.sourceRange, params.sheetId, 'source range');
  const target = normalizeRange(params.targetRange, params.sheetId, 'target range');
  assertBounds(sheet, source, 'source range');
  assertBounds(sheet, target, 'target range');
  assertGeometry(source, target, params.direction);
  const dateSystem = params.dateSystem ?? '1900';
  if (dateSystem !== '1900' && dateSystem !== '1904') throw new Error('Fill dateSystem must be 1900 or 1904');
  return { ...params, sourceRange: source, targetRange: target, dateSystem };
}

/** Build the canonical, side-effect-free plan shared by command and replay. */
export function planFill(sheet: WorksheetModel, params: FillPlanParams): FillPlan {
  const validated = validateFillPlan(sheet, params);
  const source = validated.sourceRange;
  const target = validated.targetRange;
  const dateSystem = validated.dateSystem ?? '1900';
  return validated.mode === 'copy'
    ? planCopy(sheet, source, target, validated.direction, dateSystem)
    : planSeries(sheet, source, target, validated.direction, dateSystem);
}
