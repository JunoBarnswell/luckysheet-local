import type {
  CellData,
  FilterCellValue,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
  AutoFilterColumn,
  AutoFilterModel,
  DateGroupItem,
  FilterCriterion,
  FilterScalar,
  RangeRef,
  WorksheetModel,
} from "@react-sheets/core-model";
import { clearFormulaProvenance, StructuralTransform, applyRowPermutation, columnLabel, createRowPermutationPlan, isDynamicFilterType, resolveFilterCellValue, sheetRuleRegistry } from "@react-sheets/core-model";
import { canonicalExcelDateDayOfWeek, canonicalExcelDateFromParts, canonicalExcelDateFromUtcDate, canonicalExcelDateFromValue, canonicalExcelDateToUtcDate, shiftCanonicalExcelDate, type CanonicalExcelDate, type CanonicalExcelDateParts } from '@react-sheets/formula-engine';
import { compareWorkbookValues } from '@react-sheets/formula-engine';
import { resolveAutoFilters } from './sheet-table-features';
import { assertDataRegionContextMatches, resolveDataRegionContext, type DataRegionContext } from './data-region-context';
import type { CommandContext, CommandRuntime } from "@react-sheets/command-runtime";
import {
  evaluateFormula,
  formatFormula,
  isArrayValue,
  isFormulaError,
  mapAstReferences,
  offsetAst,
  parseFormula,
  type ParsedCellReference,
  type FormulaError,
  type FormulaAst,
  type FormulaValue,
  type ScalarValue,
} from "@react-sheets/formula-engine";

// ---------- 基础 ----------

export function normalizeRangeRef(range: RangeRef): RangeRef {
  return {
    sheetId: range.sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

export interface AppliedSortState {
  sheetId: string;
  range: RangeRef;
  criteria: Array<{ column: number; ascending: boolean }>;
  hasHeader?: boolean;
  revision: number;
}

interface RowsPermutedMutationParams {
  sheetId: string;
  range: RangeRef;
  sourceRows: number[];
  affectedColumnEnd: number;
  sortState?: AppliedSortState;
  previousSortState?: AppliedSortState;
}

/**
 * A row permutation remaps row-addressed metadata for every valid worksheet
 * column, including protection rules outside the materialized grid width.
 * This is the frontend half of the persisted rows.permuted contract.
 */
function rowsPermutedAffectedRanges(params: RowsPermutedMutationParams): RangeRef[] {
  return [{
    sheetId: params.sheetId,
    startRow: params.range.startRow,
    endRow: params.range.endRow,
    startColumn: 0,
    endColumn: params.affectedColumnEnd,
  }];
}

function isRowsPermutedMutation(value: unknown): value is RowsPermutedMutationParams {
  if (!value || typeof value !== 'object') return false;
  const params = value as Record<string, unknown>;
  const range = params.range;
  if (!range || typeof range !== 'object') return false;
  const candidate = range as Record<string, unknown>;
  return typeof params.sheetId === 'string'
    && candidate.sheetId === params.sheetId
    && Number.isInteger(candidate.startRow) && Number.isInteger(candidate.endRow)
    && Number.isInteger(candidate.startColumn) && Number.isInteger(candidate.endColumn)
    && Number(candidate.startRow) >= 0 && Number(candidate.endRow) >= Number(candidate.startRow)
    && Number(candidate.startColumn) >= 0 && Number(candidate.endColumn) >= Number(candidate.startColumn)
    && Array.isArray(params.sourceRows)
    && Number.isInteger(params.affectedColumnEnd)
    && Number(params.affectedColumnEnd) >= Number(candidate.endColumn)
    && params.sourceRows.length === Number(candidate.endRow) - Number(candidate.startRow) + 1
    && new Set(params.sourceRows).size === params.sourceRows.length
    && params.sourceRows.every((row) => Number.isInteger(row) && Number(row) >= Number(candidate.startRow) && Number(row) <= Number(candidate.endRow));
}

function setAppliedSortState(sheet: WorksheetModel, state: AppliedSortState | undefined): void {
  const target = sheet as WorksheetModel & { appliedSortState?: AppliedSortState };
  if (state === undefined) delete target.appliedSortState;
  else target.appliedSortState = structuredClone(state);
}

function rowsPermutedAffectedColumnEnd(sheet: WorksheetModel, range: RangeRef): number {
  return sheetRuleRegistry.affectedColumnEnd(sheet, range.endColumn);
}

function inRange(range: RangeRef, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow
    && column >= range.startColumn && column <= range.endColumn;
}

function cellRange(sheetId: string, row: number, column: number): RangeRef {
  return { sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column };
}

function snapshotCells(sheet: WorksheetModel, range: RangeRef): Array<{ row: number; column: number; previous?: CellData }> {
  const result: Array<{ row: number; column: number; previous?: CellData }> = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      result.push({ row, column, previous: structuredClone(sheet.cells.get(row, column)) });
    }
  }
  return result;
}

function applyRangeValues(
  context: CommandContext,
  params: { sheetId: string; startRow: number; startColumn: number; values: CellData[][] },
): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  const range: RangeRef = {
    sheetId: params.sheetId,
    startRow: params.startRow,
    endRow: params.startRow + Math.max(0, params.values.length - 1),
    startColumn: params.startColumn,
    endColumn: params.startColumn + Math.max(0, Math.max(0, ...params.values.map((line) => line.length)) - 1),
  };
  const values = params.values.map((row) => row.map((value) => value ? clearFormulaProvenance(value) : value));
  const previous = snapshotCells(sheet, range);
  const affectedRanges = [range];
  context.applyMutation({
    id: 'range.set',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: {
      ...params,
      values,
      entryIntent: {
        kind: 'script' as const,
        target: {
          sheetId: params.sheetId,
          startRow: params.startRow,
          endRow: params.startRow + Math.max(0, values.length - 1),
          startColumn: params.startColumn,
          endColumn: params.startColumn + Math.max(0, Math.max(1, ...values.map((row) => row.length)) - 1),
        },
        candidate: structuredClone(values),
        validationDecision: { status: 'not-applicable' as const },
      },
    },
    affectedRanges,
    inverse: previous.map((entry) => ({
      id: 'cell.restore' as const,
      unitId: context.workbook.unitId,
      sheetId: params.sheetId,
      params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.previous },
      affectedRanges: [cellRange(params.sheetId, entry.row, entry.column)],
    })),
    apply: () => {
      for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < (values[rowOffset]?.length ?? 0); columnOffset += 1) {
          const value = values[rowOffset]?.[columnOffset];
          if (value) sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, structuredClone(value));
        }
      }
    },
  });
}

function clearRangeContents(context: CommandContext, range: RangeRef): void {
  const sheet = context.workbook.getSheet(range.sheetId);
  const previous = snapshotCells(sheet, range);
  const affectedRanges = [structuredClone(range)];
  context.applyMutation({
    id: 'range.clear',
    unitId: context.workbook.unitId,
    sheetId: range.sheetId,
    params: { sheetId: range.sheetId, range, family: 'contents' as const },
    affectedRanges,
    inverse: previous.map((entry) => ({
      id: 'cell.restore' as const,
      unitId: context.workbook.unitId,
      sheetId: range.sheetId,
      params: { sheetId: range.sheetId, row: entry.row, column: entry.column, previous: entry.previous },
      affectedRanges: [cellRange(range.sheetId, entry.row, entry.column)],
    })),
    apply: () => {
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          const current = sheet.cells.get(row, column);
          if (!current) continue;
          const next = { ...current, value: null };
          delete next.formula;
          delete next.displayValue;
          sheet.cells.set(row, column, next);
        }
      }
    },
  });
}

function applyRowsInsert(context: CommandContext, sheetId: string, at: number, count: number): void {
  if (count <= 0) return;
  const sheet = context.workbook.getSheet(sheetId);
  const affectedRanges: RangeRef[] = [{ sheetId, startRow: at, endRow: at + count - 1, startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
  context.applyMutation({
    id: 'rows.inserted',
    unitId: context.workbook.unitId,
    sheetId,
    params: { sheetId, at, count },
    affectedRanges,
    inverse: [{ id: 'rows.deleted', unitId: context.workbook.unitId, sheetId, params: { sheetId, at, count }, affectedRanges }],
    apply: () => StructuralTransform.apply(context.workbook, { kind: 'insert-rows', sheetId, at, count }),
  });
}

function applyRowsDelete(context: CommandContext, sheetId: string, at: number, count: number): void {
  if (count <= 0) return;
  const sheet = context.workbook.getSheet(sheetId);
  const end = at + count - 1;
  const removed = snapshotCells(sheet, { sheetId, startRow: at, endRow: end, startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) });
  const affectedRanges: RangeRef[] = [{ sheetId, startRow: at, endRow: end, startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
  context.applyMutation({
    id: 'rows.deleted',
    unitId: context.workbook.unitId,
    sheetId,
    params: { sheetId, at, count },
    affectedRanges,
    inverse: [
      { id: 'rows.inserted', unitId: context.workbook.unitId, sheetId, params: { sheetId, at, count }, affectedRanges },
      ...removed.map((entry) => ({
        id: 'cell.restore' as const,
        unitId: context.workbook.unitId,
        sheetId,
        params: { sheetId, row: entry.row, column: entry.column, previous: entry.previous },
        affectedRanges: [cellRange(sheetId, entry.row, entry.column)],
      })),
    ],
    apply: () => StructuralTransform.apply(context.workbook, { kind: 'delete-rows', sheetId, at, count }),
  });
}

function applyOutline(context: CommandContext, sheetId: string, next: import('@react-sheets/core-model').OutlineModel, previous: import('@react-sheets/core-model').OutlineModel, affectedRanges: RangeRef[]): void {
  context.applyMutation({
    id: 'outline.set',
    unitId: context.workbook.unitId,
    sheetId,
    params: { sheetId, outline: structuredClone(next) },
    affectedRanges,
    inverse: [{ id: 'outline.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, outline: structuredClone(previous) }, affectedRanges }],
    apply: () => { context.workbook.getSheet(sheetId).outline = structuredClone(next); },
  });
}

export function normalizeConditionalFormatRule(
  rule: ConditionalFormatRule,
  fallbackPriority = 1,
): ConditionalFormatRule {
  return sheetRuleRegistry.normalizeConditionalFormat(rule, fallbackPriority, (formula) => {
    try { parseFormula(formula); }
    catch { throw new Error(`Conditional format ${rule.id} has an invalid formula predicate`); }
  });
}

export function normalizeDataValidationRule(rule: DataValidationRule): DataValidationRule {
  return sheetRuleRegistry.normalizeDataValidation(rule, (formula) => {
    try { parseFormula(formula); }
    catch { throw new Error(`Data validation ${rule.id} has an invalid formula source`); }
  });
}

export interface ConditionalOverlay {
  style?: Partial<import("@react-sheets/core-model").CellStyle>;
  dataBar?: { color: string; ratio: number };
  colorScale?: string;
  icon?: "up" | "down" | "flat";
}

/**
 * The visual state used by AutoFilter color/icon criteria.
 *
 * This is deliberately a projection, not a CellData mutation.  Direct cell
 * style is composed with the currently winning conditional-format overlay in
 * the same order used by the canvas renderer; native filter metadata remains
 * available for OOXML identities which cannot be reconstructed from a CSS
 * color alone.
 */
export interface EffectiveFilterVisual {
  style: Partial<CellStyle>;
  nativeColor?: NonNullable<CellData['filterMetadata']>['color'];
  nativeIcon?: NonNullable<CellData['filterMetadata']>['icon'];
}

export type FilterVisualResolver = (row: number, column: number, cell?: CellData) => EffectiveFilterVisual;

export function resolveEffectiveFilterVisual(
  cell: CellData | undefined,
  overlay?: ConditionalOverlay,
  presentation?: Partial<CellStyle>,
): EffectiveFilterVisual {
  const style: Partial<CellStyle> = {
    ...(cell?.style ?? {}),
    ...(presentation ?? {}),
    ...(overlay?.style ?? {}),
  };
  const nativeColor = cell?.filterMetadata?.color;
  if (nativeColor?.value !== undefined) {
    if (nativeColor.target === 'cell' && style.background === undefined) style.background = nativeColor.value;
    if (nativeColor.target === 'font' && style.textColor === undefined) style.textColor = nativeColor.value;
  }
  // The renderer gives a color-scale result precedence over both the direct
  // fill and a differential style.  Keep that precedence in the filter
  // projection so a color filter observes the rendered color.
  if (overlay?.colorScale !== undefined) style.background = overlay.colorScale;
  return {
    style,
    nativeColor,
    nativeIcon: cell?.filterMetadata?.icon,
  };
}

export function createEffectiveFilterVisualResolver(
  overlays: ReadonlyMap<string, ConditionalOverlay> | ((row: number, column: number) => ConditionalOverlay | undefined) = new Map(),
): FilterVisualResolver {
  return (row, column, cell) => resolveEffectiveFilterVisual(cell, typeof overlays === 'function' ? overlays(row, column) : overlays.get(`${row}:${column}`));
}

function cellText(resolved: FilterCellValue | undefined): string {
  return resolved?.text ?? '';
}

function cellStorageText(cell: CellData | undefined): string {
  return resolveFilterCellValue(cell).text;
}

/** Formula/spill/data-block results arrive through this typed carrier only. */
function resolvedFilterScalar(resolved: FilterCellValue | undefined): FilterScalar {
  return resolved?.value ?? null;
}

function hasCanonicalDateNumberFormat(cell: CellData | undefined): boolean {
  const format = cell?.numberFormat ?? cell?.style?.numberFormat;
  if (!format) return false;
  // Only a numeric cell carrying an explicit date/time format is a date.  In
  // particular, a text value such as "2024-01-01" remains text.
  return /(^|[^a-z])(?:d{1,4}|m{1,4}|y{2,4}|h{1,2}|s{1,2})(?:[^a-z]|$)/i.test(format);
}

function canonicalFilterDate(resolved: FilterCellValue | undefined, dateSystem: FilterDateSystem): import('@react-sheets/formula-engine').CanonicalExcelDate | null {
  const value = resolvedFilterScalar(resolved);
  const cell = resolved?.cell;
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || !hasCanonicalDateNumberFormat(cell)) return null;
    return canonicalExcelDateFromValue(value, dateSystem);
  }
  return null;
}

function numericOf(cell: CellData | undefined): number | undefined {
  const text = cellStorageText(cell);
  if (!text) return undefined;
  const cleaned = text.replace(/[$,%\s]/g, "");
  const numeric = Number(cleaned);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function hexToRgb(color: string): [number, number, number] {
  const normalized = color.replace("#", "");
  if (normalized.length === 3) {
    return [
      parseInt(normalized[0]! + normalized[0]!, 16),
      parseInt(normalized[1]! + normalized[1]!, 16),
      parseInt(normalized[2]! + normalized[2]!, 16),
    ];
  }
  return [
    parseInt(normalized.slice(0, 2), 16) || 0,
    parseInt(normalized.slice(2, 4), 16) || 0,
    parseInt(normalized.slice(4, 6), 16) || 0,
  ];
}

function mixHex(a: string, b: string, ratio: number): string {
  const [r1, g1, b1] = hexToRgb(a);
  const [r2, g2, b2] = hexToRgb(b);
  const mix = (x: number, y: number) => Math.round(x + (y - x) * ratio);
  const toHex = (value: number) => value.toString(16).padStart(2, "0");
  return "#" + toHex(mix(r1, r2)) + toHex(mix(g1, g2)) + toHex(mix(b1, b2));
}

// ---------- 条件格式求值 ----------

interface ConditionalRangeStats {
  readonly count: number;
  readonly min: number;
  readonly max: number;
  readonly topThreshold?: number;
  readonly bottomThreshold?: number;
}

function conditionalRangeKey(range: RangeRef): string {
  return `${range.sheetId}:${range.startRow}:${range.endRow}:${range.startColumn}:${range.endColumn}`;
}

function buildConditionalRangeStats(sheet: WorksheetModel, rule: ConditionalFormatRule, range: RangeRef): ConditionalRangeStats {
  const values: number[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const value = numericOf(sheet.cells.get(row, column));
      if (value !== undefined) values.push(value);
    }
  }
  const min = values.length === 0 ? 0 : values.reduce((current, value) => Math.min(current, value), Number.POSITIVE_INFINITY);
  const max = values.length === 0 ? 1 : values.reduce((current, value) => Math.max(current, value), Number.NEGATIVE_INFINITY);
  if (rule.type !== 'topBottom' || values.length === 0) return { count: values.length, min, max };
  const config = rule.topBottom ?? {
    direction: rule.operator === 'bottom' ? 'bottom' : 'top',
    rank: Number(rule.value1 ?? 10),
  };
  values.sort((left, right) => left - right);
  const rank = config.percent
    ? Math.max(1, Math.ceil(values.length * config.rank / 100))
    : Math.max(1, Math.floor(config.rank));
  return config.direction === 'top'
    ? { count: values.length, min, max, topThreshold: values[Math.max(0, values.length - rank)] }
    : { count: values.length, min, max, bottomThreshold: values[Math.min(values.length - 1, rank - 1)] };
}

/**
 * Compiles rule formulas once and evaluates only requested cells/ranges. The
 * compatibility computeConditionalOverlays wrapper below intentionally uses
 * this same runtime, so there is one evaluator and no top/bottom inner scan.
 */
export class ConditionalFormatRuntime {
  private readonly rules: ConditionalFormatRule[];
  private readonly valueCounts: Map<string, number>;
  private readonly stats = new Map<string, ConditionalRangeStats>();
  private readonly compiledFormulas = new Map<string, FormulaAst | null>();
  private readonly cellCache = new Map<string, ConditionalOverlay | undefined>();

  constructor(private readonly sheet: WorksheetModel) {
    this.rules = [...sheet.conditionalFormats].sort((left, right) =>
      (left.priority ?? Number.MAX_SAFE_INTEGER) - (right.priority ?? Number.MAX_SAFE_INTEGER));
    this.valueCounts = buildValueCounts(sheet, this.rules);
    for (const rule of this.rules) {
      if (rule.operator === 'formula') {
        try {
          const source = String(rule.value1 ?? '').trim();
          this.compiledFormulas.set(rule.id, source ? parseFormula(source.startsWith('=') ? source : `=${source}`) : null);
        } catch {
          this.compiledFormulas.set(rule.id, null);
        }
      }
      for (const range of rule.ranges) {
        if (range.sheetId !== sheet.id) continue;
        if (rule.type === 'dataBar' || rule.type === 'colorScale' || rule.type === 'iconSet' || rule.type === 'topBottom') {
          this.stats.set(`${rule.id}:${conditionalRangeKey(range)}`, buildConditionalRangeStats(sheet, rule, range));
        }
      }
    }
  }

  resolveCell(row: number, column: number): ConditionalOverlay | undefined {
    const key = `${row}:${column}`;
    if (this.cellCache.has(key)) return this.cellCache.get(key);
    let overlay: ConditionalOverlay | undefined;
    let stoppedByRule = false;
    for (const rule of this.rules) {
      const range = rule.ranges.find((candidate) => candidate.sheetId === this.sheet.id
        && row >= candidate.startRow && row <= candidate.endRow
        && column >= candidate.startColumn && column <= candidate.endColumn);
      if (!range || stoppedByRule) continue;
      const cell = this.sheet.cells.get(row, column);
      const stats = this.stats.get(`${rule.id}:${conditionalRangeKey(range)}`);
      let matches = false;
      switch (rule.type) {
        case 'highlight':
          matches = evaluateHighlight(rule, cell, this.sheet, row, column, this.valueCounts, this.compiledFormulas.get(rule.id));
          if (matches && rule.style) overlay = { ...overlay, style: { ...overlay?.style, ...rule.style } };
          break;
        case 'dataBar': {
          const numeric = numericOf(cell);
          if (numeric !== undefined && stats) {
            matches = true;
            const span = stats.max - stats.min || 1;
            overlay = { ...overlay, dataBar: { color: rule.barColor ?? '#60a5fa', ratio: (numeric - stats.min) / span } };
          }
          break;
        }
        case 'colorScale': {
          const numeric = numericOf(cell);
          if (numeric !== undefined && stats) {
            matches = true;
            const ratio = (numeric - stats.min) / (stats.max - stats.min || 1);
            const minColor = rule.minColor ?? '#fca5a5';
            const maxColor = rule.maxColor ?? '#86efac';
            overlay = { ...overlay, colorScale: rule.midColor
              ? (ratio <= 0.5 ? mixHex(minColor, rule.midColor, ratio * 2) : mixHex(rule.midColor, maxColor, (ratio - 0.5) * 2))
              : mixHex(minColor, maxColor, ratio) };
          }
          break;
        }
        case 'iconSet': {
          const numeric = numericOf(cell);
          if (numeric !== undefined && stats) {
            matches = true;
            const ratio = (numeric - stats.min) / (stats.max - stats.min || 1);
            overlay = { ...overlay, icon: ratio >= 0.67 ? 'up' : ratio >= 0.34 ? 'flat' : 'down' };
          }
          break;
        }
        case 'topBottom': {
          matches = matchesTopBottom(rule, cell, stats);
          if (matches && rule.style) overlay = { ...overlay, style: { ...overlay?.style, ...rule.style } };
          break;
        }
      }
      if (matches && rule.stopIfTrue) stoppedByRule = true;
    }
    this.cellCache.set(key, overlay);
    return overlay;
  }

  resolveRange(range: RangeRef): Map<string, ConditionalOverlay> {
    const overlays = new Map<string, ConditionalOverlay>();
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const overlay = this.resolveCell(row, column);
        if (overlay) overlays.set(`${row}:${column}`, overlay);
      }
    }
    return overlays;
  }

  resolveAll(): Map<string, ConditionalOverlay> {
    return this.resolveRange({ sheetId: this.sheet.id, startRow: 0, endRow: Math.max(0, this.sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, this.sheet.columnCount - 1) });
  }
}

export function createConditionalFormatRuntime(sheet: WorksheetModel): ConditionalFormatRuntime {
  return new ConditionalFormatRuntime(sheet);
}

export function computeConditionalOverlays(sheet: WorksheetModel): Map<string, ConditionalOverlay> {
  return new ConditionalFormatRuntime(sheet).resolveAll();
}

function createDefaultConditionalVisualResolver(sheet: WorksheetModel): FilterVisualResolver {
  const runtime = new ConditionalFormatRuntime(sheet);
  return createEffectiveFilterVisualResolver((row, column) => runtime.resolveCell(row, column));
}

function matchesTopBottom(
  rule: ConditionalFormatRule,
  cell: CellData | undefined,
  stats: ConditionalRangeStats | undefined,
): boolean {
  if (!stats || !cell) return false;
  const config = rule.topBottom ?? { direction: rule.operator === 'bottom' ? 'bottom' : 'top', rank: Number(rule.value1 ?? 10) };
  const numeric = numericOf(cell);
  if (numeric === undefined) return false;
  return config.direction === 'top' ? stats.topThreshold !== undefined && numeric >= stats.topThreshold : stats.bottomThreshold !== undefined && numeric <= stats.bottomThreshold;
}

function evaluateHighlight(
  rule: ConditionalFormatRule,
  cell: CellData | undefined,
  sheet: WorksheetModel,
  row: number,
  column: number,
  valueCounts: Map<string, number>,
  compiledFormula?: FormulaAst | null,
): boolean {
  const text = cellStorageText(cell);
  const numeric = numericOf(cell);
  const firstValue = rule.value1;
  const firstNumber = typeof firstValue === "number" ? firstValue : Number(firstValue);
  switch (rule.operator) {
    case "greaterThan": return numeric !== undefined && Number.isFinite(firstNumber) && numeric > firstNumber;
    case "lessThan": return numeric !== undefined && Number.isFinite(firstNumber) && numeric < firstNumber;
    case "equal":
      return typeof firstValue === "string" ? text.toLowerCase() === String(firstValue).toLowerCase() : numeric === firstNumber;
    case "notEqual":
      return typeof firstValue === "string" ? text.toLowerCase() !== String(firstValue).toLowerCase() : numeric !== firstNumber;
    case "between": {
      const n2 = typeof rule.value2 === "number" ? rule.value2 : Number(rule.value2);
      return numeric !== undefined && Number.isFinite(firstNumber) && Number.isFinite(n2) && numeric >= firstNumber && numeric <= n2;
    }
    case "containsText": return typeof firstValue === "string" && text.toLowerCase().includes(String(firstValue).toLowerCase());
    case "notContainsText": return typeof firstValue === "string" && !text.toLowerCase().includes(String(firstValue).toLowerCase());
    case "duplicate": return valueCounts.get(text) !== undefined && (valueCounts.get(text) ?? 0) > 1;
    case "unique": return text !== "" && (valueCounts.get(text) ?? 0) === 1;
    case "formula": return evaluateCfFormula(String(firstValue ?? ""), sheet, row, column, cell, rule.formulaAnchor ?? (rule.ranges[0] ? { sheetId: rule.ranges[0].sheetId, row: rule.ranges[0].startRow, column: rule.ranges[0].startColumn } : undefined), compiledFormula);
    default: return false;
  }
}

function evaluateCfFormula(formula: string, sheet: WorksheetModel, row: number, column: number, cell: CellData | undefined, anchor?: { sheetId: string; row: number; column: number }, compiledFormula?: FormulaAst | null): boolean {
  const source = formula.trim();
  if (!source) return false;
  if (compiledFormula === null) return false;
  try {
    const parsed = compiledFormula ?? parseFormula(source.startsWith('=') ? source : `=${source}`);
    const ast = anchor ? offsetAst(parsed, row - anchor.row, column - anchor.column) : parsed;
    const result = evaluateFormula(ast, {
      currentCell: { sheetId: sheet.id, row, column },
      readCell: (address): FormulaValue => {
        if (address.sheetId !== sheet.id) return null;
        const target = sheet.cells.get(address.row, address.column);
        return (target?.formulaValue ?? target?.value ?? null) as FormulaValue;
      },
      readRange: (range): Iterable<FormulaValue> => {
        if (range.start.sheetId !== sheet.id || range.end.sheetId !== sheet.id) return [];
        const values: FormulaValue[] = [];
        for (let targetRow = range.start.row; targetRow <= range.end.row; targetRow += 1) {
          for (let targetColumn = range.start.column; targetColumn <= range.end.column; targetColumn += 1) {
            const target = sheet.cells.get(targetRow, targetColumn);
            values.push((target?.formulaValue ?? target?.value ?? null) as FormulaValue);
          }
        }
        return values;
      },
    });
    if (isFormulaError(result)) return false;
    if (isArrayValue(result)) return Boolean(result[0]?.[0]);
    return result === true || (typeof result === 'number' && result !== 0) || (typeof result === 'string' && result.length > 0);
  } catch {
    // An unsupported/invalid CF formula is not a successful match. Do not
    // fall back to the current cell value, which would silently apply a rule
    // whose authored predicate could not be evaluated.
    return false;
  }
}

function buildValueCounts(sheet: WorksheetModel, rules: ConditionalFormatRule[]): Map<string, number> {
  const counts = new Map<string, number>();
  for (const rule of rules) {
    if (rule.operator !== "duplicate" && rule.operator !== "unique") continue;
    for (const range of rule.ranges) {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          const text = cellStorageText(sheet.cells.get(r, c));
          counts.set(text, (counts.get(text) ?? 0) + 1);
        }
      }
    }
  }
  return counts;
}

// ---------- 筛选 ----------

const FILTER_OPERATORS = new Set([
  '>', '<', '>=', '<=', '=', '<>', 'equals', 'notequal', 'notequals',
  'contains', 'notcontains', 'beginswith', 'endswith', 'blank', 'blanks',
  'notblank', 'not blanks', 'between', 'datebefore', 'dateafter', 'dateequals',
]);

export function normalizeAutoFilterModel(filter: AutoFilterModel): AutoFilterModel {
  const range = normalizeRangeRef(filter.range);
  if (range.sheetId !== filter.sheetId) throw new Error('Filter range must target its sheetId');
  const columns: Record<number, AutoFilterColumn> = {};
  for (const [rawColumn, input] of Object.entries(filter.columns)) {
    const column = Number(rawColumn);
    if (!Number.isInteger(column) || column < range.startColumn || column > range.endColumn) throw new Error('Filter criterion is outside the filter range');
    if (!input || input.column !== column) throw new Error('Filter column identity is invalid');
    columns[column] = {
      ...structuredClone(input),
      column,
      showButton: input.showButton !== false,
      hiddenButton: input.hiddenButton === true,
      criterion: input.criterion ? normalizeCriterion(input.criterion) : undefined,
    };
  }
  return { sheetId: filter.sheetId, range, columns, sortState: filter.sortState ? structuredClone(filter.sortState) : undefined, preservedXml: structuredClone(filter.preservedXml) };
}

export type FilterCellReader = (row: number, column: number) => FilterCellValue | undefined;
export type FilterDateSystem = '1900' | '1904';
export interface FilterDateContext { referenceDate: CanonicalExcelDateParts; }

export type FilterScalarType = 'blank' | 'text' | 'number' | 'boolean' | 'date';
export type FilterFamily = 'values' | 'text' | 'number' | 'date' | 'color' | 'icon';
export type FilterDateGroupUnit = 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second';

export interface FilterDomainDescriptor {
  column: number;
  values: FilterScalar[];
  scalarTypes: FilterScalarType[];
  dominantType: FilterScalarType | 'mixed' | 'empty';
  hasBlank: boolean;
  dateDomain: Array<{ value: FilterScalar; group: DateGroupItem & { hour: number; minute: number; second: number } }>;
  dateHierarchy: FilterDateGroupUnit[];
  colorDomain: Array<{ target: 'cell' | 'font'; color: string; dxfId?: number }>;
  iconDomain: Array<{ iconSet: string; iconId: number }>;
  currentFamily?: FilterFamily;
  supportedFamilies: FilterFamily[];
}

function normalizeCriterion(criterion: FilterCriterion): FilterCriterion {
  if (criterion.kind === 'values') {
    const dateGroups = criterion.dateGroups?.map((group) => normalizeDateGroupItem(group));
    return { ...structuredClone(criterion), values: [...new Set(criterion.values)], ...(dateGroups ? { dateGroups } : {}) };
  }
  if (criterion.kind === 'custom') {
    if (!criterion.conditions[0]) throw new Error('Custom filter requires a condition');
    return structuredClone(criterion);
  }
  if (criterion.kind === 'top10' && (!Number.isSafeInteger(criterion.rank) || criterion.rank <= 0)) throw new Error('Top10 filter rank must be positive');
  if (criterion.kind === 'dynamic' && !isDynamicFilterType(criterion.type)) throw new Error(`UNSUPPORTED_FEATURE: dynamic AutoFilter type "${String(criterion.type)}" is not supported`);
  return structuredClone(criterion);
}

function normalizeDateGroupItem(raw: DateGroupItem): DateGroupItem {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('Filter date group is invalid');
  const group = raw as unknown as Record<string, unknown>;
  const allowed = new Set(['year', 'month', 'day', 'hour', 'minute', 'second']);
  if (Object.keys(group).some((key) => !allowed.has(key))) throw new Error('Filter date group contains unsupported fields');
  const ranges: Record<string, [number, number]> = { year: [1, 9999], month: [1, 12], day: [1, 31], hour: [0, 23], minute: [0, 59], second: [0, 59] };
  const units = Object.keys(ranges);
  if (!Number.isSafeInteger(group.year) || Number(group.year) < ranges.year![0] || Number(group.year) > ranges.year![1]) throw new Error('Filter date group year is invalid');
  for (let index = 1; index < units.length; index += 1) {
    const unit = units[index]!;
    const previous = units[index - 1]!;
    if (group[unit] !== undefined && group[previous] === undefined) throw new Error(`Filter date group ${unit} requires ${previous}`);
    if (group[unit] !== undefined && (!Number.isSafeInteger(group[unit]) || Number(group[unit]) < ranges[unit]![0] || Number(group[unit]) > ranges[unit]![1])) throw new Error(`Filter date group ${unit} is invalid`);
  }
  return structuredClone(raw);
}

export function computeFilterHiddenRows(
  sheet: WorksheetModel,
  readCell: FilterCellReader = (row, column) => resolveFilterCellValue(sheet.cells.get(row, column)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
  dateContext?: FilterDateContext,
): Set<number> {
  const hidden = new Set<number>();
  const resolveVisual = visualResolver ?? createDefaultConditionalVisualResolver(sheet);
  let filters: import('@react-sheets/core-model').AutoFilterModel[] = [];
  try {
    filters = resolveAutoFilters(sheet).map(({ autoFilter }) => normalizeAutoFilterModel(autoFilter));
  }
  catch {
    // Malformed filter state must not expose unfiltered data accidentally.
    for (let row = 0; row < sheet.rowCount; row += 1) hidden.add(row);
    return hidden;
  }
  for (const filter of filters) {
    const table = sheet.sheetTables.find((entry) => entry.sheetId === sheet.id
      && entry.range.startRow === filter.range.startRow && entry.range.endRow === filter.range.endRow
      && entry.range.startColumn === filter.range.startColumn && entry.range.endColumn === filter.range.endColumn);
    const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
    const rows = Array.from({ length: Math.max(0, endRow - filter.range.startRow) }, (_, index) => filter.range.startRow + 1 + index);
    const top10Matches = buildTop10Matches(filter, rows, readCell, dateSystem, resolveVisual, dateContext);
    for (const row of rows) {
      const visible = Object.values(filter.columns).every((entry) => {
        const criterion = entry.criterion;
        if (!criterion) return true;
        const resolved = readCell(row, entry.column);
        const cell = resolved?.cell;
        if (criterion.kind === 'top10') return top10Matches.get(entry.column)?.has(row) ?? false;
        return matchesFilterCriterion(resolvedFilterScalar(resolved), cellText(resolved), criterion, cell, dateSystem, resolveVisual(row, entry.column, cell), dateContext);
      });
      if (!visible) hidden.add(row);
    }
  }
  return hidden;
}

/**
 * Computes the complete value domain for one filter column. The column's own
 * criterion is deliberately ignored while all other column criteria remain
 * active, matching Excel's filter menu recovery semantics.
 */
export function getAutoFilterScalarDomain(
  sheet: WorksheetModel,
  column: number,
  readCell: FilterCellReader = (row, currentColumn) => resolveFilterCellValue(sheet.cells.get(row, currentColumn)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
  dateContext?: FilterDateContext,
): FilterScalar[] {
  const resolveVisual = visualResolver ?? createDefaultConditionalVisualResolver(sheet);
  const filter = resolveAutoFilters(sheet)
    .map(({ autoFilter }) => normalizeAutoFilterModel(autoFilter))
    .find((candidate) => column >= candidate.range.startColumn && column <= candidate.range.endColumn);
  if (!filter || column < filter.range.startColumn || column > filter.range.endColumn) return [];
  const table = sheet.sheetTables.find((entry) => entry.sheetId === sheet.id
    && entry.range.startRow === filter.range.startRow && entry.range.endRow === filter.range.endRow
    && entry.range.startColumn === filter.range.startColumn && entry.range.endColumn === filter.range.endColumn);
  const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
  const values = new Map<string, FilterScalar>();
  for (let row = filter.range.startRow + 1; row <= endRow; row += 1) {
    const otherColumnsMatch = Object.values(filter.columns).every((entry) => {
      if (entry.column === column || !entry.criterion) return true;
      const resolved = readCell(row, entry.column);
      const cell = resolved?.cell;
      return matchesFilterCriterion(resolvedFilterScalar(resolved), cellText(resolved), entry.criterion, cell, dateSystem, resolveVisual(row, entry.column, cell), dateContext);
    });
    if (!otherColumnsMatch) continue;
    const resolved = readCell(row, column);
    const cell = resolved?.cell;
    const value = resolvedFilterScalar(resolved);
    values.set(JSON.stringify(value), value);
  }
  return [...values.values()].sort(compareFilterScalars);
}

export function getAutoFilterValueDomain(
  sheet: WorksheetModel,
  column: number,
  readCell: FilterCellReader = (row, currentColumn) => resolveFilterCellValue(sheet.cells.get(row, currentColumn)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
  dateContext?: FilterDateContext,
): string[] {
  return getAutoFilterScalarDomain(sheet, column, readCell, dateSystem, visualResolver, dateContext).map(filterScalarText);
}

function filterScalarText(value: FilterScalar): string {
  return value == null ? '' : String(value);
}

function compareFilterScalars(left: FilterScalar, right: FilterScalar): number {
  return compareWorkbookValues(left, right);
}

/**
 * Returns typed date coordinates for the active filter column.  The UI uses
 * this instead of parsing the display-value strings, because locale/number
 * formatting is not a stable date identity.
 */
export interface FilterDateDomainEntry {
  value: FilterScalar;
  group: DateGroupItem & { hour: number; minute: number; second: number };
}

export function getAutoFilterDateDomain(
  sheet: WorksheetModel,
  column: number,
  readCell: FilterCellReader = (row, currentColumn) => resolveFilterCellValue(sheet.cells.get(row, currentColumn)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
  dateContext?: FilterDateContext,
): FilterDateDomainEntry[] {
  const resolveVisual = visualResolver ?? createDefaultConditionalVisualResolver(sheet);
  const filter = resolveAutoFilters(sheet)
    .map(({ autoFilter }) => normalizeAutoFilterModel(autoFilter))
    .find((candidate) => column >= candidate.range.startColumn && column <= candidate.range.endColumn);
  if (!filter) return [];
  const table = sheet.sheetTables.find((entry) => entry.sheetId === sheet.id
    && entry.range.startRow === filter.range.startRow && entry.range.endRow === filter.range.endRow
    && entry.range.startColumn === filter.range.startColumn && entry.range.endColumn === filter.range.endColumn);
  const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
  const entries = new Map<string, FilterDateDomainEntry>();
  for (let row = filter.range.startRow + 1; row <= endRow; row += 1) {
    const otherColumnsMatch = Object.values(filter.columns).every((entry) => {
      if (entry.column === column || !entry.criterion) return true;
      const resolved = readCell(row, entry.column);
      const cell = resolved?.cell;
      return matchesFilterCriterion(resolvedFilterScalar(resolved), cellText(resolved), entry.criterion, cell, dateSystem, resolveVisual(row, entry.column, cell), dateContext);
    });
    if (!otherColumnsMatch) continue;
    const resolved = readCell(row, column);
    const cell = resolved?.cell;
    const value = resolvedFilterScalar(resolved);
    const date = canonicalFilterDate(resolved, dateSystem);
    if (!date) continue;
    const group = { year: date.year, month: date.month, day: date.day, hour: date.hour, minute: date.minute, second: date.second };
    const key = `${JSON.stringify(value)}|${JSON.stringify(group)}`;
    entries.set(key, { value: value as FilterScalar, group });
  }
  return [...entries.values()].sort((left, right) => JSON.stringify(left.group).localeCompare(JSON.stringify(right.group)));
}

export function getAutoFilterColorDomain(
  sheet: WorksheetModel,
  column: number,
  readCell: FilterCellReader = (row, currentColumn) => resolveFilterCellValue(sheet.cells.get(row, currentColumn)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
): Array<{ target: 'cell' | 'font'; color: string; dxfId?: number }> {
  void dateSystem;
  const resolveVisual = visualResolver ?? createDefaultConditionalVisualResolver(sheet);
  const filter = resolveAutoFilters(sheet)
    .map(({ autoFilter }) => normalizeAutoFilterModel(autoFilter))
    .find((candidate) => column >= candidate.range.startColumn && column <= candidate.range.endColumn);
  if (!filter) return [];
  const table = sheet.sheetTables.find((entry) => entry.sheetId === sheet.id
    && entry.range.startRow === filter.range.startRow && entry.range.endRow === filter.range.endRow
    && entry.range.startColumn === filter.range.startColumn && entry.range.endColumn === filter.range.endColumn);
  const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
  const options = new Map<string, { target: 'cell' | 'font'; color: string; dxfId?: number }>();
  for (let row = filter.range.startRow + 1; row <= endRow; row += 1) {
    const visual = resolveVisual(row, column, readCell(row, column)?.cell);
    if (visual.style.background) options.set(`cell:${visual.style.background}:${visual.nativeColor?.dxfId ?? ''}`, { target: 'cell', color: visual.style.background, ...(visual.nativeColor?.target === 'cell' && visual.nativeColor.dxfId !== undefined ? { dxfId: visual.nativeColor.dxfId } : {}) });
    if (visual.style.textColor) options.set(`font:${visual.style.textColor}:${visual.nativeColor?.dxfId ?? ''}`, { target: 'font', color: visual.style.textColor, ...(visual.nativeColor?.target === 'font' && visual.nativeColor.dxfId !== undefined ? { dxfId: visual.nativeColor.dxfId } : {}) });
  }
  return [...options.values()].sort((left, right) => `${left.target}:${left.color}`.localeCompare(`${right.target}:${right.color}`));
}

export function getAutoFilterIconDomain(
  sheet: WorksheetModel,
  column: number,
  readCell: FilterCellReader = (row, currentColumn) => resolveFilterCellValue(sheet.cells.get(row, currentColumn)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
): Array<{ iconSet: string; iconId: number }> {
  void dateSystem;
  const resolveVisual = visualResolver ?? createDefaultConditionalVisualResolver(sheet);
  const filter = resolveAutoFilters(sheet)
    .map(({ autoFilter }) => normalizeAutoFilterModel(autoFilter))
    .find((candidate) => column >= candidate.range.startColumn && column <= candidate.range.endColumn);
  if (!filter) return [];
  const table = sheet.sheetTables.find((entry) => entry.sheetId === sheet.id
    && entry.range.startRow === filter.range.startRow && entry.range.endRow === filter.range.endRow
    && entry.range.startColumn === filter.range.startColumn && entry.range.endColumn === filter.range.endColumn);
  const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
  const options = new Map<string, { iconSet: string; iconId: number }>();
  for (let row = filter.range.startRow + 1; row <= endRow; row += 1) {
    const icon = resolveVisual(row, column, readCell(row, column)?.cell).nativeIcon;
    if (icon) options.set(`${icon.iconSet}:${icon.iconId}`, { ...icon });
  }
  return [...options.values()].sort((left, right) => `${left.iconSet}:${left.iconId}`.localeCompare(`${right.iconSet}:${right.iconId}`));
}

export function getAutoFilterDomainDescriptor(
  sheet: WorksheetModel,
  column: number,
  readCell: FilterCellReader = (row, currentColumn) => resolveFilterCellValue(sheet.cells.get(row, currentColumn)),
  dateSystem: FilterDateSystem = '1900',
  visualResolver?: FilterVisualResolver,
  dateContext?: FilterDateContext,
): FilterDomainDescriptor {
  const values = getAutoFilterScalarDomain(sheet, column, readCell, dateSystem, visualResolver, dateContext);
  const dateDomain = getAutoFilterDateDomain(sheet, column, readCell, dateSystem, visualResolver, dateContext);
  const colorDomain = getAutoFilterColorDomain(sheet, column, readCell, dateSystem, visualResolver);
  const iconDomain = getAutoFilterIconDomain(sheet, column, readCell, dateSystem, visualResolver);
  const dateValueKeys = new Set(dateDomain.map((entry) => JSON.stringify(entry.value)));
  const scalarTypes = new Set<FilterScalarType>();
  for (const value of values) {
    if (value == null || value === '') scalarTypes.add('blank');
    else if (dateValueKeys.has(JSON.stringify(value))) scalarTypes.add('date');
    else if (typeof value === 'number') scalarTypes.add('number');
    else if (typeof value === 'boolean') scalarTypes.add('boolean');
    else scalarTypes.add('text');
  }
  const nonBlankTypes = [...scalarTypes].filter((type): type is Exclude<FilterScalarType, 'blank'> => type !== 'blank');
  const dominantType = nonBlankTypes.length === 0
    ? (values.length === 0 || scalarTypes.has('blank') ? 'empty' : 'mixed')
    : new Set(nonBlankTypes).size === 1 ? nonBlankTypes[0]! : 'mixed';
  const currentCriterion = resolveAutoFilters(sheet)
    .map(({ autoFilter }) => normalizeAutoFilterModel(autoFilter))
    .find((filter) => column >= filter.range.startColumn && column <= filter.range.endColumn)?.columns[column]?.criterion;
  const currentFamily = currentCriterion ? criterionFamily(currentCriterion, dominantType) : undefined;
  const supportedFamilies: FilterFamily[] = ['values'];
  if (dominantType === 'text') supportedFamilies.push('text');
  if (dominantType === 'number') supportedFamilies.push('number');
  if (dominantType === 'date') supportedFamilies.push('date');
  if (colorDomain.length > 0) supportedFamilies.push('color');
  if (iconDomain.length > 0) supportedFamilies.push('icon');
  const units: FilterDateGroupUnit[] = ['year', 'month', 'day', 'hour', 'minute', 'second'];
  return {
    column,
    values,
    scalarTypes: [...scalarTypes],
    dominantType,
    hasBlank: scalarTypes.has('blank'),
    dateDomain,
    dateHierarchy: units.filter((unit) => dateDomain.some((entry) => entry.group[unit] !== undefined)),
    colorDomain,
    iconDomain,
    currentFamily,
    supportedFamilies,
  };
}

function criterionFamily(criterion: FilterCriterion, dominantType: FilterDomainDescriptor['dominantType']): FilterFamily | undefined {
  if (criterion.kind === 'values') return 'values';
  if (criterion.kind === 'color') return 'color';
  if (criterion.kind === 'icon') return 'icon';
  if (criterion.kind === 'dynamic') return 'date';
  if (criterion.kind === 'top10') return 'number';
  if (dominantType === 'text') return 'text';
  if (dominantType === 'number') return 'number';
  if (dominantType === 'date') return 'date';
  return undefined;
}

const TEXT_FILTER_OPERATORS = new Set(['equals', 'notEquals', 'contains', 'notContains', 'beginsWith', 'endsWith']);
const ORDERED_FILTER_OPERATORS = new Set(['equals', 'notEquals', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual']);

/**
 * Command-layer validation for a criterion selected from a resolved domain.
 * The UI uses the same descriptor, but replay/API callers must be checked here
 * as well; an invalid family or a visual/date criterion without a domain must
 * never reach the mutation registry.
 */
export function validateFilterCriterionAgainstDomain(descriptor: FilterDomainDescriptor, criterion: FilterCriterion | undefined): void {
  if (!criterion) return;
  const family = criterionFamily(criterion, descriptor.dominantType);
  if (!family || !descriptor.supportedFamilies.includes(family)) throw new Error(`FILTER_DOMAIN_MISMATCH: ${criterion.kind} is not supported for column ${descriptor.column}`);
  if (criterion.kind === 'custom') {
    const operators = family === 'text' ? TEXT_FILTER_OPERATORS : ORDERED_FILTER_OPERATORS;
    for (const condition of criterion.conditions) {
      if (condition && !operators.has(condition.operator)) throw new Error(`FILTER_OPERATOR_MISMATCH: ${condition.operator} is not valid for ${family}`);
    }
    return;
  }
  if (criterion.kind === 'top10') {
    if (family !== 'number' || !Number.isSafeInteger(criterion.rank) || criterion.rank <= 0) throw new Error('FILTER_DOMAIN_MISMATCH: Top/Bottom requires a numeric domain and a positive safe rank');
    return;
  }
  if (criterion.kind === 'dynamic') {
    if (family !== 'date' || !isDynamicFilterType(criterion.type)) throw new Error('FILTER_DOMAIN_MISMATCH: dynamic date criteria require a canonical date domain');
    return;
  }
  if (criterion.kind === 'color') {
    const expected = criterion.target === 'cell' ? criterion.style?.background : criterion.style?.textColor;
    const available = descriptor.colorDomain.some((entry) => entry.target === criterion.target
      && ((expected !== undefined && colorsEqual(entry.color, expected)) || (criterion.dxfId >= 0 && entry.dxfId === criterion.dxfId)));
    if (!available) throw new Error('FILTER_DOMAIN_MISMATCH: color criterion is not present in the resolved color domain');
    return;
  }
  if (criterion.kind === 'icon') {
    if (!descriptor.iconDomain.some((entry) => entry.iconSet === criterion.iconSet && entry.iconId === criterion.iconId)) throw new Error('FILTER_DOMAIN_MISMATCH: icon criterion is not present in the resolved icon domain');
  }
}

function matchesFilterCriterion(
  value: unknown,
  text: string,
  criterion: FilterCriterion,
  cell?: CellData,
  dateSystem: FilterDateSystem = '1900',
  visual?: EffectiveFilterVisual,
  dateContext?: FilterDateContext,
): boolean {
  if (criterion.kind === 'values') {
    if (criterion.dateGroups?.length) {
      const date = canonicalFilterDate({ cell, value: value as FilterScalar, text }, dateSystem);
      if (date && criterion.dateGroups.some((group) => dateMatchesGroup(date, group))) return true;
    }
    if (text === '' && criterion.includeBlank) return true;
    return criterion.values.some((candidate) => String(candidate ?? '').toLocaleLowerCase() === text.toLocaleLowerCase());
  }
  if (criterion.kind === 'custom') {
    const results = criterion.conditions.filter((condition): condition is NonNullable<typeof condition> => Boolean(condition))
      .map((condition) => evaluateFilterCondition(value as FilterScalar, text, condition.operator, String(condition.value ?? '')));
    return criterion.join === 'and' ? results.every(Boolean) : results.some(Boolean);
  }
  if (criterion.kind === 'dynamic') return matchesDynamicDateFilter(value, text, criterion.type, dateSystem, cell, dateContext);
  if (criterion.kind === 'top10') return true;
  if (criterion.kind === 'color' || criterion.kind === 'icon') {
    const effective = visual ?? resolveEffectiveFilterVisual(cell);
    if (criterion.kind === 'color') {
      const nativeMatches = criterion.dxfId >= 0 && effective.nativeColor?.dxfId === criterion.dxfId;
      const expected = criterion.target === 'cell' ? criterion.style?.background : criterion.style?.textColor;
      const actual = criterion.target === 'cell' ? effective.style.background : effective.style.textColor;
      return nativeMatches || (expected !== undefined && colorsEqual(actual, expected));
    }
    return effective.nativeIcon?.iconSet === criterion.iconSet && effective.nativeIcon.iconId === criterion.iconId;
  }
  return true;
}

function colorsEqual(left: string | undefined, right: string): boolean {
  return left !== undefined && left.trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}

function dateMatchesGroup(date: import('@react-sheets/formula-engine').CanonicalExcelDate, group: import('@react-sheets/core-model').DateGroupItem): boolean {
  return date.year === group.year
    && (group.month === undefined || date.month === group.month)
    && (group.day === undefined || date.day === group.day)
    && (group.hour === undefined || date.hour === group.hour)
    && (group.minute === undefined || date.minute === group.minute)
    && (group.second === undefined || date.second === group.second);
}

function buildTop10Matches(
  filter: import('@react-sheets/core-model').AutoFilterModel,
  rows: number[],
  readCell: FilterCellReader,
  dateSystem: FilterDateSystem,
  visualResolver: FilterVisualResolver,
  dateContext?: FilterDateContext,
): Map<number, Set<number>> {
  const result = new Map<number, Set<number>>();
  for (const entry of Object.values(filter.columns)) {
    const criterion = entry.criterion;
    if (!criterion || criterion.kind !== 'top10') continue;
    const eligible = rows.filter((row) => Object.values(filter.columns).every((other) => {
      if (other.column === entry.column || !other.criterion) return true;
      const resolved = readCell(row, other.column);
      const cell = resolved?.cell;
      return other.criterion.kind === 'top10'
        ? true
        : matchesFilterCriterion(resolvedFilterScalar(resolved), cellText(resolved), other.criterion, cell, dateSystem, visualResolver(row, other.column, cell), dateContext);
    }));
    const numeric = eligible
      .map((row) => {
        const resolved = readCell(row, entry.column);
        return { row, value: numericFilterValue(resolvedFilterScalar(resolved)) };
      })
      .filter((item): item is { row: number; value: number } => item.value !== null)
      .sort((left, right) => criterion.top ? right.value - left.value : left.value - right.value);
    const requested = criterion.percent ? Math.max(1, Math.ceil(numeric.length * criterion.rank / 100)) : criterion.rank;
    const cutoff = numeric[Math.min(requested, numeric.length) - 1]?.value;
    if (cutoff === undefined) {
      result.set(entry.column, new Set());
      continue;
    }
    result.set(entry.column, new Set(numeric.filter((item) => criterion.top ? item.value >= cutoff : item.value <= cutoff).map((item) => item.row)));
  }
  return result;
}

function numericFilterValue(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  return null;
}

function matchesDynamicDateFilter(value: unknown, text: string, type: import('@react-sheets/core-model').DynamicFilterType, dateSystem: FilterDateSystem, cell?: CellData, dateContext?: FilterDateContext): boolean {
  void value;
  void text;
  const date = canonicalFilterDate({ cell, value: value as FilterScalar, text }, dateSystem);
  if (!date) return false;
  if (!dateContext) throw new Error('Dynamic date filter requires an explicit canonical workbook reference date');
  const today = canonicalExcelDateFromParts({ ...dateContext.referenceDate, hour: 0, minute: 0, second: 0, millisecond: 0 }, dateSystem);
  const startOfWeek = shiftCanonicalExcelDate(today, -((canonicalExcelDateDayOfWeek(today) + 6) % 7), dateSystem);
  const startOfMonth = monthBoundary(today, 0, dateSystem);
  const startOfQuarter = monthBoundary(today, -((today.month - 1) % 3), dateSystem);
  const startOfYear = monthBoundary(today, -(today.month - 1), dateSystem);
  const ranges: Record<import('@react-sheets/core-model').DynamicFilterType, [CanonicalExcelDate, CanonicalExcelDate]> = {
    today: [today, shiftCanonicalExcelDate(today, 1, dateSystem)], yesterday: [shiftCanonicalExcelDate(today, -1, dateSystem), today], tomorrow: [shiftCanonicalExcelDate(today, 1, dateSystem), shiftCanonicalExcelDate(today, 2, dateSystem)],
    thisWeek: [startOfWeek, shiftCanonicalExcelDate(startOfWeek, 7, dateSystem)], lastWeek: [shiftCanonicalExcelDate(startOfWeek, -7, dateSystem), startOfWeek], nextWeek: [shiftCanonicalExcelDate(startOfWeek, 7, dateSystem), shiftCanonicalExcelDate(startOfWeek, 14, dateSystem)],
    thisMonth: [startOfMonth, monthBoundary(startOfMonth, 1, dateSystem)], lastMonth: [monthBoundary(startOfMonth, -1, dateSystem), startOfMonth], nextMonth: [monthBoundary(startOfMonth, 1, dateSystem), monthBoundary(startOfMonth, 2, dateSystem)],
    thisQuarter: [startOfQuarter, monthBoundary(startOfQuarter, 3, dateSystem)], lastQuarter: [monthBoundary(startOfQuarter, -3, dateSystem), startOfQuarter], nextQuarter: [monthBoundary(startOfQuarter, 3, dateSystem), monthBoundary(startOfQuarter, 6, dateSystem)],
    thisYear: [startOfYear, monthBoundary(startOfYear, 12, dateSystem)], lastYear: [monthBoundary(startOfYear, -12, dateSystem), startOfYear], nextYear: [monthBoundary(startOfYear, 12, dateSystem), monthBoundary(startOfYear, 24, dateSystem)], yearToDate: [startOfYear, shiftCanonicalExcelDate(today, 1, dateSystem)],
  };
  const [start, end] = ranges[type];
  return date.serial >= start.serial && date.serial < end.serial;
}

function monthBoundary(value: CanonicalExcelDate, offset: number, dateSystem: FilterDateSystem): CanonicalExcelDate {
  const date = canonicalExcelDateToUtcDate(value);
  date.setUTCDate(1);
  date.setUTCMonth(date.getUTCMonth() + offset);
  return canonicalExcelDateFromUtcDate(date, dateSystem);
}

function evaluateFilterCondition(value: FilterScalar, text: string, operator: string, operand: string, operand2?: string): boolean {
  const normalizedOperator = operator.trim().toLocaleLowerCase();
  const normalizedText = text.trim().toLocaleLowerCase();
  const numeric = typeof value === 'number' && Number.isFinite(value) ? value : Number.NaN;
  const operandNumeric = Number(operand.replace(/[$,%\s,]/g, ""));
  const hasNumbers = Number.isFinite(numeric) && Number.isFinite(operandNumeric);
  const leftDate = Date.parse(text);
  const rightDate = Date.parse(operand);
  const hasDates = looksLikeDate(text) && looksLikeDate(operand) && !Number.isNaN(leftDate) && !Number.isNaN(rightDate);
  switch (normalizedOperator) {
    case ">":
    case "greaterthan": return hasDates ? leftDate > rightDate : hasNumbers && numeric > operandNumeric;
    case "<":
    case "lessthan": return hasDates ? leftDate < rightDate : hasNumbers && numeric < operandNumeric;
    case ">=":
    case "greaterthanorequal": return hasDates ? leftDate >= rightDate : hasNumbers && numeric >= operandNumeric;
    case "<=":
    case "lessthanorequal": return hasDates ? leftDate <= rightDate : hasNumbers && numeric <= operandNumeric;
    case "=":
    case "equals": return normalizedText === operand.trim().toLocaleLowerCase();
    case "<>":
    case "notequal":
    case "notequals":
    case "not equals": return normalizedText !== operand.trim().toLocaleLowerCase();
    case "contains": return normalizedText.includes(operand.trim().toLocaleLowerCase());
    case "notcontains":
    case "not contains": return !normalizedText.includes(operand.trim().toLocaleLowerCase());
    case "beginswith":
    case "begins with": return normalizedText.startsWith(operand.trim().toLocaleLowerCase());
    case "endswith":
    case "ends with": return normalizedText.endsWith(operand.trim().toLocaleLowerCase());
    case "blank":
    case "blanks": return normalizedText === '';
    case "notblank":
    case "not blanks": return normalizedText !== '';
    case "between": {
      const upper = Number((operand2 ?? '').replace(/[$,%\s,]/g, ""));
      return hasNumbers && Number.isFinite(upper) && numeric >= operandNumeric && numeric <= upper;
    }
    case "datebefore":
    case "date before": return compareDates(text, operand) < 0;
    case "dateafter":
    case "date after": return compareDates(text, operand) > 0;
    case "dateequals":
    case "date equals": return compareDates(text, operand) === 0;
    default: return false;
  }
}

function looksLikeDate(value: string): boolean {
  return /[-/:]/.test(value.trim()) && /\d/.test(value);
}

function compareDates(left: string, right: string): number {
  const leftTime = Date.parse(left);
  const rightTime = Date.parse(right);
  if (Number.isNaN(leftTime) || Number.isNaN(rightTime)) return Number.NaN;
  return leftTime === rightTime ? 0 : leftTime < rightTime ? -1 : 1;
}

// ---------- 数据验证 ----------

export interface DataValidationResult {
  valid: boolean;
  message?: string;
  blocking: boolean;
  list?: string[];
  ruleId?: string;
  alertStyle?: 'stop' | 'warning' | 'information';
}

export function findValidationRule(sheet: WorksheetModel, row: number, column: number): DataValidationRule | undefined {
  return sheet.dataValidations.find((rule) =>
    rule.ranges.some((range) =>
      row >= range.startRow && row <= range.endRow
      && column >= range.startColumn && column <= range.endColumn));
}

export function validationList(rule: DataValidationRule, sheet?: WorksheetModel): string[] | undefined {
  if (rule.type !== "list") return undefined;
  if (rule.listSource?.kind === 'values') return [...rule.listSource.values];
  if (rule.listSource?.kind === 'range' && sheet && rule.listSource.range.sheetId === sheet.id) {
    const values: string[] = [];
    const range = normalizeRangeRef(rule.listSource.range);
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const value = sheet.cells.get(row, column)?.value;
        if (value != null && String(value) !== '') values.push(String(value));
      }
    }
    return values;
  }
  const formula = rule.listSource?.kind === 'formula' ? rule.listSource.formula : rule.formula1;
  if (!formula) return undefined;
  if (sheet && formula.trim().startsWith('=')) {
    const evaluated = evaluateValidationFormula(formula, sheet, 0, 0, undefined, rule.formulaAnchor ?? (rule.ranges[0] ? { sheetId: rule.ranges[0].sheetId, row: rule.ranges[0].startRow, column: rule.ranges[0].startColumn } : undefined));
    if (isArrayValue(evaluated)) {
      return evaluated.flat().filter((value): value is string | number | boolean =>
        value !== null && !isFormulaError(value)).map(String);
    }
  }
  return splitListLiteral(formula);
}

function splitListLiteral(source: string): string[] {
  const values: string[] = [];
  let current = '';
  let quoted = false;
  for (const character of source.replace(/^=/, '')) {
    if (character === '"') { quoted = !quoted; continue; }
    if (character === ',' && !quoted) {
      if (current.trim()) values.push(current.trim());
      current = '';
      continue;
    }
    current += character;
  }
  if (current.trim()) values.push(current.trim());
  return values;
}

function evaluateValidationFormula(
  formula: string,
  sheet: WorksheetModel,
  row: number,
  column: number,
  candidate: CellData['value'] | undefined,
  anchor?: { sheetId: string; row: number; column: number },
): FormulaValue {
  try {
    const parsed = parseFormula(formula.trim().startsWith('=') ? formula.trim() : `=${formula.trim()}`);
    const ast = anchor ? offsetAst(parsed, row - anchor.row, column - anchor.column) : parsed;
    return evaluateFormula(ast, {
      currentCell: { sheetId: sheet.id, row, column },
      readCell: (address): FormulaValue => {
        if (address.sheetId !== sheet.id) return null;
        if (address.row === row && address.column === column && candidate !== undefined) return candidate as FormulaValue;
        const target = sheet.cells.get(address.row, address.column);
        return (target?.formulaValue ?? target?.value ?? null) as FormulaValue;
      },
      readRange: (range): Iterable<FormulaValue> => {
        if (range.start.sheetId !== sheet.id || range.end.sheetId !== sheet.id) return [];
        const values: FormulaValue[] = [];
        for (let targetRow = range.start.row; targetRow <= range.end.row; targetRow += 1) {
          for (let targetColumn = range.start.column; targetColumn <= range.end.column; targetColumn += 1) {
            const target = sheet.cells.get(targetRow, targetColumn);
            values.push((target?.formulaValue ?? target?.value ?? null) as FormulaValue);
          }
        }
        return values;
      },
    });
  } catch {
    return null;
  }
}

function validationMessage(rule: DataValidationRule, fallback: string): string | undefined {
  return rule.showErrorMessage === false ? undefined : rule.errorMessage ?? fallback;
}

export function validateDataInput(
  sheet: WorksheetModel,
  row: number,
  column: number,
  value: CellData["value"],
): DataValidationResult {
  const rule = findValidationRule(sheet, row, column);
  if (!rule) return { valid: true, blocking: false };
  const withRule = (result: DataValidationResult): DataValidationResult => ({
    ...result,
    ruleId: rule.id,
    alertStyle: rule.alertStyle ?? 'stop',
  });
  if (value == null || value === "") {
    const valid = Boolean(rule.allowBlank ?? true);
    return withRule({ valid, blocking: !valid && (rule.alertStyle ?? 'stop') === 'stop', message: valid ? undefined : validationMessage(rule, "该单元格不允许为空") });
  }
  const list = validationList(rule, sheet);
  if (list) {
    const candidateValues = rule.multiSelect ? String(value).split(',').map((item) => item.trim()).filter(Boolean) : [String(value)];
    const ok = candidateValues.length > 0 && candidateValues.every((candidate) =>
      list.some((item) => item.toLowerCase() === candidate.toLowerCase()));
    return withRule({
      valid: ok,
      blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop',
      message: ok ? undefined : validationMessage(rule, "值不在允许的列表中"),
      list,
    });
  }
  const numeric = typeof value === "number" ? value : Number(value);
  const isNumberType = rule.type === "whole" || rule.type === "decimal";
  if (isNumberType) {
    if (!Number.isFinite(numeric)) {
      return withRule({ valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "需要输入数字") });
    }
    if (rule.type === "whole" && !Number.isInteger(numeric)) {
      return withRule({ valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "需要输入整数") });
    }
    const bound1 = Number(rule.formula1);
    const bound2 = Number(rule.formula2);
    switch (rule.operator) {
      case "greaterThan": return withRule(judge(Number.isFinite(bound1) && numeric > bound1, rule));
      case "lessThan": return withRule(judge(Number.isFinite(bound1) && numeric < bound1, rule));
      case "equal": return withRule(judge(numeric === bound1, rule));
      case "notEqual": return withRule(judge(numeric !== bound1, rule));
      case "notBetween": return withRule(judge(Number.isFinite(bound1) && Number.isFinite(bound2) && (numeric < bound1 || numeric > bound2), rule));
      case "between":
      default: return withRule(judge(!Number.isFinite(bound1) || (!Number.isFinite(bound2) ? numeric >= bound1 : numeric >= bound1 && numeric <= bound2), rule));
    }
  }
  if (rule.type === "textLength") {
    const length = String(value).length;
    const bound1 = Number(rule.formula1);
    const bound2 = Number(rule.formula2);
    switch (rule.operator) {
      case "greaterThan": return withRule(judge(length > bound1, rule));
      case "lessThan": return withRule(judge(length < bound1, rule));
      case "equal": return withRule(judge(length === bound1, rule));
      default: return withRule(judge(!(Number.isFinite(bound1) && length < bound1) && !(Number.isFinite(bound2) && length > bound2), rule));
    }
  }
  if (rule.type === "date" || rule.type === "time") {
    const validDate = rule.type === 'date' ? isValidDateValue(value) : isValidTimeValue(value);
    if (!validDate) {
      return withRule({ valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "需要输入有效日期/时间") });
    }
    if (rule.formula1 !== undefined || rule.formula2 !== undefined || rule.operator !== undefined) {
      const actual = rule.type === 'date' ? dateComparable(value) : timeComparable(value);
      const bound1 = rule.formula1 === undefined ? Number.NaN : rule.type === 'date' ? dateComparable(rule.formula1) : timeComparable(rule.formula1);
      const bound2 = rule.formula2 === undefined ? Number.NaN : rule.type === 'date' ? dateComparable(rule.formula2) : timeComparable(rule.formula2);
      const ok = Number.isFinite(actual) && (rule.operator === 'between'
        ? Number.isFinite(bound1) && Number.isFinite(bound2) && actual >= bound1 && actual <= bound2
        : rule.operator === 'notBetween'
          ? Number.isFinite(bound1) && Number.isFinite(bound2) && (actual < bound1 || actual > bound2)
          : rule.operator === 'greaterThan' ? Number.isFinite(bound1) && actual > bound1
            : rule.operator === 'lessThan' ? Number.isFinite(bound1) && actual < bound1
              : rule.operator === 'equal' ? Number.isFinite(bound1) && actual === bound1
                : rule.operator === 'notEqual' ? Number.isFinite(bound1) && actual !== bound1
                  : true);
      return withRule(judge(ok, rule));
    }
  }
  if (rule.type === "checkbox") {
    const ok = value === true || value === false || String(value).toUpperCase() === "TRUE" || String(value).toUpperCase() === "FALSE";
    return withRule({ valid: ok, blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop', message: ok ? undefined : validationMessage(rule, "需要 TRUE/FALSE") });
  }
  if (rule.type === "custom") {
    if (!rule.formula1) return withRule({ valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "自定义验证公式缺失") });
    const evaluated = evaluateValidationFormula(rule.formula1, sheet, row, column, value, rule.formulaAnchor ?? (rule.ranges[0] ? { sheetId: rule.ranges[0].sheetId, row: rule.ranges[0].startRow, column: rule.ranges[0].startColumn } : undefined));
    const ok = evaluated === true || (typeof evaluated === 'number' && evaluated !== 0) || (typeof evaluated === 'string' && evaluated.length > 0);
    return withRule(judge(ok, rule));
  }
  return { valid: true, blocking: false };
}

function isValidDateValue(value: CellData['value']): boolean {
  if (typeof value === 'number') return Number.isFinite(value);
  if (typeof value !== 'string' || !value.trim()) return false;
  const date = new Date(value);
  return !Number.isNaN(date.getTime());
}

function isValidTimeValue(value: CellData['value']): boolean {
  if (typeof value === 'number') return Number.isFinite(value) && value >= 0 && value < 1;
  if (typeof value !== 'string') return false;
  const text = value.trim();
  if (/^(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?$/.test(text)) return true;
  const date = new Date(`1970-01-01T${text}`);
  return !Number.isNaN(date.getTime());
}

function dateComparable(value: CellData['value'] | string): number {
  if (typeof value === 'number') return value;
  const time = Date.parse(String(value));
  return Number.isNaN(time) ? Number.NaN : time;
}

function timeComparable(value: CellData['value'] | string): number {
  if (typeof value === 'number') return value * 86400;
  const match = String(value).trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return Number.NaN;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] ?? 0);
  return hours * 3600 + minutes * 60 + seconds;
}

function judge(ok: boolean, rule: DataValidationRule): DataValidationResult {
  return {
    valid: ok,
    blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop',
    message: ok ? undefined : validationMessage(rule, "不符合数据验证规则"),
  };
}

/** 大纲折叠隐藏行 — 独立于 filter hiddenRows 语义；保留分组首行可见 */
export function computeOutlineHiddenRows(sheet: WorksheetModel): Set<number> {
  const hidden = new Set<number>();
  const groups = sheet.outline?.groups ?? [];
  for (const group of groups) {
    if (group.axis !== 'row' || !group.collapsed) continue;
    for (let row = group.start + 1; row <= group.end; row++) hidden.add(row);
  }
  return hidden;
}

export function computeOutlineHiddenColumns(sheet: WorksheetModel): Set<number> {
  const hidden = new Set<number>();
  const groups = sheet.outline?.groups ?? [];
  for (const group of groups) {
    if (group.axis !== 'column' || !group.collapsed) continue;
    for (let column = group.start + 1; column <= group.end; column++) hidden.add(column);
  }
  return hidden;
}

export interface TextToColumnsParams {
  sheetId: string;
  range: RangeRef;
  delimiter: string;
  maxColumns?: number;
}

export interface RemoveDuplicatesParams {
  sheetId: string;
  range: RangeRef;
  columns: number[];
  hasHeader?: boolean;
}

export interface DataSortParams {
  sheetId: string;
  range: RangeRef;
  criteria: Array<{ column: number; ascending: boolean }>;
  hasHeader?: boolean;
  dataRegionContext?: DataRegionContext;
}

export interface SubtotalParams {
  sheetId: string;
  range: RangeRef;
  groupColumn: number;
  valueColumn: number;
  functionName: 'SUM' | 'COUNT' | 'AVERAGE';
}

/** A resolved scalar only; array/range formula results are not sortable. */
export type SortCellValue = ScalarValue | FormulaError;

function normalizeSortCellValue(value: unknown): SortCellValue {
  if (Array.isArray(value)) {
    if (value.length === 1 && Array.isArray(value[0]) && value[0].length === 1) return normalizeSortCellValue(value[0][0]);
    throw new Error('Sort key cannot be an unresolved array result');
  }
  if (value === undefined || value === null) return null;
  if (isFormulaError(value)) return value;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('Sort key contains a non-finite numeric result');
    return value;
  }
  if (typeof value === 'string' || typeof value === 'boolean') return value;
  throw new Error(`Sort key has unsupported resolved value type: ${typeof value}`);
}

export function compareSortValues(left: SortCellValue, right: SortCellValue): number {
  return compareWorkbookValues(left, right);
}

export function resolveSortCellValue(
  sheet: WorksheetModel,
  row: number,
  column: number,
  resolver?: (sheet: WorksheetModel, row: number, column: number) => unknown,
): SortCellValue {
  const cell = sheet.cells.get(row, column);
  const resolved = resolver?.(sheet, row, column);
  if (resolved !== undefined) return normalizeSortCellValue(resolved);
  if (cell?.formula !== undefined && cell.formulaValue === undefined) {
    throw new Error(`Sort formula result unavailable at ${sheet.id}!${row}:${column}`);
  }
  return normalizeSortCellValue(cell?.formulaValue ?? cell?.value ?? null);
}

function sortedSourceRows(
  sheet: WorksheetModel,
  params: DataSortParams,
  resolver?: (sheet: WorksheetModel, row: number, column: number) => unknown,
): number[] {
  const range = normalizeRangeRef(params.range);
  const startRow = (params.hasHeader ?? false) ? range.startRow + 1 : range.startRow;
  if (startRow > range.endRow) return [];
  for (const criterion of params.criteria) {
    if (!Number.isInteger(criterion.column) || criterion.column < range.startColumn || criterion.column > range.endColumn) {
      throw new Error('Sort criterion is outside the selected range');
    }
  }
  const rows = Array.from({ length: range.endRow - startRow + 1 }, (_, offset) => startRow + offset);
  rows.sort((leftRow, rightRow) => {
    for (const criterion of params.criteria) {
      const result = compareSortValues(
        resolveSortCellValue(sheet, leftRow, criterion.column, resolver),
        resolveSortCellValue(sheet, rightRow, criterion.column, resolver),
      );
      if (result !== 0) return criterion.ascending ? result : -result;
    }
    return leftRow - rightRow;
  });
  return rows;
}

function selectedRange(params: { sheetId: string; range: RangeRef }): RangeRef {
  return normalizeRangeRef({ ...params.range, sheetId: params.sheetId });
}

function assertNoDataRegionIntersection(sheet: WorksheetModel, range: RangeRef, operation: string): void {
  const region = sheet.dataRegions.find((candidate) => candidate.range.sheetId === range.sheetId
    && candidate.range.startRow <= range.endRow && candidate.range.endRow >= range.startRow
    && candidate.range.startColumn <= range.endColumn && candidate.range.endColumn >= range.startColumn);
  if (region) throw new Error(`${operation} does not support data-region ${region.id} without the canonical resolved-cell transaction`);
}

function subtotalFormula(functionName: SubtotalParams['functionName'], column: number, startRow: number, endRow: number): string {
  const code = functionName === 'SUM' ? 9 : functionName === 'COUNT' ? 3 : 1;
  return `=SUBTOTAL(${code},${columnLabel(column)}${startRow + 1}:${columnLabel(column)}${endRow + 1})`;
}

function transformMatrixFormula(
  formula: string,
  range: RangeRef,
  mapCoordinate: (row: number, column: number) => { row: number; column: number },
): string {
  if (!formula.trim().startsWith('=')) return formula;
  const ast = parseFormula(formula);
  return formatFormula(mapAstReferences(ast, (reference: ParsedCellReference) => {
    if (reference.sheetId !== undefined && reference.sheetId.toLocaleLowerCase() !== range.sheetId.toLocaleLowerCase()) return reference;
    if (reference.row < range.startRow || reference.row > range.endRow
      || reference.column < range.startColumn || reference.column > range.endColumn) return reference;
    const mapped = mapCoordinate(reference.row, reference.column);
    return { ...reference, row: mapped.row, column: mapped.column };
  }));
}

function assertMatrixTransformSupported(sheet: WorksheetModel, range: RangeRef): void {
  const intersects = (candidate: RangeRef): boolean => candidate.sheetId === range.sheetId
    && candidate.startRow <= range.endRow && candidate.endRow >= range.startRow
    && candidate.startColumn <= range.endColumn && candidate.endColumn >= range.startColumn;
  if (sheet.merges.some((merge) => intersects(merge.range))) throw new Error('Matrix transform cannot partially or silently rewrite merged cells');
  if (sheet.sheetTables.some((table) => intersects(table.range))) throw new Error('Matrix transform cannot rewrite a Sheet Table');
  if (sheet.conditionalFormats.some((rule) => rule.ranges.some(intersects))) throw new Error('Matrix transform cannot rewrite conditional-format ranges');
  if (sheet.dataValidations.some((rule) => rule.ranges.some(intersects))) throw new Error('Matrix transform cannot rewrite validation ranges');
  if (sheet.autoFilter && intersects(sheet.autoFilter.range)) throw new Error('Matrix transform cannot rewrite a filtered range');
  if (sheet.drawings.some((drawing) => drawing.anchor.kind !== 'absolute'
    && drawing.anchor.row !== undefined && drawing.anchor.row >= range.startRow && drawing.anchor.row <= range.endRow)) {
    throw new Error('Matrix transform cannot rewrite drawing anchors');
  }
  if (sheet.review.noteCount > 0 || sheet.review.threadEntries().some((thread) => thread.row >= range.startRow && thread.row <= range.endRow)) {
    throw new Error('Matrix transform cannot rewrite review objects');
  }
  if (sheet.sparklines.some((sparkline) => sparkline.anchor.row >= range.startRow && sparkline.anchor.row <= range.endRow)) {
    throw new Error('Matrix transform cannot rewrite sparklines');
  }
}

function matrixTargetRange(range: RangeRef, transpose: boolean): RangeRef {
  return transpose
    ? {
      sheetId: range.sheetId,
      startRow: range.startRow,
      endRow: range.startRow + (range.endColumn - range.startColumn),
      startColumn: range.startColumn,
      endColumn: range.startColumn + (range.endRow - range.startRow),
    }
    : structuredClone(range);
}

function matrixClearRange(source: RangeRef, target: RangeRef): RangeRef {
  return {
    sheetId: source.sheetId,
    startRow: Math.min(source.startRow, target.startRow),
    endRow: Math.max(source.endRow, target.endRow),
    startColumn: Math.min(source.startColumn, target.startColumn),
    endColumn: Math.max(source.endColumn, target.endColumn),
  };
}

function executeMatrixTransform(
  runtime: CommandRuntime,
  context: CommandContext,
  params: { sheetId: string; range: RangeRef; direction?: 'horizontal' | 'vertical' },
  transpose: boolean,
): ReturnType<CommandRuntime['execute']> {
  const sheet = runtime.workbook.getSheet(params.sheetId);
  const range = normalizeRangeRef({ ...params.range, sheetId: params.sheetId });
  assertMatrixTransformSupported(sheet, range);
  const target = matrixTargetRange(range, transpose);
  if (target.endRow >= sheet.rowCount || target.endColumn >= sheet.columnCount) throw new Error('Matrix transform exceeds worksheet bounds');
  const clearRange = matrixClearRange(range, target);
  if (target.startRow !== range.startRow || target.startColumn !== range.startColumn
    || target.endRow !== range.endRow || target.endColumn !== range.endColumn) {
    for (let row = target.startRow; row <= target.endRow; row += 1) {
      for (let column = target.startColumn; column <= target.endColumn; column += 1) {
        if (!inRange(range, row, column) && sheet.cells.get(row, column)) {
          throw new Error('Matrix transform would overwrite cells outside the selected range');
        }
      }
    }
  }
  const values: CellData[][] = [];
  for (let row = target.startRow; row <= target.endRow; row += 1) {
    const line: CellData[] = [];
    for (let column = target.startColumn; column <= target.endColumn; column += 1) {
      const sourceRow = transpose ? range.startRow + (column - target.startColumn) : (
        params.direction === 'vertical' ? range.endRow - (row - target.startRow) : row);
      const sourceColumn = transpose ? range.startColumn + (row - target.startRow) : (
        params.direction === 'horizontal' ? range.endColumn - (column - target.startColumn) : column);
      const source = structuredClone(sheet.cells.get(sourceRow, sourceColumn) ?? { value: null });
      if (source.formula) {
        source.formula = transformMatrixFormula(source.formula, range, (refRow, refColumn) => transpose
          ? { row: range.startRow + (refColumn - range.startColumn), column: range.startColumn + (refRow - range.startRow) }
          : params.direction === 'horizontal'
            ? { row: refRow, column: range.startColumn + range.endColumn - refColumn }
            : { row: range.startRow + range.endRow - refRow, column: refColumn });
      }
      line.push(source);
    }
    values.push(line);
  }
  clearRangeContents(context, clearRange);
  applyRangeValues(context, {
    sheetId: params.sheetId,
    startRow: target.startRow,
    startColumn: target.startColumn,
    values,
  });
  return { operationId: context.operationId, mutationCount: 2, affectedRanges: [clearRange, target] };
}

function contiguousGroups(sheet: WorksheetModel, params: SubtotalParams): Array<{ start: number; end: number; key: string }> {
  const range = normalizeRangeRef(params.range);
  if (params.groupColumn < range.startColumn || params.groupColumn > range.endColumn) throw new Error('Subtotal group column is outside the range');
  if (params.valueColumn < range.startColumn || params.valueColumn > range.endColumn) throw new Error('Subtotal value column is outside the range');
  const groups: Array<{ start: number; end: number; key: string }> = [];
  let start = range.startRow + 1;
  if (start > range.endRow) return groups;
  let key = cellStorageText(sheet.cells.get(start, params.groupColumn));
  for (let row = start + 1; row <= range.endRow + 1; row += 1) {
    const next = row <= range.endRow ? cellStorageText(sheet.cells.get(row, params.groupColumn)) : undefined;
    if (next !== key) {
      groups.push({ start, end: row - 1, key });
      if (next !== undefined) {
        start = row;
        key = next;
      }
    }
  }
  return groups;
}

export function registerDataToolCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<RowsPermutedMutationParams>({
    id: 'rows.permuted',
    handler: (item, context) => {
      if (!isRowsPermutedMutation(item.params)) throw new Error('Invalid rows.permuted mutation payload');
      const params = item.params;
      const range = params.range;
      const sheet = context.workbook.getSheet(params.sheetId);
      applyRowPermutation(sheet, createRowPermutationPlan(range, params.sourceRows));
      setAppliedSortState(sheet, params.sortState);
    },
    metadata: {
      schema: { name: 'RowsPermuted', validate: isRowsPermutedMutation },
      permission: { capability: 'sheet.sort.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: rowsPermutedAffectedRanges, mode: 'exact' },
      inverseIds: ['rows.permuted'],
    },
  });

  runtime.registry.registerCommand<DataSortParams>({
    id: 'data.sort.rows',
    execute: (params, context) => {
      const requestedRange = selectedRange(params);
      const sheet = context.workbook.getSheet(params.sheetId);
      const regionContext = params.dataRegionContext ?? resolveDataRegionContext(context.workbook, {
        selection: requestedRange,
        activeRow: requestedRange.startRow,
        activeColumn: params.criteria[0]?.column ?? requestedRange.startColumn,
      });
      if (params.dataRegionContext) {
        const actual = resolveDataRegionContext(context.workbook, {
          selection: requestedRange,
          activeRow: requestedRange.startRow,
          activeColumn: params.criteria[0]?.column ?? requestedRange.startColumn,
        });
        assertDataRegionContextMatches(params.dataRegionContext, actual);
      }
      const range = normalizeRangeRef(regionContext.range);
      const hasHeader = params.hasHeader ?? regionContext.header.kind === 'present';
      assertNoDataRegionIntersection(sheet, range, 'Sort');
      if (params.criteria.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const sourceRows = sortedSourceRows(sheet, { ...params, range, hasHeader }, context.resolveCellValue);
      if (sourceRows.length <= 1 || sourceRows.every((row, offset) => row === range.startRow + (hasHeader ? 1 : 0) + offset)) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const startRow = hasHeader ? range.startRow + 1 : range.startRow;
      const bodyRange: RangeRef = { ...range, startRow };
      const inverseRows = new Array<number>(sourceRows.length);
      sourceRows.forEach((sourceRow, offset) => { inverseRows[sourceRow - startRow] = startRow + offset; });
      const affectedColumnEnd = rowsPermutedAffectedColumnEnd(sheet, bodyRange);
      const affectedRanges = rowsPermutedAffectedRanges({ sheetId: params.sheetId, range: bodyRange, sourceRows, affectedColumnEnd });
      context.applyMutation({
        id: 'rows.permuted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: {
          ...params,
          hasHeader,
          dataRegionContext: regionContext,
          range: bodyRange,
          sourceRows,
          affectedColumnEnd,
          sortState: {
            sheetId: params.sheetId,
            range,
            criteria: structuredClone(params.criteria),
            hasHeader,
            revision: ((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState?.revision ?? 0) + 1,
          },
          previousSortState: ((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState
            ? structuredClone((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState)
            : undefined),
        },
        affectedRanges,
        inverse: [{
          id: 'rows.permuted',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: {
            ...params,
            hasHeader,
            dataRegionContext: regionContext,
            range: bodyRange,
            sourceRows: inverseRows,
            affectedColumnEnd,
            sortState: ((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState
              ? structuredClone((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState)
              : undefined),
            previousSortState: ((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState
              ? structuredClone((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState)
              : undefined),
          },
          affectedRanges,
        }],
        apply: () => applyRowPermutation(sheet, createRowPermutationPlan(bodyRange, sourceRows)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; range: RangeRef }>({
    id: 'matrix.transpose',
    execute: (params, context) => executeMatrixTransform(runtime, context, params, true),
  });

  runtime.registry.registerCommand<{ sheetId: string; range: RangeRef; direction: 'horizontal' | 'vertical' }>({
    id: 'matrix.flip',
    execute: (params, context) => executeMatrixTransform(runtime, context, params, false),
  });

  runtime.registry.registerCommand<TextToColumnsParams>({
    id: 'data.textToColumns',
    execute: (params, context) => {
      if (!params.delimiter) throw new Error('Text to Columns delimiter is required');
      const sheet = context.workbook.getSheet(params.sheetId);
      const range = normalizeRangeRef(params.range);
      const maxColumns = Math.max(2, params.maxColumns ?? 8);
      if (range.startColumn + maxColumns > sheet.columnCount) throw new Error('Text to Columns exceeds worksheet bounds');
      const values: CellData[][] = [];
      for (let row = range.startRow; row <= range.endRow; row++) {
        const cell = sheet.cells.get(row, range.startColumn);
        const text = cell?.value == null ? '' : String(cell.value);
        const parts = text.split(params.delimiter).slice(0, maxColumns);
        values.push(parts.map((part) => ({ value: part })));
      }
      clearRangeContents(context, {
        sheetId: params.sheetId,
        startRow: range.startRow,
        endRow: range.endRow,
        startColumn: range.startColumn,
        endColumn: range.startColumn + maxColumns - 1,
      });
      applyRangeValues(context, {
        sheetId: params.sheetId,
        startRow: range.startRow,
        startColumn: range.startColumn,
        values,
      });
      return { operationId: context.operationId, mutationCount: 2, affectedRanges: [range] };
    },
  });

  runtime.registry.registerCommand<RemoveDuplicatesParams>({
    id: 'data.removeDuplicates',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const range = selectedRange(params);
      if (params.columns.length === 0 || params.columns.some((column) => column < range.startColumn || column > range.endColumn)) {
        throw new Error('Remove Duplicates columns must be inside the selected range');
      }
      const startRow = params.hasHeader ? range.startRow + 1 : range.startRow;
      const seen = new Set<string>();
      const duplicateRows: number[] = [];
      for (let row = startRow; row <= range.endRow; row++) {
        const key = params.columns.map((column) => cellStorageText(sheet.cells.get(row, column))).join('\u0001');
        if (seen.has(key)) { duplicateRows.push(row); continue; }
        seen.add(key);
      }
      if (duplicateRows.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      // Delete complete worksheet rows from the bottom. StructuralTransform
      // carries formulas, drawings, notes, validation and table metadata with
      // each row, instead of clearing the selected cells and losing them.
      const runs: Array<{ at: number; count: number }> = [];
      for (const row of duplicateRows) {
        const last = runs.at(-1);
        if (last && last.at + last.count === row) last.count += 1;
        else runs.push({ at: row, count: 1 });
      }
      let mutationCount = 0;
      const affectedRanges: RangeRef[] = [];
      for (const run of [...runs].reverse()) {
        applyRowsDelete(context, params.sheetId, run.at, run.count);
        mutationCount += 1;
        affectedRanges.push({ sheetId: params.sheetId, startRow: run.at, endRow: run.at + run.count - 1, startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) });
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });

  runtime.registry.registerCommand<SubtotalParams>({
    id: 'data.subtotal',
    execute: (params, context) => {
      if (!['SUM', 'COUNT', 'AVERAGE'].includes(params.functionName)) throw new Error('Unsupported Subtotal function');
      const sheet = context.workbook.getSheet(params.sheetId);
      const range = selectedRange(params);
      const groups = contiguousGroups(sheet, { ...params, range });
      if (groups.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      let mutationCount = 0;
      const targetRow = range.endRow + 2;
      const summaryEnd = targetRow + groups.length;
      // The summary block is appended after one blank row, matching the
      // workbook's existing Subtotal presentation. If the destination is
      // occupied, insert the complete block first so no user data is lost.
      let occupied = false;
      for (let row = targetRow; row <= summaryEnd; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          if (sheet.cells.get(row, column)) occupied = true;
        }
      }
      if (occupied) {
        applyRowsInsert(context, params.sheetId, targetRow, groups.length + 1);
        mutationCount += 1;
      }
      const values: CellData[][] = [];
      const header: CellData[] = [];
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        header.push({ value: column === params.groupColumn ? 'Group' : column === params.valueColumn ? params.functionName : null, style: { bold: true } });
      }
      values.push(header);
      for (const group of groups) {
        const rowValues: CellData[] = [];
        const numbers: number[] = [];
        for (let row = group.start; row <= group.end; row += 1) {
          const numeric = numericOf(sheet.cells.get(row, params.valueColumn));
          if (numeric !== undefined) numbers.push(numeric);
        }
        const cachedValue = params.functionName === 'SUM'
          ? numbers.reduce((sum, value) => sum + value, 0)
          : params.functionName === 'COUNT'
            ? numbers.length
            : numbers.length > 0 ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null;
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          if (column === params.groupColumn) rowValues.push({ value: group.key, style: { bold: true } });
          else if (column === params.valueColumn) rowValues.push({ value: cachedValue, formula: subtotalFormula(params.functionName, params.valueColumn, group.start, group.end), style: { bold: true } });
          else rowValues.push({ value: null });
        }
        values.push(rowValues);
      }
      applyRangeValues(context, {
        sheetId: params.sheetId,
        startRow: targetRow,
        startColumn: range.startColumn,
        values,
      });
      mutationCount += 1;
      const affectedRanges: RangeRef[] = [structuredClone(range), { sheetId: params.sheetId, startRow: targetRow, endRow: targetRow + groups.length, startColumn: range.startColumn, endColumn: range.endColumn }];
      for (const group of groups) {
        const sheetOutline = sheet.outline ? structuredClone(sheet.outline) : { groups: [] };
        const nextOutline = structuredClone(sheetOutline);
        nextOutline.groups.push({ id: `subtotal-${context.operationId}-${group.start}-${group.end}`, axis: 'row', start: group.start, end: group.end, level: 1, collapsed: false });
        applyOutline(context, params.sheetId, nextOutline, sheetOutline, [{ sheetId: params.sheetId, startRow: group.start, endRow: group.end, startColumn: range.startColumn, endColumn: range.endColumn }]);
        mutationCount += 1;
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; row: number; column: number; delimiter: string; maxColumns?: number }>({
    id: 'data.splitColumn',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const cell = sheet.cells.get(params.row, params.column);
      const text = cell?.value == null ? '' : String(cell.value);
      const maxColumns = Math.max(2, params.maxColumns ?? 4);
      if (params.column + maxColumns > sheet.columnCount) throw new Error('Split Column exceeds worksheet bounds');
      const parts = text.split(params.delimiter).slice(0, maxColumns);
      if (parts.length <= 1 && parts[0] === text) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const values = [parts.map((part) => ({ value: coerceDataText(part, cell), style: cell?.style ? structuredClone(cell.style) : undefined }))];
      const range: RangeRef = { sheetId: params.sheetId, startRow: params.row, endRow: params.row, startColumn: params.column, endColumn: params.column + maxColumns - 1 };
      clearRangeContents(context, range);
      applyRangeValues(context, { sheetId: params.sheetId, startRow: params.row, startColumn: params.column, values });
      return { operationId: context.operationId, mutationCount: 2, affectedRanges: [range] };
    },
  });

  runtime.registry.registerCommand<{ name: string }>({
    id: 'workbook.name.list',
    execute: (_params, context) => {
      const names = context.workbook.listDefinedNames().map((entry) => entry.name).sort();
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: names.map(() => ({ sheetId: context.workbook.primarySheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })) };
    },
  });
}

function coerceDataText(text: string, previousCell: CellData | undefined): CellData['value'] {
  if (typeof previousCell?.value === 'number') {
    const numeric = Number(text.replace(/[$,%]/g, ''));
    if (Number.isFinite(numeric) && text.trim() !== '') return numeric;
  }
  return text;
}
