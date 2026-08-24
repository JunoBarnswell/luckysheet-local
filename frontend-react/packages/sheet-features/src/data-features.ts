import type {
  CellData,
  ConditionalFormatRule,
  DataValidationRule,
  AutoFilterColumn,
  AutoFilterModel,
  FilterCriterion,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
  ConditionalFormatTopBottom,
} from "@react-sheets/core-model";
import { StructuralTransform, applyRowPermutation, columnLabel, validatePermutationMetadata } from "@react-sheets/core-model";
import { resolveWorksheetAutoFilter } from './sheet-table-features';
import type { CommandContext, CommandRuntime } from "@react-sheets/command-runtime";
import {
  evaluateFormula,
  formatFormula,
  isArrayValue,
  isFormulaError,
  mapAstReferences,
  parseFormula,
  type ParsedCellReference,
  type FormulaValue,
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
  sortState?: AppliedSortState;
  previousSortState?: AppliedSortState;
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
    && params.sourceRows.length === Number(candidate.endRow) - Number(candidate.startRow) + 1
    && params.sourceRows.every((row) => Number.isInteger(row) && Number(row) >= Number(candidate.startRow) && Number(row) <= Number(candidate.endRow));
}

function setAppliedSortState(sheet: WorksheetModel, state: AppliedSortState | undefined): void {
  const target = sheet as WorksheetModel & { appliedSortState?: AppliedSortState };
  if (state === undefined) delete target.appliedSortState;
  else target.appliedSortState = structuredClone(state);
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
  const previous = snapshotCells(sheet, range);
  const affectedRanges = [range];
  context.applyMutation({
    id: 'range.set',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: previous.map((entry) => ({
      id: 'cell.restore' as const,
      unitId: context.workbook.unitId,
      sheetId: params.sheetId,
      params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.previous },
      affectedRanges: [cellRange(params.sheetId, entry.row, entry.column)],
    })),
    apply: () => {
      for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < (params.values[rowOffset]?.length ?? 0); columnOffset += 1) {
          const value = params.values[rowOffset]?.[columnOffset];
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
    params: { sheetId: range.sheetId, range, mode: 'contents' as const },
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
  const ranges = rule.ranges.map(normalizeRangeRef);
  if (ranges.length === 0) throw new Error(`Conditional format ${rule.id} requires at least one range`);
  if (!rule.id.trim()) throw new Error('Conditional format id is required');
  if (ranges.some((range) => range.sheetId !== rule.sheetId)) throw new Error('Conditional format ranges must target the rule sheet');
  if (rule.priority !== undefined && (!Number.isInteger(rule.priority) || rule.priority <= 0)) throw new Error('Conditional format priority must be a positive integer');
  if (rule.stopIfTrue !== undefined && typeof rule.stopIfTrue !== 'boolean') throw new Error('Conditional format stopIfTrue must be boolean');
  if (rule.operator === 'formula') {
    const formula = String(rule.value1 ?? '').trim();
    if (!formula) throw new Error(`Conditional format ${rule.id} requires a formula predicate`);
    try { parseFormula(formula.startsWith('=') ? formula : `=${formula}`); }
    catch { throw new Error(`Conditional format ${rule.id} has an invalid formula predicate`); }
  }
  if (rule.type === 'topBottom') {
    const topBottom: ConditionalFormatTopBottom = rule.topBottom ?? {
      direction: rule.operator === 'bottom' ? 'bottom' : 'top',
      rank: Number(rule.value1 ?? 10),
    };
    if (!Number.isFinite(topBottom.rank) || topBottom.rank <= 0) throw new Error('Top/Bottom rank must be positive');
    if (topBottom.percent && topBottom.rank > 100) throw new Error('Top/Bottom percentage cannot exceed 100');
    return {
      ...structuredClone(rule),
      ranges,
      priority: Number.isFinite(rule.priority) && (rule.priority ?? 0) > 0 ? rule.priority : fallbackPriority,
      stopIfTrue: rule.stopIfTrue ?? false,
      topBottom: { ...topBottom },
    };
  }
  return {
    ...structuredClone(rule),
    ranges,
    priority: Number.isFinite(rule.priority) && (rule.priority ?? 0) > 0 ? rule.priority : fallbackPriority,
    stopIfTrue: rule.stopIfTrue ?? false,
  };
}

export function normalizeDataValidationRule(rule: DataValidationRule): DataValidationRule {
  if (!rule.id.trim()) throw new Error('Data validation id is required');
  if (rule.ranges.length === 0) throw new Error(`Data validation ${rule.id} requires at least one range`);
  if (rule.type === 'list' && !rule.formula1 && !rule.listSource) {
    throw new Error(`List validation ${rule.id} requires a list source`);
  }
  if ((rule.type === 'whole' || rule.type === 'decimal' || rule.type === 'textLength') && rule.formula1 === undefined) {
    throw new Error(`Data validation ${rule.id} requires a lower bound`);
  }
  if (rule.type === 'checkbox' && rule.operator !== undefined) {
    throw new Error('Checkbox validation does not accept a comparison operator');
  }
  if (rule.alertStyle !== undefined && !['stop', 'warning', 'information'].includes(rule.alertStyle)) {
    throw new Error(`Unsupported data validation alert style: ${rule.alertStyle}`);
  }
  if (rule.listSource?.kind === 'range' && rule.listSource.range.sheetId !== rule.sheetId) {
    throw new Error('Validation list range must target the validation sheet');
  }
  const formula = rule.type === 'custom' ? rule.formula1 : rule.listSource?.kind === 'formula' ? rule.listSource.formula : undefined;
  if (formula?.trim().startsWith('=')) {
    try { parseFormula(formula.trim()); }
    catch { throw new Error(`Data validation ${rule.id} has an invalid formula source`); }
  }
  return {
    ...structuredClone(rule),
    ranges: rule.ranges.map(normalizeRangeRef),
    alertStyle: rule.alertStyle ?? 'stop',
    showErrorMessage: rule.showErrorMessage ?? true,
    showInputMessage: rule.showInputMessage ?? false,
    allowBlank: rule.allowBlank ?? true,
  };
}

export interface ConditionalOverlay {
  style?: Partial<import("@react-sheets/core-model").CellStyle>;
  dataBar?: { color: string; ratio: number };
  colorScale?: string;
  icon?: "up" | "down" | "flat";
}

function cellText(cell: CellData | undefined): string {
  if (!cell || cell.value == null) return "";
  return String(cell.value);
}

function numericOf(cell: CellData | undefined): number | undefined {
  const text = cellText(cell);
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

export function computeConditionalOverlays(sheet: WorksheetModel): Map<string, ConditionalOverlay> {
  const overlays = new Map<string, ConditionalOverlay>();
  const rules = [...sheet.conditionalFormats].sort((left, right) =>
    (left.priority ?? Number.MAX_SAFE_INTEGER) - (right.priority ?? Number.MAX_SAFE_INTEGER));
  if (rules.length === 0) return overlays;
  const valueCounts = buildValueCounts(sheet, rules);
  const stopped = new Set<string>();

  // 预收集各规则范围的数值用于色阶/数据条归一化
  for (const rule of rules) {
    for (const range of rule.ranges) {
      if (range.sheetId !== sheet.id) continue;
      let min = Number.POSITIVE_INFINITY;
      let max = Number.NEGATIVE_INFINITY;
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          const numeric = numericOf(sheet.cells.get(r, c));
          if (numeric !== undefined) {
            min = Math.min(min, numeric);
            max = Math.max(max, numeric);
          }
        }
      }
      const boundedMin = Number.isFinite(min) ? min : 0;
      const boundedMax = Number.isFinite(max) ? max : 1;
      const span = boundedMax - boundedMin || 1;

      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          const cell = sheet.cells.get(r, c);
          const key = r + ":" + c;
          if (stopped.has(key)) continue;
          let overlay = overlays.get(key) ?? {};
          let matches = false;

          switch (rule.type) {
            case "highlight": {
              matches = evaluateHighlight(rule, cell, sheet, r, c, valueCounts);
              if (matches && rule.style) overlay = { ...overlay, style: { ...overlay.style, ...rule.style } };
              break;
            }
            case "dataBar": {
            const numeric = numericOf(cell);
            if (numeric !== undefined) {
              matches = true;
              overlay = {
                ...overlay,
                dataBar: { color: rule.barColor ?? "#60a5fa", ratio: (numeric - boundedMin) / span },
              };
            }
            break;
            }
            case "colorScale": {
            const numeric = numericOf(cell);
            if (numeric !== undefined) {
              matches = true;
              const ratio = (numeric - boundedMin) / span;
              const minColor = rule.minColor ?? "#fca5a5";
              const midColor = rule.midColor;
              const maxColor = rule.maxColor ?? "#86efac";
              overlay = {
                ...overlay,
                colorScale: midColor
                  ? (ratio <= 0.5 ? mixHex(minColor, midColor, ratio * 2) : mixHex(midColor, maxColor, (ratio - 0.5) * 2))
                  : mixHex(minColor, maxColor, ratio),
              };
            }
            break;
            }
            case "iconSet": {
            const numeric = numericOf(cell);
            if (numeric !== undefined) {
              matches = true;
              const ratio = (numeric - boundedMin) / span;
              overlay = {
                ...overlay,
                icon: ratio >= 0.67 ? "up" : ratio >= 0.34 ? "flat" : "down",
              };
            }
            break;
            }
            case "topBottom": {
              matches = matchesTopBottom(rule, cell, sheet, range);
              if (matches && rule.style) overlay = { ...overlay, style: { ...overlay.style, ...rule.style } };
              break;
            }
          }
          if (matches || Object.keys(overlay).length > 0) overlays.set(key, overlay);
          if (matches && rule.stopIfTrue) stopped.add(key);
        }
      }
    }
  }
  return overlays;
}

function matchesTopBottom(
  rule: ConditionalFormatRule,
  cell: CellData | undefined,
  sheet: WorksheetModel,
  range: RangeRef,
): boolean {
  const config = rule.topBottom ?? {
    direction: rule.operator === 'bottom' ? 'bottom' : 'top',
    rank: Number(rule.value1 ?? 10),
  };
  const values: Array<{ row: number; column: number; value: number }> = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const value = numericOf(sheet.cells.get(row, column));
      if (value !== undefined) values.push({ row, column, value });
    }
  }
  if (values.length === 0 || !cell) return false;
  values.sort((left, right) => left.value - right.value);
  const rank = config.percent
    ? Math.max(1, Math.ceil(values.length * config.rank / 100))
    : Math.max(1, Math.floor(config.rank));
  const threshold = config.direction === 'top'
    ? values[Math.max(0, values.length - rank)]?.value
    : values[Math.min(values.length - 1, rank - 1)]?.value;
  const numeric = numericOf(cell);
  if (numeric === undefined || threshold === undefined) return false;
  return config.direction === 'top' ? numeric >= threshold : numeric <= threshold;
}

function evaluateHighlight(
  rule: ConditionalFormatRule,
  cell: CellData | undefined,
  sheet: WorksheetModel,
  row: number,
  column: number,
  valueCounts: Map<string, number>,
): boolean {
  const text = cellText(cell);
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
    case "formula": return evaluateCfFormula(String(firstValue ?? ""), sheet, row, column, cell);
    default: return false;
  }
}

function evaluateCfFormula(formula: string, sheet: WorksheetModel, row: number, column: number, cell: CellData | undefined): boolean {
  const source = formula.trim();
  if (!source) return false;
  try {
    const ast = parseFormula(source.startsWith('=') ? source : `=${source}`);
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
          const text = cellText(sheet.cells.get(r, c));
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

function normalizeCriterion(criterion: FilterCriterion): FilterCriterion {
  if (criterion.kind === 'values') return { ...structuredClone(criterion), values: [...new Set(criterion.values)] };
  if (criterion.kind === 'custom') {
    if (!criterion.conditions[0]) throw new Error('Custom filter requires a condition');
    return structuredClone(criterion);
  }
  if (criterion.kind === 'top10' && (!Number.isSafeInteger(criterion.rank) || criterion.rank <= 0)) throw new Error('Top10 filter rank must be positive');
  return structuredClone(criterion);
}

export function computeFilterHiddenRows(sheet: WorksheetModel): Set<number> {
  const hidden = new Set<number>();
  let filter: import('@react-sheets/core-model').AutoFilterModel | undefined;
  try {
    const source = resolveWorksheetAutoFilter(sheet);
    filter = source ? normalizeAutoFilterModel(source) : undefined;
  }
  catch {
    // Malformed filter state must not expose unfiltered data accidentally.
    for (let row = 0; row < sheet.rowCount; row += 1) hidden.add(row);
    return hidden;
  }
  if (!filter) return hidden;
  const table = sheet.sheetTables.find((entry) => entry.sheetId === sheet.id
    && entry.range.startRow === filter.range.startRow
    && entry.range.endRow === filter.range.endRow
    && entry.range.startColumn === filter.range.startColumn
    && entry.range.endColumn === filter.range.endColumn);
  const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
  for (const rawKey of Object.keys(filter.columns)) {
    const column = Number(rawKey);
    const criterion = filter.columns[column]?.criterion;
    if (!criterion) continue;
    for (let row = filter.range.startRow + 1; row <= endRow; row++) {
      const cell = sheet.cells.get(row, column);
      const text = cellText(cell);
      const visible = matchesFilterCriterion(cell?.value ?? null, text, criterion);
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
export function getAutoFilterValueDomain(sheet: WorksheetModel, column: number): string[] {
  const source = resolveWorksheetAutoFilter(sheet);
  const filter = source ? normalizeAutoFilterModel(source) : undefined;
  if (!filter || column < filter.range.startColumn || column > filter.range.endColumn) return [];
  const values = new Set<string>();
  for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
    const otherColumnsMatch = Object.values(filter.columns).every((entry) => {
      if (entry.column === column || !entry.criterion) return true;
      const cell = sheet.cells.get(row, entry.column);
      return matchesFilterCriterion(cell?.value ?? null, cellText(cell), entry.criterion);
    });
    if (!otherColumnsMatch) continue;
    const cell = sheet.cells.get(row, column);
    values.add(cellText(cell));
  }
  return [...values].sort((left, right) => left.localeCompare(right));
}

function matchesFilterCriterion(value: unknown, text: string, criterion: FilterCriterion): boolean {
  if (criterion.kind === 'values') {
    if (text === '' && criterion.includeBlank) return true;
    return criterion.values.some((candidate) => String(candidate ?? '').toLocaleLowerCase() === text.toLocaleLowerCase());
  }
  if (criterion.kind === 'custom') {
    const results = criterion.conditions.filter((condition): condition is NonNullable<typeof condition> => Boolean(condition))
      .map((condition) => evaluateFilterCondition(text, condition.operator, String(condition.value ?? '')));
    return criterion.join === 'and' ? results.every(Boolean) : results.some(Boolean);
  }
  if (criterion.kind === 'top10' || criterion.kind === 'dynamic' || criterion.kind === 'color' || criterion.kind === 'icon') {
    // Ranking, dynamic date and visual criteria are resolved by the typed index
    // before row projection. They are intentionally not treated as value filters.
    void value;
    return true;
  }
  return true;
}

function evaluateFilterCondition(text: string, operator: string, operand: string, operand2?: string): boolean {
  const normalizedOperator = operator.trim().toLocaleLowerCase();
  const normalizedText = text.trim().toLocaleLowerCase();
  const numeric = Number(text.replace(/[$,%\s,]/g, ""));
  const operandNumeric = Number(operand.replace(/[$,%\s,]/g, ""));
  const hasNumbers = Number.isFinite(numeric) && Number.isFinite(operandNumeric);
  switch (normalizedOperator) {
    case ">": return hasNumbers && numeric > operandNumeric;
    case "<": return hasNumbers && numeric < operandNumeric;
    case ">=": return hasNumbers && numeric >= operandNumeric;
    case "<=": return hasNumbers && numeric <= operandNumeric;
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
    const evaluated = evaluateValidationFormula(formula, sheet, 0, 0, undefined);
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
): FormulaValue {
  try {
    const ast = parseFormula(formula.trim().startsWith('=') ? formula.trim() : `=${formula.trim()}`);
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
  if (value == null || value === "") {
    const valid = Boolean(rule.allowBlank ?? true);
    return { valid, blocking: !valid && (rule.alertStyle ?? 'stop') === 'stop', message: valid ? undefined : validationMessage(rule, "该单元格不允许为空") };
  }
  const list = validationList(rule, sheet);
  if (list) {
    const candidateValues = rule.multiSelect ? String(value).split(',').map((item) => item.trim()).filter(Boolean) : [String(value)];
    const ok = candidateValues.length > 0 && candidateValues.every((candidate) =>
      list.some((item) => item.toLowerCase() === candidate.toLowerCase()));
    return {
      valid: ok,
      blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop',
      message: ok ? undefined : validationMessage(rule, "值不在允许的列表中"),
      list,
    };
  }
  const numeric = typeof value === "number" ? value : Number(value);
  const isNumberType = rule.type === "whole" || rule.type === "decimal";
  if (isNumberType) {
    if (!Number.isFinite(numeric)) {
      return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "需要输入数字") };
    }
    if (rule.type === "whole" && !Number.isInteger(numeric)) {
      return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "需要输入整数") };
    }
    const bound1 = Number(rule.formula1);
    const bound2 = Number(rule.formula2);
    switch (rule.operator) {
      case "greaterThan": return judge(Number.isFinite(bound1) && numeric > bound1, rule);
      case "lessThan": return judge(Number.isFinite(bound1) && numeric < bound1, rule);
      case "equal": return judge(numeric === bound1, rule);
      case "notEqual": return judge(numeric !== bound1, rule);
      case "notBetween": return judge(Number.isFinite(bound1) && Number.isFinite(bound2) && (numeric < bound1 || numeric > bound2), rule);
      case "between":
      default: return judge(!Number.isFinite(bound1) || (!Number.isFinite(bound2) ? numeric >= bound1 : numeric >= bound1 && numeric <= bound2), rule);
    }
  }
  if (rule.type === "textLength") {
    const length = String(value).length;
    const bound1 = Number(rule.formula1);
    const bound2 = Number(rule.formula2);
    switch (rule.operator) {
      case "greaterThan": return judge(length > bound1, rule);
      case "lessThan": return judge(length < bound1, rule);
      case "equal": return judge(length === bound1, rule);
      default: return judge(!(Number.isFinite(bound1) && length < bound1) && !(Number.isFinite(bound2) && length > bound2), rule);
    }
  }
  if (rule.type === "date" || rule.type === "time") {
    const validDate = rule.type === 'date' ? isValidDateValue(value) : isValidTimeValue(value);
    if (!validDate) {
      return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "需要输入有效日期/时间") };
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
      return judge(ok, rule);
    }
  }
  if (rule.type === "checkbox") {
    const ok = value === true || value === false || String(value).toUpperCase() === "TRUE" || String(value).toUpperCase() === "FALSE";
    return { valid: ok, blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop', message: ok ? undefined : validationMessage(rule, "需要 TRUE/FALSE") };
  }
  if (rule.type === "custom") {
    if (!rule.formula1) return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: validationMessage(rule, "自定义验证公式缺失") };
    const evaluated = evaluateValidationFormula(rule.formula1, sheet, row, column, value);
    const ok = evaluated === true || (typeof evaluated === 'number' && evaluated !== 0) || (typeof evaluated === 'string' && evaluated.length > 0);
    return judge(ok, rule);
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

// ---------- 查找替换 ----------

export interface FindReplaceParams {
  find: string;
  replace: string;
  matchCase: boolean;
  entireCell: boolean;
  scope: "sheet" | "workbook";
}

export interface RangeValuesPatch {
  sheetId: string;
  startRow: number;
  startColumn: number;
  values: CellData[][];
}

/** 收集需要替换的区域补丁(每个命中行生成一个连续区段补丁) */
export function collectFindReplacements(
  workbook: WorkbookModel,
  params: FindReplaceParams,
): RangeValuesPatch[] {
  if (!params.find) return [];
  const sheets = params.scope === "workbook"
    ? workbook.getSheets()
    : [workbook.getSheet(workbook.primarySheetId)];
  const patches: RangeValuesPatch[] = [];

  for (const sheet of sheets) {
    const byRow = new Map<number, Map<number, CellData>>();
    sheet.cells.forEach((cell, row, column) => {
      const original = cellText(cell);
      if (!original) return;
      const haystack = params.matchCase ? original : original.toLowerCase();
      const needle = params.matchCase ? params.find : params.find.toLowerCase();
      const hit = params.entireCell ? haystack === needle : haystack.includes(needle);
      if (!hit) return;
      const replaced = params.entireCell
        ? params.replace
        : (params.matchCase ? original : original)
            .split(params.find)
            .join(params.replace);
      const next: CellData = { ...cell, value: coerceLike(replaced, cell.value) };
      delete next.displayValue;
      let rowMap = byRow.get(row);
      if (!rowMap) { rowMap = new Map(); byRow.set(row, rowMap); }
      rowMap.set(column, next);
    });

    for (const [row, columns] of byRow) {
      const columnList = [...columns.keys()].sort((a, b) => a - b);
      const startColumn = columnList[0]!;
      const endColumn = columnList.at(-1)!;
      const values: CellData[][] = [[]];
      for (let c = startColumn; c <= endColumn; c++) {
        values[0]!.push(columns.get(c) ?? { value: null });
      }
      patches.push({ sheetId: sheet.id, startRow: row, startColumn, values });
    }
  }
  return patches;
}

function coerceLike(text: string, previous: CellData["value"]): CellData["value"] {
  if (typeof previous === "number") {
    const numeric = Number(text.replace(/[$,%]/g, ""));
    if (Number.isFinite(numeric) && text !== "") return numeric;
  }
  return text;
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
}

export interface SubtotalParams {
  sheetId: string;
  range: RangeRef;
  groupColumn: number;
  valueColumn: number;
  functionName: 'SUM' | 'COUNT' | 'AVERAGE';
}

function compareSortValues(left: CellData['value'], right: CellData['value']): number {
  if (left === right) return 0;
  if (left == null) return 1;
  if (right == null) return -1;
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  if (typeof left === 'boolean' && typeof right === 'boolean') return Number(left) - Number(right);
  return String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: 'base' });
}

function sortedSourceRows(sheet: WorksheetModel, params: DataSortParams): number[] {
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
        sheet.cells.get(leftRow, criterion.column)?.value ?? null,
        sheet.cells.get(rightRow, criterion.column)?.value ?? null,
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
  if (sheet.notes.size > 0 || sheet.commentThreads.some((thread) => thread.row >= range.startRow && thread.row <= range.endRow)) {
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
  let key = cellText(sheet.cells.get(start, params.groupColumn));
  for (let row = start + 1; row <= range.endRow + 1; row += 1) {
    const next = row <= range.endRow ? cellText(sheet.cells.get(row, params.groupColumn)) : undefined;
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
      validatePermutationMetadata(sheet, range);
      applyRowPermutation(sheet, { range, sourceRows: params.sourceRows });
      setAppliedSortState(sheet, params.sortState);
    },
    metadata: {
      schema: { name: 'RowsPermuted', validate: isRowsPermutedMutation },
      permission: { capability: 'sheet.sort.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['rows.permuted'],
    },
  });

  runtime.registry.registerCommand<DataSortParams>({
    id: 'data.sort.rows',
    execute: (params, context) => {
      const range = selectedRange(params);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertNoDataRegionIntersection(sheet, range, 'Sort');
      if (params.criteria.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const sourceRows = sortedSourceRows(sheet, { ...params, range });
      if (sourceRows.length <= 1 || sourceRows.every((row, offset) => row === range.startRow + (params.hasHeader ? 1 : 0) + offset)) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const startRow = (params.hasHeader ?? false) ? range.startRow + 1 : range.startRow;
      const bodyRange: RangeRef = { ...range, startRow };
      const inverseRows = new Array<number>(sourceRows.length);
      sourceRows.forEach((sourceRow, offset) => { inverseRows[sourceRow - startRow] = startRow + offset; });
      const affectedRanges = [structuredClone(bodyRange)];
      context.applyMutation({
        id: 'rows.permuted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: {
          ...params,
          range: bodyRange,
          sourceRows,
          sortState: {
            sheetId: params.sheetId,
            range,
            criteria: structuredClone(params.criteria),
            hasHeader: params.hasHeader,
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
            range: bodyRange,
            sourceRows: inverseRows,
            sortState: ((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState
              ? structuredClone((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState)
              : undefined),
            previousSortState: ((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState
              ? structuredClone((sheet as WorksheetModel & { appliedSortState?: AppliedSortState }).appliedSortState)
              : undefined),
          },
          affectedRanges,
        }],
        apply: () => applyRowPermutation(sheet, { range: bodyRange, sourceRows }),
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
        const key = params.columns.map((column) => cellText(sheet.cells.get(row, column))).join('\u0001');
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
