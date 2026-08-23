import type {
  CellData,
  ConditionalFormatRule,
  DataValidationRule,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
  ConditionalFormatTopBottom,
} from "@react-sheets/core-model";
import { applyRowPermutation, columnLabel, validatePermutationMetadata } from "@react-sheets/core-model";
import type { CommandRuntime } from "@react-sheets/command-runtime";
import { formatFormula, mapAstReferences, parseFormula, type ParsedCellReference } from "@react-sheets/formula-engine";

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

function inRange(range: RangeRef, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow
    && column >= range.startColumn && column <= range.endColumn;
}

export function normalizeConditionalFormatRule(
  rule: ConditionalFormatRule,
  fallbackPriority = 1,
): ConditionalFormatRule {
  const ranges = rule.ranges.map(normalizeRangeRef);
  if (ranges.length === 0) throw new Error(`Conditional format ${rule.id} requires at least one range`);
  if (!rule.id.trim()) throw new Error('Conditional format id is required');
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
  if (rule.type === 'checkbox' && rule.operator !== undefined) {
    throw new Error('Checkbox validation does not accept a comparison operator');
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
  const v1 = rule.value1;
  const n1 = typeof v1 === "number" ? v1 : Number(v1);
  switch (rule.operator) {
    case "greaterThan": return numeric !== undefined && Number.isFinite(n1) && numeric > n1;
    case "lessThan": return numeric !== undefined && Number.isFinite(n1) && numeric < n1;
    case "equal":
      return typeof v1 === "string" ? text.toLowerCase() === String(v1).toLowerCase() : numeric === n1;
    case "notEqual":
      return typeof v1 === "string" ? text.toLowerCase() !== String(v1).toLowerCase() : numeric !== n1;
    case "between": {
      const n2 = typeof rule.value2 === "number" ? rule.value2 : Number(rule.value2);
      return numeric !== undefined && Number.isFinite(n1) && Number.isFinite(n2) && numeric >= n1 && numeric <= n2;
    }
    case "containsText": return typeof v1 === "string" && text.toLowerCase().includes(String(v1).toLowerCase());
    case "notContainsText": return typeof v1 === "string" && !text.toLowerCase().includes(String(v1).toLowerCase());
    case "duplicate": return valueCounts.get(text) !== undefined && (valueCounts.get(text) ?? 0) > 1;
    case "unique": return text !== "" && (valueCounts.get(text) ?? 0) === 1;
    case "formula": return evaluateCfFormula(String(v1 ?? ""), sheet, row, column, cell);
    default: return false;
  }
}

function evaluateCfFormula(formula: string, sheet: WorksheetModel, row: number, column: number, cell: CellData | undefined): boolean {
  const source = formula.trim().startsWith("=") ? formula.trim().slice(1) : formula.trim();
  if (!source) return false;
  const replaced = source
    .replace(/\bROW\(\)/gi, String(row + 1))
    .replace(/\bCOLUMN\(\)/gi, String(column + 1))
    .replace(/\bROW\b/gi, String(row + 1))
    .replace(/\bCOLUMN\b/gi, String(column + 1));
  const cellRef = replaced.match(/^([A-Z]+\d+)\s*(=|<>|>=|<=|>|<)\s*(.+)$/i);
  if (cellRef) {
    const ref = parseSimpleA1(cellRef[1]!);
    const operator = cellRef[2]!;
    const operand = cellRef[3]!.trim();
    const target = ref ? sheet.cells.get(ref.row, ref.column) : cell;
    const left = numericOf(target) ?? cellText(target);
    const right = Number(operand);
    if (Number.isFinite(right)) return compareValues(left, right, operator);
    return compareValues(String(left).toLowerCase(), operand.replace(/^"|"$/g, "").toLowerCase(), operator);
  }
  if (/^MOD\s*\(/i.test(replaced)) {
    const modMatch = replaced.match(/^MOD\s*\(\s*ROW\s*\(\s*\)\s*,\s*(\d+)\s*\)\s*=\s*(\d+)/i);
    if (modMatch) return (row + 1) % Number(modMatch[1]) === Number(modMatch[2]);
  }
  return Boolean(cell?.value);
}

function parseSimpleA1(reference: string): { row: number; column: number } | null {
  const match = reference.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  let column = 0;
  for (const char of match[1]!.toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function compareValues(left: string | number, right: string | number, operator: string): boolean {
  switch (operator) {
    case "=": return left === right;
    case "<>": return left !== right;
    case ">": return Number(left) > Number(right);
    case "<": return Number(left) < Number(right);
    case ">=": return Number(left) >= Number(right);
    case "<=": return Number(left) <= Number(right);
    default: return false;
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

export function computeFilterHiddenRows(sheet: WorksheetModel): Set<number> {
  const hidden = new Set<number>();
  const filter = sheet.filter;
  if (!filter) return hidden;
  for (const rawKey of Object.keys(filter.criteria)) {
    const column = Number(rawKey);
    const criterion = filter.criteria[column];
    if (!criterion) continue;
    for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row++) {
      const cell = sheet.cells.get(row, column);
      const text = cellText(cell);
      let visible = true;
      if (criterion.selectedValues != null) {
        visible = criterion.selectedValues.includes(text);
      }
      if (visible && criterion.excludeBlanks && text === "") {
        visible = false;
      }
      if (visible && criterion.conditionOperator && criterion.conditionValue != null) {
        visible = evaluateFilterCondition(text, criterion.conditionOperator, criterion.conditionValue, criterion.conditionValue2);
      }
      if (!visible) hidden.add(row);
    }
  }
  return hidden;
}

function evaluateFilterCondition(text: string, operator: string, operand: string, operand2?: string): boolean {
  const numeric = Number(text.replace(/[$,%]/g, ""));
  const operandNumeric = Number(operand);
  const hasNumbers = Number.isFinite(numeric) && Number.isFinite(operandNumeric);
  switch (operator) {
    case ">": return hasNumbers && numeric > operandNumeric;
    case "<": return hasNumbers && numeric < operandNumeric;
    case ">=": return hasNumbers && numeric >= operandNumeric;
    case "<=": return hasNumbers && numeric <= operandNumeric;
    case "=": return text.toLowerCase() === operand.toLowerCase();
    case "<>": return text.toLowerCase() !== operand.toLowerCase();
    case "contains": return text.toLowerCase().includes(operand.toLowerCase());
    case "notContains": return !text.toLowerCase().includes(operand.toLowerCase());
    case "beginsWith": return text.toLowerCase().startsWith(operand.toLowerCase());
    case "endsWith": return text.toLowerCase().endsWith(operand.toLowerCase());
    case "between": {
      const upper = Number(operand2);
      return hasNumbers && Number.isFinite(upper) && numeric >= operandNumeric && numeric <= upper;
    }
    default: return true;
  }
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
  return formula.split(",").map((item) => item.trim().replace(/^"|"$/g, "")).filter(Boolean);
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
    return { valid, blocking: !valid && (rule.alertStyle ?? 'stop') === 'stop', message: valid ? undefined : rule.errorMessage ?? "该单元格不允许为空" };
  }
  const list = validationList(rule, sheet);
  if (list) {
    const candidateValues = rule.multiSelect ? String(value).split(',').map((item) => item.trim()).filter(Boolean) : [String(value)];
    const ok = candidateValues.length > 0 && candidateValues.every((candidate) =>
      list.some((item) => item.toLowerCase() === candidate.toLowerCase()));
    return {
      valid: ok,
      blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop',
      message: ok ? undefined : rule.errorMessage ?? "值不在允许的列表中",
      list,
    };
  }
  const numeric = typeof value === "number" ? value : Number(value);
  const isNumberType = rule.type === "whole" || rule.type === "decimal";
  if (isNumberType) {
    if (!Number.isFinite(numeric)) {
      return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: rule.errorMessage ?? "需要输入数字" };
    }
    if (rule.type === "whole" && !Number.isInteger(numeric)) {
      return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: rule.errorMessage ?? "需要输入整数" };
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
    const date = new Date(String(value));
    if (Number.isNaN(date.getTime())) {
      return { valid: false, blocking: (rule.alertStyle ?? 'stop') === 'stop', message: rule.errorMessage ?? "需要输入有效日期/时间" };
    }
  }
  if (rule.type === "checkbox") {
    const ok = value === true || value === false || String(value).toUpperCase() === "TRUE" || String(value).toUpperCase() === "FALSE";
    return { valid: ok, blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop', message: ok ? undefined : rule.errorMessage ?? "需要 TRUE/FALSE" };
  }
  return { valid: true, blocking: false };
}

function judge(ok: boolean, rule: DataValidationRule): DataValidationResult {
  return {
    valid: ok,
    blocking: !ok && (rule.alertStyle ?? 'stop') === 'stop',
    message: ok ? undefined : rule.errorMessage ?? "不符合数据验证规则",
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
    : [workbook.getSheet(workbook.activeSheetId)];
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
  if (sheet.filter && intersects(sheet.filter.range)) throw new Error('Matrix transform cannot rewrite a filtered range');
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
  params: { sheetId: string; range: RangeRef; direction?: 'horizontal' | 'vertical' },
  transpose: boolean,
): ReturnType<CommandRuntime['execute']> {
  const sheet = runtime.workbook.getSheet(params.sheetId);
  const range = normalizeRangeRef({ ...params.range, sheetId: params.sheetId });
  assertMatrixTransformSupported(sheet, range);
  const target = matrixTargetRange(range, transpose);
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
  runtime.execute('sheet.range.clear', { sheetId: params.sheetId, range: clearRange, mode: 'all' });
  return runtime.execute('sheet.range.set', {
    sheetId: params.sheetId,
    startRow: target.startRow,
    startColumn: target.startColumn,
    values,
  });
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
  runtime.registry.registerMutation('rows.permuted', (item, context) => {
    const params = item.params as DataSortParams & { sourceRows: number[] };
    const range = selectedRange(params);
    const sheet = context.workbook.getSheet(params.sheetId);
    validatePermutationMetadata(sheet, range);
    applyRowPermutation(sheet, { range, sourceRows: params.sourceRows });
  });

  runtime.registry.registerCommand<DataSortParams>({
    id: 'data.sort.rows',
    execute: (params, context) => {
      const range = selectedRange(params);
      const sheet = context.workbook.getSheet(params.sheetId);
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
        params: { ...params, range: bodyRange, sourceRows },
        affectedRanges,
        inverse: [{
          id: 'rows.permuted',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { ...params, range: bodyRange, sourceRows: inverseRows },
          affectedRanges,
        }],
        apply: () => applyRowPermutation(sheet, { range: bodyRange, sourceRows }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; range: RangeRef }>({
    id: 'matrix.transpose',
    execute: (params) => executeMatrixTransform(runtime, params, true),
  });

  runtime.registry.registerCommand<{ sheetId: string; range: RangeRef; direction: 'horizontal' | 'vertical' }>({
    id: 'matrix.flip',
    execute: (params) => executeMatrixTransform(runtime, params, false),
  });

  runtime.registry.registerCommand<TextToColumnsParams>({
    id: 'data.textToColumns',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const range = params.range;
      const maxColumns = Math.max(2, params.maxColumns ?? 8);
      const values: CellData[][] = [];
      for (let row = range.startRow; row <= range.endRow; row++) {
        const cell = sheet.cells.get(row, range.startColumn);
        const text = cell?.value == null ? '' : String(cell.value);
        const parts = text.split(params.delimiter).slice(0, maxColumns);
        values.push(parts.map((part) => ({ value: part })));
      }
      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: range.startRow,
        startColumn: range.startColumn,
        values,
      });
    },
  });

  runtime.registry.registerCommand<RemoveDuplicatesParams>({
    id: 'data.removeDuplicates',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const range = selectedRange(params);
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
        const result = runtime.execute('sheet.rows.delete', { sheetId: params.sheetId, at: run.at, count: run.count });
        mutationCount += result.mutationCount;
        affectedRanges.push({ sheetId: params.sheetId, startRow: run.at, endRow: run.at + run.count - 1, startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) });
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });

  runtime.registry.registerCommand<SubtotalParams>({
    id: 'data.subtotal',
    execute: (params, context) => {
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
        const result = runtime.execute('sheet.rows.insert', { sheetId: params.sheetId, at: targetRow, count: groups.length + 1 });
        mutationCount += result.mutationCount;
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
      const setResult = runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: targetRow,
        startColumn: range.startColumn,
        values,
      });
      mutationCount += setResult.mutationCount;
      const affectedRanges: RangeRef[] = [structuredClone(range), { sheetId: params.sheetId, startRow: targetRow, endRow: targetRow + groups.length, startColumn: range.startColumn, endColumn: range.endColumn }];
      for (const group of groups) {
        runtime.execute('outline.group.add', {
          sheetId: params.sheetId,
          group: { id: `subtotal-${context.operationId}-${group.start}-${group.end}`, axis: 'row', start: group.start, end: group.end, level: 1, collapsed: false },
        });
        mutationCount += 1;
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; row: number; column: number; delimiter: string; maxColumns?: number }>({
    id: 'data.splitColumn',
    execute: (params, context) => runtime.execute('sheet.splitColumn', params),
  });

  runtime.registry.registerCommand<{ name: string }>({
    id: 'workbook.name.list',
    execute: (_params, context) => {
      const names = context.workbook.listDefinedNames().map((entry) => entry.name).sort();
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: names.map(() => ({ sheetId: context.workbook.activeSheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })) };
    },
  });
}
