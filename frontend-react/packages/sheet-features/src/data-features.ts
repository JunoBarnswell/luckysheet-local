import type {
  CellData,
  ConditionalFormatRule,
  DataValidationRule,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
} from "@react-sheets/core-model";

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
  const rules = sheet.conditionalFormats;
  if (rules.length === 0) return overlays;

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
          let overlay = overlays.get(key) ?? {};

          switch (rule.type) {
            case "highlight": {
              const matches = evaluateHighlight(rule, cell);
              if (matches && rule.style) overlay = { ...overlay, style: { ...overlay.style, ...rule.style } };
              break;
            }
            case "dataBar": {
            const numeric = numericOf(cell);
            if (numeric !== undefined) {
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
              const ratio = (numeric - boundedMin) / span;
              overlay = {
                ...overlay,
                icon: ratio >= 0.67 ? "up" : ratio >= 0.34 ? "flat" : "down",
              };
            }
            break;
            }
          }
          overlays.set(key, overlay);
        }
      }
    }
  }
  return overlays;
}

function evaluateHighlight(rule: ConditionalFormatRule, cell: CellData | undefined): boolean {
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
    case "duplicate": return false; // 单规则内无法判定,需跨范围统计(留待增强)
    case "unique": return false;
    case "formula": return false;
    default: return false;
  }
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
      if (visible && criterion.conditionOperator && criterion.conditionValue != null) {
        visible = evaluateFilterCondition(text, criterion.conditionOperator, criterion.conditionValue);
      }
      if (!visible) hidden.add(row);
    }
  }
  return hidden;
}

function evaluateFilterCondition(text: string, operator: string, operand: string): boolean {
  const numeric = Number(text.replace(/[$,%]/g, ""));
  const operandNumeric = Number(operand);
  const hasNumbers = Number.isFinite(numeric) && Number.isFinite(operandNumeric);
  switch (operator) {
    case ">": return hasNumbers && numeric > operandNumeric;
    case "<": return hasNumbers && numeric < operandNumeric;
    case ">=": return hasNumbers && numeric >= operandNumeric;
    case "<=": return hasNumbers && numeric <= operandNumeric;
    case "=": return text.toLowerCase() === operand.toLowerCase();
    case "contains": return text.toLowerCase().includes(operand.toLowerCase());
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

export function validationList(rule: DataValidationRule): string[] | undefined {
  if (rule.type !== "list" || !rule.formula1) return undefined;
  return rule.formula1.split(",").map((item) => item.trim().replace(/^"|"$/g, "")).filter(Boolean);
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
    return { valid: Boolean(rule.allowBlank ?? true), blocking: !(rule.allowBlank ?? true), message: rule.errorMessage ?? "该单元格不允许为空" };
  }
  const list = validationList(rule);
  if (list) {
    const ok = list.some((item) => item.toLowerCase() === String(value).toLowerCase());
    return { valid: ok, blocking: !ok && Boolean(rule.errorMessage), message: rule.errorMessage ?? "值不在允许的列表中", list };
  }
  const numeric = typeof value === "number" ? value : Number(value);
  const isNumberType = rule.type === "whole" || rule.type === "decimal";
  if (isNumberType) {
    if (!Number.isFinite(numeric)) {
      return { valid: false, blocking: Boolean(rule.errorMessage), message: rule.errorMessage ?? "需要输入数字" };
    }
    if (rule.type === "whole" && !Number.isInteger(numeric)) {
      return { valid: false, blocking: Boolean(rule.errorMessage), message: rule.errorMessage ?? "需要输入整数" };
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
  if (rule.type === "date") {
    const date = new Date(String(value));
    if (Number.isNaN(date.getTime())) {
      return { valid: false, blocking: Boolean(rule.errorMessage), message: rule.errorMessage ?? "需要输入有效日期" };
    }
  }
  return { valid: true, blocking: false };
}

function judge(ok: boolean, rule: DataValidationRule): DataValidationResult {
  return {
    valid: ok,
    blocking: !ok && Boolean(rule.errorMessage),
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
