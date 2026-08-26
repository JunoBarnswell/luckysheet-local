import type { FormulaValue } from "../values";
import { createFormulaError, isFormulaError } from "../values";
import { matchesCriteria, parseCriteria, sameCriteriaShape, toCriteriaRange, type CriteriaExpression, type CriteriaRange } from "../criteria";
import type { RangeDependency, FormulaDependency } from "../range-index";
import type { CellAddress } from "../ast";

export interface AdvancedFunctionArgs {
  values: FormulaValue[];
  ranges: unknown[]; // EvaluationValue[](含 range 包装)
}

export interface AdvancedContext {
  currentCell: CellAddress;
  readMatrix(range: RangeDependency): FormulaValue[][];
  toRange(value: unknown): RangeDependency | undefined;
}

type AdvancedFn = (args: AdvancedFunctionArgs, context: AdvancedContext) => FormulaValue | undefined;

function flattenMatrix(value: unknown): FormulaValue[] {
  if (!Array.isArray(value)) return [(value ?? null) as FormulaValue];
  const out: FormulaValue[] = [];
  for (const row of value as unknown[]) {
    if (Array.isArray(row)) for (const cell of row) out.push((cell ?? null) as FormulaValue);
    else out.push((row ?? null) as FormulaValue);
  }
  return out;
}

function isMatrix(value: FormulaValue | undefined): boolean {
  return Array.isArray(value);
}

function numbersOf(values: FormulaValue[]): number[] {
  const out: number[] = [];
  for (const value of values) {
    if (typeof value === "number") out.push(value);
    else if (typeof value === "boolean") out.push(value ? 1 : 0);
    else if (typeof value === "string") {
      const parsed = Number(value.replace(/[$,%]/g, ""));
      if (Number.isFinite(parsed) && value.trim() !== "") out.push(parsed);
    }
  }
  return out;
}

interface CriteriaPair { range: CriteriaRange; criteria: CriteriaExpression }

function collectCriteriaPairs(args: AdvancedFunctionArgs, startIndex: number): CriteriaPair[] {
  const pairs: CriteriaPair[] = [];
  for (let i = startIndex; i + 1 < args.values.length; i += 2) {
    pairs.push({
      range: toCriteriaRange(args.values[i] ?? null),
      criteria: parseCriteria(args.values[i + 1] ?? null),
    });
  }
  return pairs;
}

function rowsPass(pairs: CriteriaPair[]): boolean[] | ReturnType<typeof createFormulaError> {
  if (pairs.length === 0) return createFormulaError('#VALUE!', 'At least one criteria range is required');
  const shape = pairs[0]!.range;
  if (shape.columns < 0 || pairs.some((pair) => !sameCriteriaShape(shape, pair.range))) {
    return createFormulaError('#VALUE!', 'Criteria ranges must have identical shape');
  }
  const pass: boolean[] = new Array(shape.rows * shape.columns).fill(true);
  for (let row = 0; row < shape.rows; row += 1) {
    for (let column = 0; column < shape.columns; column += 1) {
      const index = row * shape.columns + column;
      for (const pair of pairs) {
        if (!matchesCriteria(pair.range.values[row]![column]!, pair.criteria)) pass[index] = false;
      }
    }
  }
  return pass;
}

function targetWithPass(target: CriteriaRange, pairs: CriteriaPair[], pass: boolean[] | ReturnType<typeof createFormulaError>): FormulaValue | undefined {
  if (isFormulaError(pass)) return pass;
  if (target.columns < 0 || pairs.length === 0 || !sameCriteriaShape(target, pairs[0]!.range)) {
    return createFormulaError('#VALUE!', 'Target and criteria ranges must have identical shape');
  }
  return undefined;
}

export const ADVANCED_FUNCTIONS: Record<string, AdvancedFn> = {
  SUMIFS: (args) => {
    const sumRange = toCriteriaRange(args.values[0] ?? null);
    const pairs = collectCriteriaPairs(args, 1);
    const pass = rowsPass(pairs);
    const shapeError = targetWithPass(sumRange, pairs, pass);
    if (shapeError !== undefined) return shapeError;
    if (isFormulaError(pass)) return pass;
    let total = 0;
    for (let row = 0; row < sumRange.rows; row += 1) {
      for (let column = 0; column < sumRange.columns; column += 1) {
        const index = row * sumRange.columns + column;
        const value = sumRange.values[row]![column]!;
        if (pass[index] && typeof value === "number") total += value;
      }
    }
    return total;
  },
  COUNTIFS: (args) => {
    const pairs = collectCriteriaPairs(args, 0);
    const pass = rowsPass(pairs);
    return isFormulaError(pass) ? pass : pass.filter(Boolean).length;
  },
  AVERAGEIFS: (args) => {
    const avgRange = toCriteriaRange(args.values[0] ?? null);
    const pairs = collectCriteriaPairs(args, 1);
    const pass = rowsPass(pairs);
    const shapeError = targetWithPass(avgRange, pairs, pass);
    if (shapeError !== undefined) return shapeError;
    if (isFormulaError(pass)) return pass;
    let total = 0;
    let count = 0;
    for (let row = 0; row < avgRange.rows; row += 1) {
      for (let column = 0; column < avgRange.columns; column += 1) {
        const index = row * avgRange.columns + column;
        const value = avgRange.values[row]![column]!;
        if (pass[index] && typeof value === "number") { total += value; count++; }
      }
    }
    return count === 0 ? createFormulaError("#DIV/0!", "No matching values") : total / count;
  },
  MAXIFS: (args) => {
    const target = toCriteriaRange(args.values[0] ?? null);
    const pairs = collectCriteriaPairs(args, 1);
    const pass = rowsPass(pairs);
    const shapeError = targetWithPass(target, pairs, pass);
    if (shapeError !== undefined) return shapeError;
    if (isFormulaError(pass)) return pass;
    let max = -Infinity;
    for (let row = 0; row < target.rows; row += 1) {
      for (let column = 0; column < target.columns; column += 1) {
        const index = row * target.columns + column;
        const value = target.values[row]![column]!;
        if (pass[index] && typeof value === "number") max = Math.max(max, value);
      }
    }
    return max === -Infinity ? 0 : max;
  },
  MINIFS: (args) => {
    const target = toCriteriaRange(args.values[0] ?? null);
    const pairs = collectCriteriaPairs(args, 1);
    const pass = rowsPass(pairs);
    const shapeError = targetWithPass(target, pairs, pass);
    if (shapeError !== undefined) return shapeError;
    if (isFormulaError(pass)) return pass;
    let min = Infinity;
    for (let row = 0; row < target.rows; row += 1) {
      for (let column = 0; column < target.columns; column += 1) {
        const index = row * target.columns + column;
        const value = target.values[row]![column]!;
        if (pass[index] && typeof value === "number") min = Math.min(min, value);
      }
    }
    return min === Infinity ? 0 : min;
  },
  SUMPRODUCT: (args) => {
    const matrices = args.values.filter((value) => Array.isArray(value)) as FormulaValue[][][];
    if (matrices.length === 0) return createFormulaError("#VALUE!", "SUMPRODUCT expects arrays");
    const first = matrices[0]!;
    let total = 0;
    for (let r = 0; r < first.length; r++) {
      const row = first[r]!;
      for (let c = 0; c < row.length; c++) {
        let product = typeof row[c] === "number" ? row[c] as number : 0;
        for (let m = 1; m < matrices.length; m++) {
          const other = matrices[m]?.[r]?.[c];
          product *= typeof other === "number" ? other : 0;
        }
        total += product;
      }
    }
    return total;
  },
  SUBTOTAL: (args) => {
    const mode = args.values[0];
    const data = args.values.slice(1).flatMap((value) => (flattenMatrix(value)));
    const numbers = numbersOf(data);
    switch (mode) {
      case 1: return numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : createFormulaError("#DIV/0!", "No data");
      case 2: return numbers.length;
      case 3: return data.filter((value) => value != null && value !== "").length;
      case 6: return numbers.reduce((a, b) => a * b, 1);
      case 9:
      default: return numbers.reduce((a, b) => a + b, 0);
    }
  },
  TEXTJOIN: (args) => {
    const delimiter = String(args.values[0] ?? "");
    const skipEmpty = args.values[1] === true || args.values[1] === 1;
    const parts: string[] = [];
    for (const value of args.values.slice(2)) {
      const items = flattenMatrix(value);
      for (const item of items) {
        if (skipEmpty && (item == null || item === "")) continue;
        parts.push(String(item ?? ""));
      }
    }
    return parts.join(delimiter);
  },
  IFS: (args) => {
    for (let i = 0; i + 1 < args.values.length; i += 2) {
      if (args.values[i] === true) return args.values[i + 1]!;
    }
    return createFormulaError("#N/A", "No IFS condition matched");
  },
  XLOOKUP: (args) => {
    const lookup = args.values[0] ?? null;
    const lookupArray = Array.isArray(args.values[1]) ? flattenMatrix(args.values[1] as FormulaValue[]) : [];
    const returnArray = Array.isArray(args.values[2]) ? flattenMatrix(args.values[2] as FormulaValue[]) : [];
    for (let i = 0; i < lookupArray.length; i++) {
      if (String(lookupArray[i]).toLowerCase() === String(lookup).toLowerCase()) {
        return returnArray[i] ?? null;
      }
    }
    return args.values[3] ?? createFormulaError("#N/A", "XLOOKUP no match");
  },
  ROMAN: () => createFormulaError("#VALUE!", "ROMAN not supported"),
};

export function evaluateAdvancedFunction(
  name: string,
  args: AdvancedFunctionArgs,
  context: AdvancedContext,
): FormulaValue | undefined {
  const fn = ADVANCED_FUNCTIONS[name.toUpperCase()];
  if (!fn) return undefined;
  try {
    return fn(args, context);
  } catch (error) {
    return createFormulaError("#VALUE!", error instanceof Error ? error.message : "Advanced function error");
  }
}
