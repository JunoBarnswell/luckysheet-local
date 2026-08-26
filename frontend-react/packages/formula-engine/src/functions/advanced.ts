import type { FormulaValue } from "../values";
import { createFormulaError, isFormulaError } from "../values";
import { matchesCriteria, parseCriteria, sameCriteriaShape, toCriteriaRange, type CriteriaExpression, type CriteriaRange } from "../criteria";
import { coerceExcelNumber, normalizeExcelPrecision } from "../numeric";
import type { ReferenceCell } from "../reference-cursor";
import type { RangeDependency } from "../range-index";

export interface AdvancedFunctionArgs {
  values: FormulaValue[];
  ranges: unknown[]; // EvaluationValue[](含 range 包装)
}

export interface AdvancedContext {
  toRanges(value: unknown): readonly RangeDependency[] | undefined;
  readCursor(range: RangeDependency): Iterable<ReferenceCell>;
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

function referenceCells(args: AdvancedFunctionArgs, context: AdvancedContext, startIndex: number): ReferenceCell[] | ReturnType<typeof createFormulaError> {
  const cells: ReferenceCell[] = [];
  const seen = new Set<string>();
  for (let index = startIndex; index < args.ranges.length; index += 1) {
    const ranges = context.toRanges(args.ranges[index]);
    if (!ranges || ranges.length === 0) return createFormulaError('#VALUE!', 'Aggregate functions require cell references');
    for (const range of ranges) {
      for (const cell of context.readCursor(range)) {
        const key = `${cell.address.sheetId}:${cell.address.row}:${cell.address.column}`;
        if (seen.has(key)) continue;
        seen.add(key);
        cells.push(cell);
      }
    }
  }
  if (cells.length === 0) return createFormulaError('#VALUE!', 'Aggregate reference is empty');
  return cells;
}

function visibleAggregateCells(cells: readonly ReferenceCell[], options: number, nestedKind: ReferenceCell['formulaKind'] | 'both'): ReferenceCell[] | ReturnType<typeof createFormulaError> {
  const ignoreManualHidden = options === 1 || options === 3 || options === 5 || options === 7;
  const ignoreErrors = options === 2 || options === 3 || options === 6 || options === 7;
  const ignoreNested = options <= 3;
  const visible: ReferenceCell[] = [];
  for (const cell of cells) {
    if (cell.visibility.filterHidden) continue;
    if (ignoreManualHidden && (cell.visibility.manualHidden || cell.visibility.outlineHidden)) continue;
    if (ignoreNested && (nestedKind === 'both'
      ? (cell.formulaKind === 'subtotal' || cell.formulaKind === 'aggregate')
      : cell.formulaKind === nestedKind)) continue;
    if (isFormulaError(cell.value)) {
      if (!ignoreErrors) return cell.value;
      continue;
    }
    visible.push(cell);
  }
  return visible;
}

function aggregateNumbers(cells: readonly ReferenceCell[]): number[] {
  return cells.flatMap((cell) => typeof cell.value === 'number' && Number.isFinite(cell.value) ? [cell.value] : []);
}

function aggregateResult(functionNumber: number, cells: readonly ReferenceCell[]): FormulaValue {
  const numbers = aggregateNumbers(cells);
  switch (functionNumber) {
    case 1: return numbers.length === 0 ? createFormulaError('#DIV/0!', 'No values for AVERAGE') : normalizeExcelPrecision(numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
    case 2: return numbers.length;
    case 3: return cells.filter((cell) => cell.value !== null && cell.value !== '').length;
    case 4: return numbers.length === 0 ? 0 : Math.max(...numbers);
    case 5: return numbers.length === 0 ? 0 : Math.min(...numbers);
    case 6: return numbers.reduce((product, value) => product * value, 1);
    case 7: return deviation(numbers, true);
    case 8: return deviation(numbers, false);
    case 9: return normalizeExcelPrecision(numbers.reduce((sum, value) => sum + value, 0));
    case 10: return variance(numbers, true);
    case 11: return variance(numbers, false);
    case 12: return median(numbers);
    case 13: return mode(numbers);
    default: return createFormulaError('#VALUE!', `Unsupported AGGREGATE function number: ${functionNumber}`);
  }
}

function deviation(values: readonly number[], sample: boolean): number | ReturnType<typeof createFormulaError> {
  const result = variance(values, sample);
  return isFormulaError(result) ? result : normalizeExcelPrecision(Math.sqrt(result));
}

function variance(values: readonly number[], sample: boolean): number | ReturnType<typeof createFormulaError> {
  const divisor = sample ? values.length - 1 : values.length;
  if (divisor <= 0) return createFormulaError('#DIV/0!', 'Insufficient values for variance');
  const average = values.reduce((sum, value) => sum + value, 0) / values.length;
  return normalizeExcelPrecision(values.reduce((sum, value) => sum + (value - average) ** 2, 0) / divisor);
}

function median(values: readonly number[]): FormulaValue {
  if (values.length === 0) return createFormulaError('#NUM!', 'No values for MEDIAN');
  const ordered = [...values].sort((left, right) => left - right);
  const middle = Math.floor(ordered.length / 2);
  return normalizeExcelPrecision(ordered.length % 2 === 0 ? (ordered[middle - 1]! + ordered[middle]!) / 2 : ordered[middle]!);
}

function mode(values: readonly number[]): FormulaValue {
  if (values.length === 0) return createFormulaError('#N/A', 'No values for MODE');
  const counts = new Map<number, number>();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);
  const highest = Math.max(...counts.values());
  if (highest < 2) return createFormulaError('#N/A', 'No repeated values for MODE');
  return [...counts.entries()].find(([, count]) => count === highest)![0];
}

function percentile(values: readonly number[], k: number, exclusive: boolean): FormulaValue {
  if (values.length === 0 || k < 0 || k > 1) return createFormulaError('#NUM!', 'Invalid percentile arguments');
  const ordered = [...values].sort((left, right) => left - right);
  const position = (exclusive ? (ordered.length + 1) * k : (ordered.length - 1) * k);
  if (position < 0 || position > ordered.length - 1) return createFormulaError('#NUM!', 'Percentile is outside the data range');
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return normalizeExcelPrecision(ordered[lower]! + (ordered[upper]! - ordered[lower]!) * (position - lower));
}

function aggregateFunctionNumber(args: AdvancedFunctionArgs, index: number): number | ReturnType<typeof createFormulaError> {
  const number = coerceExcelNumber(args.values[index] ?? null);
  if (isFormulaError(number) || !Number.isSafeInteger(number)) return createFormulaError('#VALUE!', 'Aggregate function number must be an integer');
  return number;
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
  SUBTOTAL: (args, context) => {
    const functionNumber = aggregateFunctionNumber(args, 0);
    if (isFormulaError(functionNumber) || functionNumber < 1 || functionNumber > 111 || (functionNumber > 11 && functionNumber < 101)) {
      return createFormulaError('#VALUE!', 'SUBTOTAL function number must be 1-11 or 101-111');
    }
    const cells = referenceCells(args, context, 1);
    if (isFormulaError(cells)) return cells;
    const visible = visibleAggregateCells(cells, functionNumber >= 101 ? 1 : 0, 'both');
    if (isFormulaError(visible)) return visible;
    const mode = functionNumber >= 101 ? functionNumber - 100 : functionNumber;
    return aggregateResult(mode, visible);
  },
  AGGREGATE: (args, context) => {
    const functionNumber = aggregateFunctionNumber(args, 0);
    const options = aggregateFunctionNumber(args, 1);
    if (isFormulaError(functionNumber) || functionNumber < 1 || functionNumber > 19) return createFormulaError('#VALUE!', 'AGGREGATE function number must be 1-19');
    if (isFormulaError(options) || options < 0 || options > 7) return createFormulaError('#VALUE!', 'AGGREGATE options must be 0-7');
    const cells = referenceCells(args, context, 3);
    if (isFormulaError(cells)) return cells;
    const visible = visibleAggregateCells(cells, options, 'both');
    if (isFormulaError(visible)) return visible;
    if (functionNumber <= 13) return aggregateResult(functionNumber, visible);
    const k = coerceExcelNumber(args.values[2] ?? null);
    if (isFormulaError(k) || k < 1) return createFormulaError('#NUM!', 'AGGREGATE k must be positive');
    const numbers = aggregateNumbers(visible);
    const ordered = [...numbers].sort((left, right) => left - right);
    if (functionNumber === 14 || functionNumber === 15) {
      const index = functionNumber === 14 ? ordered.length - Math.trunc(k) : Math.trunc(k) - 1;
      return index < 0 || index >= ordered.length ? createFormulaError('#NUM!', 'AGGREGATE k is outside the data range') : ordered[index]!;
    }
    if (functionNumber === 16 || functionNumber === 17) return percentile(numbers, k, false);
    if (functionNumber === 18 || functionNumber === 19) return percentile(numbers, k, true);
    return createFormulaError('#VALUE!', `Unsupported AGGREGATE function number: ${functionNumber}`);
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
