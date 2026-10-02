import type { FormulaEvaluationValue } from '../evaluator';
import type { FormulaValue } from '../values';
import { createFormulaError, isFormulaError } from '../values';
import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import type { ReferenceCell } from '../reference-cursor';
import type { RangeDependency } from '../range-index';
import { evaluateAggregate } from './aggregate';

export interface AdvancedFunctionArgs {
  values: FormulaValue[];
  ranges: FormulaEvaluationValue[];
}
export interface AdvancedContext {
  toRanges(value: FormulaEvaluationValue): readonly RangeDependency[] | undefined;
  readCursor(range: RangeDependency): Iterable<ReferenceCell>;
}
type AdvancedFn = (args: AdvancedFunctionArgs, context: AdvancedContext) => FormulaValue;
const AGGREGATIONS = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P'];

function* referenceValues(args: AdvancedFunctionArgs, context: AdvancedContext, start: number, end: number, options: number, subtotal: boolean): Iterable<FormulaValue> {
  for (let index = start; index < end; index++) {
    const ranges = context.toRanges(args.ranges[index]!);
    if (!ranges?.length) { yield createFormulaError('#VALUE!', 'Aggregate functions require cell references'); return; }
    for (const range of ranges) for (const cell of context.readCursor(range)) {
      if (cell.visibility.filterHidden) continue;
      if ([1, 3, 5, 7].includes(options) && (cell.visibility.manualHidden || cell.visibility.outlineHidden)) continue;
      if ((subtotal || options <= 3) && (cell.formulaKind === 'subtotal' || cell.formulaKind === 'aggregate')) continue;
      if ([2, 3, 6, 7].includes(options) && isFormulaError(cell.value)) continue;
      yield cell.value;
    }
  }
}
function parameter(args: AdvancedFunctionArgs, index: number): number | ReturnType<typeof createFormulaError> {
  const value = coerceExcelNumber(args.values[index]);
  return isFormulaError(value) || !Number.isInteger(value) ? createFormulaError('#VALUE!', 'Aggregate parameter must be an integer') : value;
}
function percentile(numbers: readonly number[], k: number, exclusive: boolean): FormulaValue {
  if (!numbers.length || k < 0 || k > 1) return createFormulaError('#NUM!', 'Invalid percentile');
  const position = exclusive ? (numbers.length + 1) * k - 1 : (numbers.length - 1) * k;
  if (position < 0 || position > numbers.length - 1) return createFormulaError('#NUM!', 'Percentile is outside the data range');
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return normalizeExcelPrecision(numbers[lower]! + (numbers[upper]! - numbers[lower]!) * (position - lower));
}
export const ADVANCED_FUNCTIONS: Record<string, AdvancedFn> = {
  SUBTOTAL: (args, context) => {
    const mode = parameter(args, 0);
    if (args.ranges.length < 2 || isFormulaError(mode) || !((mode >= 1 && mode <= 11) || (mode >= 101 && mode <= 111))) return createFormulaError('#VALUE!', 'SUBTOTAL requires a function number and references');
    return evaluateAggregate(AGGREGATIONS[(mode >= 101 ? mode - 100 : mode) - 1]!, [{ reference: true, values: referenceValues(args, context, 1, args.ranges.length, mode >= 101 ? 1 : 0, true) }]);
  },
  AGGREGATE: (args, context) => {
    const mode = parameter(args, 0);
    const options = parameter(args, 1);
    if (isFormulaError(mode) || mode < 1 || mode > 19 || isFormulaError(options) || options < 0 || options > 7 || args.ranges.length < 3) return createFormulaError('#VALUE!', 'Invalid AGGREGATE arguments');
    if (mode <= 11) return evaluateAggregate(AGGREGATIONS[mode - 1]!, [{ reference: true, values: referenceValues(args, context, 2, args.ranges.length, options, false) }]);
    if (mode >= 14 && args.ranges.length !== 4) return createFormulaError('#VALUE!', 'AGGREGATE array form requires array and k');
    const values = referenceValues(args, context, 2, mode >= 14 ? 3 : args.ranges.length, options, false);
    const numbers: number[] = [];
    for (const value of values) { if (isFormulaError(value)) return value; if (typeof value === 'number') numbers.push(value); }
    if (!numbers.length) return createFormulaError('#NUM!', 'AGGREGATE has no numeric values');
    if (mode === 13) {
      const counts = new Map<number, number>();
      let best = numbers[0]!;
      let highest = 1;
      for (const value of numbers) { const count = (counts.get(value) ?? 0) + 1; counts.set(value, count); if (count > highest) { best = value; highest = count; } }
      return highest < 2 ? createFormulaError('#N/A', 'No repeated values') : best;
    }
    numbers.sort((a, b) => a - b);
    if (mode === 12) { const middle = Math.floor(numbers.length / 2); return normalizeExcelPrecision(numbers.length % 2 ? numbers[middle]! : (numbers[middle - 1]! + numbers[middle]!) / 2); }
    const k = coerceExcelNumber(args.values[3]);
    if (isFormulaError(k)) return k;
    if (mode === 14 || mode === 15) { const ordinal = Math.trunc(k); return ordinal < 1 || ordinal > numbers.length ? createFormulaError('#NUM!', 'AGGREGATE k is out of bounds') : numbers[mode === 14 ? numbers.length - ordinal : ordinal - 1]!; }
    if (mode === 17 || mode === 19) { if (!Number.isInteger(k)) return createFormulaError('#NUM!', 'Quartile must be an integer'); return percentile(numbers, k / 4, mode === 19); }
    return percentile(numbers, k, mode === 18);
  },
};
export function evaluateAdvancedFunction(name: string, args: AdvancedFunctionArgs, context: AdvancedContext): FormulaValue | undefined {
  const fn = ADVANCED_FUNCTIONS[name.toUpperCase()];
  if (!fn) return undefined;
  return fn(args, context);
}
