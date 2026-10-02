import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import { createFormulaError, isFormulaError, type FormulaValue } from '../values';

export const STREAMING_AGGREGATES = new Set([
  'SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'PRODUCT',
  'VAR', 'VAR.S', 'VARP', 'VAR.P', 'STDEV', 'STDEV.S', 'STDEVP', 'STDEV.P',
]);

/** References exclude text/booleans; scalar arguments use Excel coercion. */
export interface AggregateArgument {
  readonly values: Iterable<FormulaValue>;
  readonly reference: boolean;
}

export function evaluateAggregate(name: string, args: Iterable<AggregateArgument>): FormulaValue {
  let count = 0;
  let occupied = 0;
  let sum = 0;
  let product = 1;
  let minimum = Infinity;
  let maximum = -Infinity;
  let mean = 0;
  let m2 = 0;
  for (const argument of args) {
    for (const value of argument.values) {
      if (Array.isArray(value)) return createFormulaError('#VALUE!', 'Aggregate input must contain scalar values');
      if (value !== null) occupied += 1;
      if (name === 'COUNTA') continue;
      if (isFormulaError(value)) {
        if (name === 'COUNT') continue;
        return value;
      }
      let number: number;
      if (typeof value === 'number') number = value;
      else if (argument.reference || value === null) continue;
      else {
        const coerced = coerceExcelNumber(value);
        if (isFormulaError(coerced)) {
          if (name === 'COUNT') continue;
          return coerced;
        }
        number = coerced;
      }
      if (!Number.isFinite(number)) return createFormulaError('#NUM!', 'Aggregate input is not finite');
      count += 1;
      sum += number;
      product *= number;
      minimum = Math.min(minimum, number);
      maximum = Math.max(maximum, number);
      const delta = number - mean;
      mean += delta / count;
      m2 += delta * (number - mean);
    }
  }
  if (name === 'COUNT') return count;
  if (name === 'COUNTA') return occupied;
  let result: number;
  switch (name) {
    case 'SUM': result = sum; break;
    case 'AVERAGE':
      if (!count) return createFormulaError('#DIV/0!', 'No numbers to average');
      result = sum / count;
      break;
    case 'MIN': result = count ? minimum : 0; break;
    case 'MAX': result = count ? maximum : 0; break;
    case 'PRODUCT': result = count ? product : 0; break;
    default: {
      const population = name === 'VARP' || name === 'VAR.P' || name === 'STDEVP' || name === 'STDEV.P';
      const denominator = count - (population ? 0 : 1);
      if (denominator <= 0) return createFormulaError('#DIV/0!', 'Insufficient numbers for variance');
      const variance = Math.max(0, m2 / denominator);
      result = name.startsWith('STDEV') ? Math.sqrt(variance) : variance;
    }
  }
  return Number.isFinite(result) ? normalizeExcelPrecision(result) : createFormulaError('#NUM!', 'Aggregate result is not finite');
}

export function* flattenAggregateValue(value: FormulaValue): Iterable<FormulaValue> {
  if (Array.isArray(value)) {
    for (const row of value) for (const cell of row) yield* flattenAggregateValue(cell);
  } else yield value;
}
