import type { FormulaEvaluationContext } from '../evaluator';
import type { FormulaValue } from '../values';
import { evaluateAggregate, flattenAggregateValue } from './aggregate';
import { evaluateRangeFunction } from './range-functions';
export const statisticalFunctions: Record<string, (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue> = {
  ...Object.fromEntries(['AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'VAR', 'VAR.S', 'VARP', 'VAR.P', 'STDEV', 'STDEV.S', 'STDEVP', 'STDEV.P'].map((name) => [name, (args: FormulaValue[]) => evaluateAggregate(name, args.map((value) => ({ values: flattenAggregateValue(value), reference: Array.isArray(value) })))])),
  ...Object.fromEntries(['COUNTBLANK', 'MEDIAN', 'LARGE', 'SMALL', 'COUNTIF', 'SUMIF', 'AVERAGEIF', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS'].map((name) => [name, (args: FormulaValue[], context?: FormulaEvaluationContext) => evaluateRangeFunction(name, args, context) as FormulaValue])),
};
