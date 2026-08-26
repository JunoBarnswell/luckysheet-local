import type { FormulaValue } from '../values';
import type { FormulaEvaluationContext } from '../evaluator';
import { mathFunctions } from './math';
import { statisticalFunctions } from './statistical';
import { logicalFunctions } from './logical';
import { textFunctions } from './text';
import { lookupFunctions } from './lookup';
import { datetimeFunctions } from './datetime';
import { informationFunctions } from './information';
import { extendedMatrixFunctions } from './extended-matrix';
import { dynamicArrayFunctions } from './dynamic-array';
import { ADVANCED_FUNCTIONS } from './advanced';

export type BuiltinFunction = (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue;

export const BUILTIN_FUNCTIONS: Record<string, BuiltinFunction> = {
  ...mathFunctions,
  ...statisticalFunctions,
  ...logicalFunctions,
  ...textFunctions,
  ...lookupFunctions,
  ...datetimeFunctions,
  ...informationFunctions,
  ...extendedMatrixFunctions,
  ...dynamicArrayFunctions,
};

export interface FunctionDescriptor {
  id: string;
  cost: 'scalar' | 'range' | 'sort' | 'volatile' | 'external';
  streaming: boolean;
  volatile: boolean;
}

const RANGE_FUNCTIONS = new Set(['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX', 'PRODUCT', 'VAR', 'VARP', 'STDEV', 'STDEVP', 'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'SUBTOTAL', 'AGGREGATE', 'SUMPRODUCT']);
const SORT_FUNCTIONS = new Set(['LOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'XMATCH', 'MEDIAN', 'PERCENTILE', 'SORT', 'FILTER', 'UNIQUE']);
const VOLATILE_FUNCTIONS = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT', 'RANDARRAY']);
const MATRIX_FUNCTIONS = new Set(['GROUPBY', 'PIVOTBY', 'FILTER', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE', 'RANDARRAY', 'HSTACK', 'VSTACK', 'TAKE', 'DROP']);

export const FUNCTION_DESCRIPTORS: ReadonlyMap<string, FunctionDescriptor> = new Map(
  [...new Set([...Object.keys(BUILTIN_FUNCTIONS), ...Object.keys(ADVANCED_FUNCTIONS)])].map((name) => {
    const id = name.toUpperCase();
    const volatile = VOLATILE_FUNCTIONS.has(id);
    const cost = volatile ? 'volatile' : MATRIX_FUNCTIONS.has(id) ? 'range' : RANGE_FUNCTIONS.has(id) ? 'range' : SORT_FUNCTIONS.has(id) ? 'sort' : 'scalar';
    return [id, { id, cost, streaming: cost === 'range', volatile }];
  }),
);

export function getBuiltinFunction(name: string): BuiltinFunction | undefined {
  return BUILTIN_FUNCTIONS[name.trim().toUpperCase()];
}

export function getFunctionDescriptor(name: string): FunctionDescriptor | undefined {
  return FUNCTION_DESCRIPTORS.get(name.trim().toUpperCase());
}
