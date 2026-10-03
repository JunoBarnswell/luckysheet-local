import type { FormulaValue } from '../values';
import type { FormulaEvaluationContext } from '../evaluator';
import { mathFunctions } from './math';
import { statisticalFunctions } from './statistical';
import { logicalFunctions } from './logical';
import { textFunctions } from './text';
import { lookupFunctions } from './lookup';
import { datetimeFunctions } from './datetime';
import { informationFunctions } from './information';
import { financialFunctions } from './financial';
import { extendedMatrixFunctions } from './extended-matrix';
import { dynamicArrayFunctions } from './dynamic-array';
import { ADVANCED_FUNCTIONS } from './advanced';
import { STREAMING_AGGREGATES } from './aggregate';
import { RANGE_CONSUMER_FUNCTIONS } from './range-functions';
export * from './contracts';

export type BuiltinFunction = (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue;

export const BUILTIN_FUNCTIONS: Record<string, BuiltinFunction> = {
  ...mathFunctions,
  ...statisticalFunctions,
  ...logicalFunctions,
  ...textFunctions,
  ...lookupFunctions,
  ...datetimeFunctions,
  ...informationFunctions,
  ...financialFunctions,
  ...extendedMatrixFunctions,
  ...dynamicArrayFunctions,
};

export interface FunctionDescriptor {
  id: string;
  cost: 'scalar' | 'range' | 'sort' | 'volatile' | 'external';
  streaming: boolean;
  volatile: boolean;
}

export type FunctionLibraryCategory = 'logical' | 'text' | 'date-time' | 'lookup-reference' | 'math-trig' | 'more-functions';

export interface FunctionLibraryEntry extends FunctionDescriptor {
  category: FunctionLibraryCategory;
}

const RANGE_FUNCTIONS = new Set(['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX', 'PRODUCT', 'VAR', 'VARP', 'STDEV', 'STDEVP', 'VAR.S', 'VAR.P', 'STDEV.S', 'STDEV.P', 'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'SUBTOTAL', 'AGGREGATE', 'SUMPRODUCT']);
const SORT_FUNCTIONS = new Set(['LOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'XMATCH', 'MEDIAN', 'PERCENTILE', 'SORT', 'FILTER', 'UNIQUE']);
const VOLATILE_FUNCTIONS = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT', 'RANDARRAY']);
const MATRIX_FUNCTIONS = new Set(['GROUPBY', 'PIVOTBY', 'FILTER', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE', 'RANDARRAY', 'HSTACK', 'VSTACK', 'TAKE', 'DROP']);
const NATIVE_FUNCTIONS = new Set(['SJS.TABLE', 'LET', 'LAMBDA', 'ROW', 'COLUMN', 'ADDRESS', 'OFFSET', 'INDIRECT']);

export const FUNCTION_DESCRIPTORS: ReadonlyMap<string, FunctionDescriptor> = new Map(
  [...new Set([...Object.keys(BUILTIN_FUNCTIONS), ...Object.keys(ADVANCED_FUNCTIONS), ...NATIVE_FUNCTIONS])].map((name) => {
    const id = name.toUpperCase();
    const volatile = VOLATILE_FUNCTIONS.has(id);
    const cost = volatile ? 'volatile' : MATRIX_FUNCTIONS.has(id) || NATIVE_FUNCTIONS.has(id) ? 'range' : RANGE_FUNCTIONS.has(id) ? 'range' : SORT_FUNCTIONS.has(id) ? 'sort' : 'scalar';
    return [id, { id, cost, streaming: STREAMING_AGGREGATES.has(id) || RANGE_CONSUMER_FUNCTIONS.has(id) || id === 'SUBTOTAL' || id === 'AGGREGATE', volatile }];
  }),
);

const categoryForFunction = (name: string): FunctionLibraryCategory => {
  if (Object.prototype.hasOwnProperty.call(logicalFunctions, name)) return 'logical';
  if (Object.prototype.hasOwnProperty.call(textFunctions, name)) return 'text';
  if (Object.prototype.hasOwnProperty.call(datetimeFunctions, name)) return 'date-time';
  if (Object.prototype.hasOwnProperty.call(lookupFunctions, name)) return 'lookup-reference';
  if (['ROW', 'COLUMN', 'ADDRESS', 'OFFSET', 'INDIRECT', 'XMATCH'].includes(name)) return 'lookup-reference';
  if (name === 'LET' || name === 'LAMBDA') return 'logical';
  if (Object.prototype.hasOwnProperty.call(mathFunctions, name)) return 'math-trig';
  return 'more-functions';
};

/** The only UI-facing function list; it is derived from the executable registry. */
export const FUNCTION_LIBRARY: readonly FunctionLibraryEntry[] = [...FUNCTION_DESCRIPTORS.values()]
  .map((descriptor) => ({ ...descriptor, category: categoryForFunction(descriptor.id) }))
  .sort((left, right) => left.id.localeCompare(right.id));

export function listFunctionLibrary(category?: FunctionLibraryCategory): readonly FunctionLibraryEntry[] {
  return category ? FUNCTION_LIBRARY.filter((entry) => entry.category === category) : FUNCTION_LIBRARY;
}

export function getBuiltinFunction(name: string): BuiltinFunction | undefined {
  const id = name.trim().toUpperCase();
  return Object.prototype.hasOwnProperty.call(BUILTIN_FUNCTIONS, id) ? BUILTIN_FUNCTIONS[id] : undefined;
}

export function getFunctionDescriptor(name: string): FunctionDescriptor | undefined {
  return FUNCTION_DESCRIPTORS.get(name.trim().toUpperCase());
}
