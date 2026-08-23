import type { FormulaValue } from '../values';
import { mathFunctions } from './math';
import { statisticalFunctions } from './statistical';
import { logicalFunctions } from './logical';
import { textFunctions } from './text';
import { lookupFunctions } from './lookup';
import { datetimeFunctions } from './datetime';
import { informationFunctions } from './information';

export const BUILTIN_FUNCTIONS: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  ...mathFunctions,
  ...statisticalFunctions,
  ...logicalFunctions,
  ...textFunctions,
  ...lookupFunctions,
  ...datetimeFunctions,
  ...informationFunctions,
};

export const FORMULA_ALIASES: Readonly<Record<string, string>> = {
  '求和': 'SUM',
  '计数': 'COUNT',
  '平均值': 'AVERAGE',
  '最大值': 'MAX',
  '最小值': 'MIN',
  '如果': 'IF',
  '查找': 'LOOKUP',
  '四舍五入': 'ROUND',
};

export interface FunctionDescriptor {
  id: string;
  cost: 'scalar' | 'range' | 'sort' | 'volatile' | 'external';
  streaming: boolean;
  volatile: boolean;
}

const RANGE_FUNCTIONS = new Set(['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX', 'PRODUCT', 'VAR', 'VARP', 'STDEV', 'STDEVP']);
const SORT_FUNCTIONS = new Set(['LOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'MEDIAN', 'PERCENTILE']);
const VOLATILE_FUNCTIONS = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT']);

export const FUNCTION_DESCRIPTORS: ReadonlyMap<string, FunctionDescriptor> = new Map(
  Object.keys(BUILTIN_FUNCTIONS).map((name) => {
    const id = name.toUpperCase();
    const volatile = VOLATILE_FUNCTIONS.has(id);
    const cost = volatile ? 'volatile' : RANGE_FUNCTIONS.has(id) ? 'range' : SORT_FUNCTIONS.has(id) ? 'sort' : 'scalar';
    return [id, { id, cost, streaming: cost === 'range', volatile }];
  }),
);

export function getBuiltinFunction(name: string): ((args: FormulaValue[]) => FormulaValue) | undefined {
  const normalized = name.trim().toUpperCase();
  return BUILTIN_FUNCTIONS[FORMULA_ALIASES[name.trim()] ?? normalized];
}

export function getFunctionDescriptor(name: string): FunctionDescriptor | undefined {
  return FUNCTION_DESCRIPTORS.get(FORMULA_ALIASES[name.trim()] ?? name.trim().toUpperCase());
}
