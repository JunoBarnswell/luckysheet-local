import type { FormulaValue } from '../values';
import { mathFunctions } from './math';
import { statisticalFunctions } from './statistical';
import { logicalFunctions } from './logical';
import { textFunctions } from './text';
import { lookupFunctions } from './lookup';
import { datetimeFunctions } from './datetime';
import { informationFunctions } from './information';
import { extendedMatrixFunctions } from './extended-matrix';
import { dynamicArrayFunctions } from './dynamic-array';
import {
  createFormulaCapabilityError,
  DEFAULT_FORMULA_CAPABILITIES,
  getFormulaCapability,
  isFormulaCapabilityEnabled,
  type FormulaCapabilities,
} from '../capabilities';

export const BUILTIN_FUNCTIONS: Record<string, (args: FormulaValue[]) => FormulaValue> = {
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

const RANGE_FUNCTIONS = new Set(['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX', 'PRODUCT', 'VAR', 'VARP', 'STDEV', 'STDEVP', 'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'SUBTOTAL', 'SUMPRODUCT']);
const SORT_FUNCTIONS = new Set(['LOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'XMATCH', 'MEDIAN', 'PERCENTILE', 'SORT', 'FILTER', 'UNIQUE']);
const VOLATILE_FUNCTIONS = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT', 'RANDARRAY']);
const MATRIX_FUNCTIONS = new Set(['GROUPBY', 'PIVOTBY', 'FILTER', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE', 'RANDARRAY', 'HSTACK', 'VSTACK', 'TAKE', 'DROP']);

export const FUNCTION_DESCRIPTORS: ReadonlyMap<string, FunctionDescriptor> = new Map(
  Object.keys(BUILTIN_FUNCTIONS).map((name) => {
    const id = name.toUpperCase();
    const volatile = VOLATILE_FUNCTIONS.has(id);
    const cost = volatile ? 'volatile' : MATRIX_FUNCTIONS.has(id) ? 'range' : RANGE_FUNCTIONS.has(id) ? 'range' : SORT_FUNCTIONS.has(id) ? 'sort' : 'scalar';
    return [id, { id, cost, streaming: cost === 'range', volatile }];
  }),
);

export function getBuiltinFunction(
  name: string,
  capabilities: FormulaCapabilities = DEFAULT_FORMULA_CAPABILITIES,
): ((args: FormulaValue[]) => FormulaValue) | undefined {
  const normalized = name.trim().toUpperCase();
  const resolvedName = FORMULA_ALIASES[name.trim()] ?? normalized;
  const fn = BUILTIN_FUNCTIONS[resolvedName];
  if (!fn) return undefined;

  const capability = getFormulaCapability(resolvedName);
  if (capability && !isFormulaCapabilityEnabled(capabilities, capability)) {
    return () => createFormulaCapabilityError(resolvedName, capability, capabilities);
  }
  return fn;
}

export function getFunctionDescriptor(name: string): FunctionDescriptor | undefined {
  return FUNCTION_DESCRIPTORS.get(FORMULA_ALIASES[name.trim()] ?? name.trim().toUpperCase());
}
