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

export function getBuiltinFunction(name: string): ((args: FormulaValue[]) => FormulaValue) | undefined {
  return BUILTIN_FUNCTIONS[name.toUpperCase()];
}
