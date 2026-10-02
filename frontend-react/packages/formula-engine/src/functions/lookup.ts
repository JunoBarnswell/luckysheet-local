import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { evaluateRangeFunction } from './range-functions';
import { coerceExcelNumber } from '../numeric';
import type { FormulaEvaluationContext } from '../evaluator';

function to2DArray(val: FormulaValue | undefined): FormulaValue[][] {
  if (val === undefined || val === null) return [[]];
  if (Array.isArray(val)) {
    if (val.length === 0) return [[]];
    if (Array.isArray(val[0])) return val as FormulaValue[][];
    return [val as FormulaValue[]];
  }
  return [[val]];
}

export const lookupFunctions: Record<string, (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue> = {
  ...Object.fromEntries(['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'ROWS', 'COLUMNS'].map((name) => [name, (args: FormulaValue[], context?: FormulaEvaluationContext) => evaluateRangeFunction(name, args, context) as FormulaValue])),

  CHOOSE: (args) => {
    const index = coerceExcelNumber(args[0]);
    if (isFormulaError(index) || index < 1 || index >= args.length) {
      return createFormulaError('#VALUE!', 'Index out of bounds in CHOOSE');
    }
    return args[index] ?? null;
  },

  TRANSPOSE: (args) => {
    const table = to2DArray(args[0]);
    if (table.length === 0 || !table[0]) return [[]];
    const rowCount = table.length;
    const colCount = table[0].length;
    const result: FormulaValue[][] = [];

    for (let c = 0; c < colCount; c++) {
      const newRow: FormulaValue[] = [];
      for (let r = 0; r < rowCount; r++) {
        newRow.push(table[r]?.[c] ?? null);
      }
      result.push(newRow);
    }
    return result;
  },
};
