import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { findLookupIndex, type LookupMatchMode } from './lookup-engine';
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

function to1DArray(val: FormulaValue | undefined): FormulaValue[] {
  if (val === undefined || val === null) return [];
  if (Array.isArray(val)) {
    const list: FormulaValue[] = [];
    for (const row of val) {
      if (Array.isArray(row)) {
        for (const cell of row) list.push(cell);
      } else {
        list.push(row);
      }
    }
    return list;
  }
  return [val];
}

export const lookupFunctions: Record<string, (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue> = {
  VLOOKUP: (args, context) => {
    const lookupValue = args[0];
    const table = to2DArray(args[1]);
    const colIndex = coerceExcelNumber(args[2]);
    const exactMatch = args[3] === undefined ? false : !Boolean(args[3]) || String(args[3]).toLowerCase() === 'false';

    if (isFormulaError(colIndex) || colIndex < 1 || colIndex > (table[0]?.length ?? 0)) {
      return createFormulaError('#REF!', 'Column index out of bounds in VLOOKUP');
    }

    const index = findLookupIndex(lookupValue, table.map((row) => row[0] ?? null), exactMatch ? 0 : -1, 1, context?.collationContext);
    if (index >= 0) return table[index]?.[colIndex - 1] ?? null;
    return createFormulaError('#N/A', 'Value not found in VLOOKUP');
  },

  HLOOKUP: (args, context) => {
    const lookupValue = args[0];
    const table = to2DArray(args[1]);
    const rowIndex = coerceExcelNumber(args[2]);
    const exactMatch = args[3] === undefined ? false : !Boolean(args[3]) || String(args[3]).toLowerCase() === 'false';

    if (isFormulaError(rowIndex) || rowIndex < 1 || rowIndex > table.length) {
      return createFormulaError('#REF!', 'Row index out of bounds in HLOOKUP');
    }

    const firstRow = table[0] ?? [];
    const column = findLookupIndex(lookupValue, firstRow, exactMatch ? 0 : -1, 1, context?.collationContext);
    if (column >= 0) return table[rowIndex - 1]?.[column] ?? null;
    return createFormulaError('#N/A', 'Value not found in HLOOKUP');
  },

  INDEX: (args) => {
    const table = to2DArray(args[0]);
    const rowNum = args[1] !== undefined ? coerceExcelNumber(args[1]) : 1;
    const colNum = args[2] !== undefined ? coerceExcelNumber(args[2]) : 1;

    if (isFormulaError(rowNum) || isFormulaError(colNum) || rowNum < 0 || colNum < 0) {
      return createFormulaError('#VALUE!', 'Invalid index in INDEX');
    }
    if (rowNum === 0 && colNum === 0) return table;
    if (rowNum === 0) {
      // Return whole column
      return table.map((r) => [r[colNum - 1] ?? null]);
    }
    if (colNum === 0) {
      // Return whole row
      return [table[rowNum - 1] ?? []];
    }
    const row = table[rowNum - 1];
    if (!row) return createFormulaError('#REF!', 'Row out of bounds in INDEX');
    const cell = row[colNum - 1];
    return cell !== undefined ? cell : createFormulaError('#REF!', 'Column out of bounds in INDEX');
  },

  MATCH: (args, context) => {
    const lookupValue = args[0];
    const array = to1DArray(args[1]);
    const matchType = args[2] !== undefined ? coerceExcelNumber(args[2]) : 1;

    if (isFormulaError(matchType) || ![1, 0, -1].includes(matchType)) return createFormulaError('#N/A', 'Invalid match type in MATCH');
    const index = findLookupIndex(lookupValue, array, matchType as LookupMatchMode, 1, context?.collationContext);
    if (index >= 0) return index + 1;
    return createFormulaError('#N/A', 'Value not found in MATCH');
  },

  XLOOKUP: (args, context) => {
    const lookupValue = args[0];
    const lookupMatrix = to2DArray(args[1]);
    const returnMatrix = to2DArray(args[2]);
    const horizontal = lookupMatrix.length === 1;
    const lookupArray = horizontal ? lookupMatrix[0] ?? [] : lookupMatrix.map((row) => row[0] ?? null);
    const matchMode = coerceExcelNumber(args[4] ?? 0);
    const searchMode = coerceExcelNumber(args[5] ?? 1);
    if (isFormulaError(matchMode) || isFormulaError(searchMode) || ![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return createFormulaError('#VALUE!', 'Invalid XLOOKUP mode');
    const ifNotFound = args[3] !== undefined ? args[3] : createFormulaError('#N/A', 'Value not found in XLOOKUP');

    const index = findLookupIndex(lookupValue, lookupArray, matchMode as LookupMatchMode, searchMode, context?.collationContext);
    if (index >= 0) {
      if (horizontal) {
        if (returnMatrix.length === 1) return returnMatrix[0]?.[index] ?? null;
        return returnMatrix.map((row) => [row[index] ?? null]);
      }
      const row = returnMatrix[index] ?? [];
      return row.length <= 1 ? row[0] ?? null : [row.map((value) => value ?? null)];
    }
    return ifNotFound;
  },

  CHOOSE: (args) => {
    const index = coerceExcelNumber(args[0]);
    if (isFormulaError(index) || index < 1 || index >= args.length) {
      return createFormulaError('#VALUE!', 'Index out of bounds in CHOOSE');
    }
    return args[index] ?? null;
  },

  COLUMNS: (args) => {
    const table = to2DArray(args[0]);
    return table[0]?.length ?? 0;
  },

  ROWS: (args) => {
    const table = to2DArray(args[0]);
    return table.length;
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
