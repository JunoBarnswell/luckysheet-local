import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import { flattenNumericArgs } from './math';

export const statisticalFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  AVERAGE: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    if (nums.length === 0) return createFormulaError('#DIV/0!', 'No numbers to average');
    return normalizeExcelPrecision(nums.reduce((acc, n) => acc + n, 0) / nums.length);
  },

  COUNT: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    return nums.length;
  },

  COUNTA: (args) => {
    let count = 0;
    for (const arg of args) {
      if (Array.isArray(arg)) {
        for (const row of arg) {
          if (Array.isArray(row)) {
            for (const cell of row) {
              if (cell !== null && cell !== undefined && cell !== '') count += 1;
            }
          } else if (row !== null && row !== undefined && row !== '') {
            count += 1;
          }
        }
      } else if (arg !== null && arg !== undefined && arg !== '') {
        count += 1;
      }
    }
    return count;
  },

  COUNTBLANK: (args) => {
    let count = 0;
    for (const arg of args) {
      if (Array.isArray(arg)) {
        for (const row of arg) {
          if (Array.isArray(row)) {
            for (const cell of row) {
              if (cell === null || cell === undefined || cell === '') count += 1;
            }
          } else if (row === null || row === undefined || row === '') {
            count += 1;
          }
        }
      } else if (arg === null || arg === undefined || arg === '') {
        count += 1;
      }
    }
    return count;
  },

  MIN: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    if (nums.length === 0) return 0;
    return normalizeExcelPrecision(Math.min(...nums));
  },

  MAX: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    if (nums.length === 0) return 0;
    return normalizeExcelPrecision(Math.max(...nums));
  },

  MEDIAN: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    if (nums.length === 0) return createFormulaError('#NUM!', 'No numbers for MEDIAN');
    nums.sort((a, b) => a - b);
    const mid = Math.floor(nums.length / 2);
    return normalizeExcelPrecision(nums.length % 2 !== 0 ? nums[mid]! : (nums[mid - 1]! + nums[mid]!) / 2);
  },

  LARGE: (args) => {
    const nums = flattenNumericArgs([args[0]!]);
    if (isFormulaError(nums)) return nums;
    const k = coerceExcelNumber(args[1]);
    if (isFormulaError(k) || k <= 0 || k > nums.length) return isFormulaError(k) ? k : createFormulaError('#NUM!', 'Invalid k in LARGE');
    nums.sort((a, b) => b - a);
    return normalizeExcelPrecision(nums[k - 1]!);
  },

  SMALL: (args) => {
    const nums = flattenNumericArgs([args[0]!]);
    if (isFormulaError(nums)) return nums;
    const k = coerceExcelNumber(args[1]);
    if (isFormulaError(k) || k <= 0 || k > nums.length) return isFormulaError(k) ? k : createFormulaError('#NUM!', 'Invalid k in SMALL');
    nums.sort((a, b) => a - b);
    return normalizeExcelPrecision(nums[k - 1]!);
  },

  COUNTIF: (args) => {
    const range = args[0];
    const criteria = args[1];
    if (!Array.isArray(range)) return 0;

    let count = 0;
    const criteriaStr = String(criteria ?? '');
    const isComparison = /^[><]=?/.test(criteriaStr);

    for (const row of range) {
      if (Array.isArray(row)) {
        for (const cell of row) {
          if (matchesCriteria(cell, criteriaStr, isComparison)) count += 1;
        }
      }
    }
    return count;
  },

  SUMIF: (args) => {
    const range = args[0];
    const criteria = args[1];
    const sumRange = args[2] !== undefined ? args[2] : range;
    if (!Array.isArray(range) || !Array.isArray(sumRange)) return 0;

    let sum = 0;
    const criteriaStr = String(criteria ?? '');
    const isComparison = /^[><]=?/.test(criteriaStr);

    for (let r = 0; r < range.length; r++) {
      const row = range[r];
      const sumRow = sumRange[r];
      if (Array.isArray(row)) {
        for (let c = 0; c < row.length; c++) {
          if (matchesCriteria(row[c], criteriaStr, isComparison)) {
            const sumVal = Array.isArray(sumRow) ? sumRow[c] : sumRow;
            if (typeof sumVal === 'number') sum += sumVal;
            else if (typeof sumVal === 'string') {
              const numeric = coerceExcelNumber(sumVal);
              if (!isFormulaError(numeric)) sum += numeric;
            }
          }
        }
      }
    }
    return normalizeExcelPrecision(sum);
  },
};

function matchesCriteria(cellValue: unknown, criteria: string, isComparison: boolean): boolean {
  if (isComparison) {
    const num = coerceExcelNumber(cellValue as FormulaValue);
    const threshold = (text: string) => coerceExcelNumber(text);
    if (isFormulaError(num)) return false;
    if (criteria.startsWith('>=')) { const value = threshold(criteria.slice(2)); return !isFormulaError(value) && num >= value; }
    if (criteria.startsWith('<=')) { const value = threshold(criteria.slice(2)); return !isFormulaError(value) && num <= value; }
    if (criteria.startsWith('>')) { const value = threshold(criteria.slice(1)); return !isFormulaError(value) && num > value; }
    if (criteria.startsWith('<')) { const value = threshold(criteria.slice(1)); return !isFormulaError(value) && num < value; }
  }
  if (typeof cellValue === 'number') {
    const numericCriteria = coerceExcelNumber(criteria);
    if (!isFormulaError(numericCriteria)) return cellValue === numericCriteria;
  }
  return String(cellValue ?? '').toLowerCase() === criteria.toLowerCase();
}
