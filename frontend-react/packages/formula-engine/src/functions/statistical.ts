import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import { matchesCriteria, parseCriteria, projectCriteriaRange, toCriteriaRange } from '../criteria';
import { flattenNumericArgs } from './math';
import { evaluateAggregate, flattenAggregateValue } from './aggregate';

export const statisticalFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  ...Object.fromEntries(['AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'VAR', 'VAR.S', 'VARP', 'VAR.P', 'STDEV', 'STDEV.S', 'STDEVP', 'STDEV.P'].map((name) => [name, (args: FormulaValue[]) => evaluateAggregate(name, args.map((value) => ({ values: flattenAggregateValue(value), reference: Array.isArray(value) })))])),

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
    const range = toCriteriaRange(args[0] ?? null);
    const criteria = parseCriteria(args[1] ?? null);
    if (range.columns < 0) return createFormulaError('#VALUE!', 'COUNTIF range must be rectangular');
    let count = 0;
    for (const row of range.values) {
      for (const cell of row) {
        if (matchesCriteria(cell, criteria)) count += 1;
      }
    }
    return count;
  },

  SUMIF: (args) => {
    const criteriaRange = toCriteriaRange(args[0] ?? null);
    const criteria = parseCriteria(args[1] ?? null);
    const sumRange = toCriteriaRange(args[2] ?? args[0] ?? null);
    if (criteriaRange.columns < 0 || sumRange.columns < 0) return createFormulaError('#VALUE!', 'SUMIF ranges must be rectangular');
    const projectedSumRange = projectCriteriaRange(sumRange, criteriaRange.rows, criteriaRange.columns);
    let sum = 0;
    for (let row = 0; row < criteriaRange.rows; row += 1) {
      for (let column = 0; column < criteriaRange.columns; column += 1) {
        if (!matchesCriteria(criteriaRange.values[row]![column]!, criteria)) continue;
        const numeric = coerceExcelNumber(projectedSumRange.values[row]![column]!);
        if (!isFormulaError(numeric)) sum += numeric;
      }
    }
    return normalizeExcelPrecision(sum);
  },

  AVERAGEIF: (args) => {
    const criteriaRange = toCriteriaRange(args[0] ?? null);
    const criteria = parseCriteria(args[1] ?? null);
    const averageRange = toCriteriaRange(args[2] ?? args[0] ?? null);
    if (criteriaRange.columns < 0 || averageRange.columns < 0) return createFormulaError('#VALUE!', 'AVERAGEIF ranges must be rectangular');
    const projectedAverageRange = projectCriteriaRange(averageRange, criteriaRange.rows, criteriaRange.columns);
    let sum = 0;
    let count = 0;
    for (let row = 0; row < criteriaRange.rows; row += 1) {
      for (let column = 0; column < criteriaRange.columns; column += 1) {
        if (!matchesCriteria(criteriaRange.values[row]![column]!, criteria)) continue;
        const numeric = coerceExcelNumber(projectedAverageRange.values[row]![column]!);
        if (!isFormulaError(numeric)) {
          sum += numeric;
          count += 1;
        }
      }
    }
    return count === 0 ? createFormulaError('#DIV/0!', 'No matching values') : normalizeExcelPrecision(sum / count);
  },
};
