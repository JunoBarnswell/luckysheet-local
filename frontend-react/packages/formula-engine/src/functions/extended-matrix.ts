import type { FormulaValue } from '../values';
import { createFormulaError } from '../values';

type ScalarMatrix = FormulaValue[][];

function toMatrix(value: FormulaValue): ScalarMatrix {
  if (Array.isArray(value)) {
    if (value.length > 0 && Array.isArray(value[0])) return value as ScalarMatrix;
    return (value as FormulaValue[]).map((v) => [v]);
  }
  return [[value]];
}

function flattenColumn(matrix: ScalarMatrix, colIndex: number): FormulaValue[] {
  return matrix.map((row) => row[colIndex] ?? null);
}

function uniqueValues(values: FormulaValue[]): FormulaValue[] {
  const seen = new Set<string>();
  const result: FormulaValue[] = [];
  for (const v of values) {
    const key = JSON.stringify(v);
    if (!seen.has(key)) {
      seen.add(key);
      result.push(v);
    }
  }
  return result;
}

/** Capability-gated GROUPBY implementation — 按行分组聚合. */
function GROUPBY(args: FormulaValue[]): FormulaValue {
  if (args.length < 3) return createFormulaError('#VALUE!', 'GROUPBY requires at least 3 arguments');
  const rowGroups = toMatrix(args[0]!);
  const values = toMatrix(args[1]!);
  const functionIndex = args[2];
  if (typeof functionIndex !== 'number') return createFormulaError('#VALUE!', 'function_num must be a number');

  const groupCol = 0;
  const groups = uniqueValues(flattenColumn(rowGroups, groupCol));
  const aggFn = functionIndex === 1 ? (nums: number[]) => nums.reduce((a, b) => a + b, 0) : (nums: number[]) => nums.length;

  const result: FormulaValue[][] = groups.map((g) => {
    const indices = flattenColumn(rowGroups, groupCol)
      .map((v, i) => (v === g ? i : -1))
      .filter((i) => i >= 0);
    const nums = indices.map((i) => {
      const v = values[i]?.[0];
      return typeof v === 'number' ? v : 0;
    });
    return [g, aggFn(nums)];
  });

  return result;
}

/** Capability-gated PIVOTBY implementation — 行列双维度透视. */
function PIVOTBY(args: FormulaValue[]): FormulaValue {
  if (args.length < 4) return createFormulaError('#VALUE!', 'PIVOTBY requires at least 4 arguments');
  const rowFields = toMatrix(args[0]!);
  const colFields = toMatrix(args[1]!);
  const values = toMatrix(args[2]!);
  const _functionIndex = args[3];

  const rowKeys = uniqueValues(flattenColumn(rowFields, 0));
  const colKeys = uniqueValues(flattenColumn(colFields, 0));

  const header: FormulaValue[] = ['', ...colKeys];
  const body: FormulaValue[][] = rowKeys.map((rk) => {
    const row: FormulaValue[] = [rk];
    for (const ck of colKeys) {
      let sum = 0;
      for (let i = 0; i < rowFields.length; i++) {
        if (flattenColumn(rowFields, 0)[i] === rk && flattenColumn(colFields, 0)[i] === ck) {
          const v = values[i]?.[0];
          sum += typeof v === 'number' ? v : 0;
        }
      }
      row.push(sum);
    }
    return row;
  });

  return [header, ...body];
}

export const extendedMatrixFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  GROUPBY,
  PIVOTBY,
};
