import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import type { FormulaEvaluationContext } from '../evaluator';
import { compareWorkbookValues } from '../collation';
import { evaluateRangeFunction } from './range-functions';

const MAX_GENERATED_ARRAY_CELLS = 100_000;

function validateGeneratedArraySize(functionName: string, rows: number, columns: number): FormulaValue | undefined {
  if (rows > MAX_GENERATED_ARRAY_CELLS || columns > Math.floor(MAX_GENERATED_ARRAY_CELLS / rows)) {
    return createFormulaError(
      '#VALUE!',
      `${functionName} output exceeds the ${MAX_GENERATED_ARRAY_CELLS}-cell limit`,
    );
  }
  return undefined;
}

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
      if (Array.isArray(row)) for (const cell of row) list.push(cell);
      else list.push(row);
    }
    return list;
  }
  return [val];
}

function isTruthy(value: FormulaValue): boolean {
  if (isFormulaError(value)) return false;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (value == null || value === '') return false;
  return true;
}

function compareValues(left: FormulaValue, right: FormulaValue, context?: FormulaEvaluationContext): number {
  return compareWorkbookValues(left, right, context?.collationContext);
}

function matrixHeight(matrix: FormulaValue[][]): number {
  return matrix.length;
}

function matrixWidth(matrix: FormulaValue[][]): number {
  return matrix.reduce((width, row) => Math.max(width, row.length), 0);
}

export const dynamicArrayFunctions: Record<string, (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue> = {
  FILTER: (args) => {
    const array = to2DArray(args[0]);
    const include = to2DArray(args[1]);
    const rows = matrixHeight(array);
    const columns = matrixWidth(array);
    const byRow = include.length === rows && matrixWidth(include) === 1;
    const byColumn = include.length === 1 && matrixWidth(include) === columns;
    if (!byRow && !byColumn) return createFormulaError('#VALUE!', 'FILTER include must align with one array dimension');
    const selected: boolean[] = [];
    for (let index = 0; index < (byRow ? rows : columns); index++) {
      const value = include[byRow ? index : 0]?.[byRow ? 0 : index] ?? null;
      if (isFormulaError(value)) return value;
      if (typeof value === 'string' && !['TRUE', 'FALSE'].includes(value.toUpperCase())) return createFormulaError('#VALUE!', 'FILTER include is not logical');
      selected.push(typeof value === 'string' ? value.toUpperCase() === 'TRUE' : Boolean(value));
    }
    if (!selected.some(Boolean)) return args[2] !== undefined ? to2DArray(args[2]) : createFormulaError('#CALC!', 'FILTER returned no results');
    return byRow ? array.filter((_, row) => selected[row]).map((row) => [...row]) : array.map((row) => row.filter((_, column) => selected[column]));
  },

  UNIQUE: (args) => {
    if (args.length < 1) return createFormulaError('#VALUE!', 'UNIQUE requires an array');
    const byCol = args[1] === true || args[1] === 1;
    const exactlyOnce = args[2] === true || args[2] === 1;
    const array = to2DArray(args[0]);
    if (array.length === 0) return [[]];

    if (byCol) {
      const width = matrixWidth(array);
      const seen = new Map<string, number>();
      const columns: FormulaValue[][] = [];
      for (let column = 0; column < width; column++) {
        const colValues = array.map((row) => row[column] ?? null);
        const key = JSON.stringify(colValues);
        seen.set(key, (seen.get(key) ?? 0) + 1);
        if (seen.get(key) === 1) columns.push(colValues);
      }
      if (exactlyOnce) {
        const filtered = columns.filter((col) => seen.get(JSON.stringify(col)) === 1);
        if (filtered.length === 0) return [[]];
        const height = filtered[0]?.length ?? 0;
        return Array.from({ length: height }, (_, row) => filtered.map((col) => col[row] ?? null));
      }
      const height = columns[0]?.length ?? 0;
      return Array.from({ length: height }, (_, row) => columns.map((col) => col[row] ?? null));
    }

    const counts = new Map<string, number>();
    for (const row of array) counts.set(JSON.stringify(row), (counts.get(JSON.stringify(row)) ?? 0) + 1);
    const unique = array.filter((row) => {
      const key = JSON.stringify(row);
      return exactlyOnce ? counts.get(key) === 1 : true;
    });
    const deduped: FormulaValue[][] = [];
    const seen = new Set<string>();
    for (const row of unique) {
      const key = JSON.stringify(row);
      if (seen.has(key)) continue;
      seen.add(key);
      deduped.push(row);
    }
    return deduped.length > 0 ? deduped : [[]];
  },

  SORT: (args, context) => {
    if (args.length < 1) return createFormulaError('#VALUE!', 'SORT requires an array');
    const array = to2DArray(args[0]);
    if (array.length === 0) return [[]];
    const sortIndex = coerceExcelNumber(args[1] ?? 1);
    const sortOrderValue = coerceExcelNumber(args[2] ?? 1);
    if (isFormulaError(sortIndex) || isFormulaError(sortOrderValue)) return isFormulaError(sortIndex) ? sortIndex : sortOrderValue;
    const normalizedSortIndex = Math.trunc(sortIndex);
    if (![1, -1].includes(sortOrderValue) || normalizedSortIndex < 1) return createFormulaError('#VALUE!', 'Invalid SORT index or order');
    const sortOrder = sortOrderValue;
    const byCol = args[3] === true || args[3] === 1;
    const column = normalizedSortIndex - 1;

    if (byCol) {
      const width = matrixWidth(array);
      if (column >= array.length) return createFormulaError('#VALUE!', 'SORT sort_index out of bounds');
      const indices = Array.from({ length: width }, (_, index) => index);
      indices.sort((left, right) => sortOrder * compareValues(array[column]?.[left] ?? null, array[column]?.[right] ?? null, context));
      return array.map((row) => indices.map((index) => row[index] ?? null));
    }

    if (column >= matrixWidth(array)) return createFormulaError('#VALUE!', 'SORT sort_index out of bounds');
    const sorted = [...array].sort((left, right) => sortOrder * compareValues(left[column] ?? null, right[column] ?? null, context));
    return sorted;
  },

  SEQUENCE: (args) => {
    const rows = coerceExcelNumber(args[0] ?? 1);
    const columns = coerceExcelNumber(args[1] ?? 1);
    const start = coerceExcelNumber(args[2] ?? 1);
    const step = coerceExcelNumber(args[3] ?? 1);
    if (isFormulaError(rows) || isFormulaError(columns) || isFormulaError(start) || isFormulaError(step)) return [rows, columns, start, step].find(isFormulaError)!;
    if (!Number.isFinite(rows) || rows < 1) return createFormulaError('#VALUE!', 'SEQUENCE rows must be >= 1');
    if (!Number.isFinite(columns) || columns < 1) return createFormulaError('#VALUE!', 'SEQUENCE columns must be >= 1');
    const sizeError = validateGeneratedArraySize('SEQUENCE', rows, columns);
    if (sizeError) return sizeError;
    const result: FormulaValue[][] = [];
    let value = start;
    for (let row = 0; row < Math.trunc(rows); row++) {
      const line: FormulaValue[] = [];
      for (let column = 0; column < Math.trunc(columns); column++) {
        line.push(normalizeExcelPrecision(value));
        value = normalizeExcelPrecision(value + step);
      }
      result.push(line);
    }
    return result;
  },

  XMATCH: (args, context) => evaluateRangeFunction('XMATCH', args, context) as FormulaValue,

  HSTACK: (args) => {
    if (args.length === 0) return createFormulaError('#VALUE!', 'HSTACK requires arrays');
    const matrices = args.map((arg) => to2DArray(arg));
    const height = matrices.reduce((height, matrix) => Math.max(height, matrix.length), 0);
    const result: FormulaValue[][] = [];
    for (let row = 0; row < height; row++) {
      const line: FormulaValue[] = [];
      for (const matrix of matrices) for (let column = 0; column < matrixWidth(matrix); column++) line.push(matrix[row]?.[column] ?? createFormulaError('#N/A', 'HSTACK input has no value at this position'));
      result.push(line);
    }
    return result;
  },

  VSTACK: (args) => {
    if (args.length === 0) return createFormulaError('#VALUE!', 'VSTACK requires arrays');
    const matrices = args.map((arg) => to2DArray(arg));
    const width = matrices.reduce((width, matrix) => Math.max(width, matrixWidth(matrix)), 0);
    const result: FormulaValue[][] = [];
    for (const matrix of matrices) {
      for (const row of matrix) {
        const padded = [...row];
        while (padded.length < width) padded.push(createFormulaError('#N/A', 'VSTACK input has no value at this position'));
        result.push(padded);
      }
    }
    return result;
  },

  TAKE: (args) => {
    if (args.length < 2) return createFormulaError('#VALUE!', 'TAKE requires array and rows');
    const array = to2DArray(args[0]);
    const rows = coerceExcelNumber(args[1]);
    const columns = args[2] === undefined ? undefined : coerceExcelNumber(args[2]);
    if (isFormulaError(rows) || isFormulaError(columns) || !Number.isFinite(rows)) return isFormulaError(rows) ? rows : isFormulaError(columns) ? columns : createFormulaError('#VALUE!', 'TAKE rows must be numeric');
    if (Math.trunc(rows) === 0 || (columns !== undefined && Math.trunc(columns) === 0)) return createFormulaError('#CALC!', 'TAKE cannot return an empty array');
    const rowSlice = rows >= 0 ? array.slice(0, Math.trunc(rows)) : array.slice(Math.max(0, array.length + Math.trunc(rows)));
    if (columns === undefined) return rowSlice;
    return rowSlice.map((row) => columns >= 0 ? row.slice(0, Math.trunc(columns)) : row.slice(Math.trunc(columns)));
  },

  DROP: (args) => {
    if (args.length < 2) return createFormulaError('#VALUE!', 'DROP requires array and rows');
    const array = to2DArray(args[0]);
    const rows = coerceExcelNumber(args[1]);
    const columns = args[2] === undefined ? undefined : coerceExcelNumber(args[2]);
    if (isFormulaError(rows) || isFormulaError(columns) || !Number.isFinite(rows)) return isFormulaError(rows) ? rows : isFormulaError(columns) ? columns : createFormulaError('#VALUE!', 'DROP rows must be numeric');
    const rowSlice = rows >= 0 ? array.slice(rows) : array.slice(0, Math.max(0, array.length + rows));
    const result = columns === undefined ? rowSlice : rowSlice.map((row) => columns >= 0 ? row.slice(Math.trunc(columns)) : row.slice(0, Math.max(0, row.length + Math.trunc(columns))));
    return !result.length || !matrixWidth(result) ? createFormulaError('#CALC!', 'DROP cannot return an empty array') : result;
  },

  SORTBY: (args, context) => {
    if (args.length < 2) return createFormulaError('#VALUE!', 'SORTBY requires array and by_array');
    const array = to2DArray(args[0]);
    if (array.length === 0) return [[]];
    const indices = array.map((_, index) => index);
    const sortKeys: Array<{ values: FormulaValue[]; order: number }> = [];
    let argIndex = 1;
    while (argIndex < args.length) {
      const byArray = to1DArray(args[argIndex]);
      argIndex += 1;
      let sortOrder = 1;
      if (argIndex < args.length && !Array.isArray(args[argIndex])) {
        const order = coerceExcelNumber(args[argIndex] ?? 1);
        if (isFormulaError(order)) return order;
        if (![1, -1].includes(order)) return createFormulaError('#VALUE!', 'SORTBY order must be 1 or -1');
        sortOrder = order;
        argIndex += 1;
      }
      if (byArray.length !== array.length) return createFormulaError('#VALUE!', 'SORTBY keys must align with the array');
      sortKeys.push({ values: byArray, order: sortOrder });
    }
    indices.sort((left, right) => {
      for (const key of sortKeys) {
        const delta = key.order * compareValues(key.values[left] ?? null, key.values[right] ?? null, context);
        if (delta !== 0) return delta;
      }
      return left - right;
    });
    return indices.map((index) => array[index] ?? []);
  },

  RANDARRAY: (args, context) => {
    const rows = coerceExcelNumber(args[0] ?? 1);
    const columns = coerceExcelNumber(args[1] ?? 1);
    const min = coerceExcelNumber(args[2] ?? 0);
    const max = coerceExcelNumber(args[3] ?? 1);
    const whole = args[4] === true || args[4] === 1;
    if (isFormulaError(rows) || isFormulaError(columns) || isFormulaError(min) || isFormulaError(max)) return [rows, columns, min, max].find(isFormulaError)!;
    if (!Number.isFinite(rows) || rows < 1) return createFormulaError('#VALUE!', 'RANDARRAY rows must be >= 1');
    if (!Number.isFinite(columns) || columns < 1) return createFormulaError('#VALUE!', 'RANDARRAY columns must be >= 1');
    if (min > max || (whole && Math.ceil(min) > Math.floor(max))) return createFormulaError('#VALUE!', 'RANDARRAY minimum exceeds maximum');
    const sizeError = validateGeneratedArraySize('RANDARRAY', rows, columns);
    if (sizeError) return sizeError;
    const result: FormulaValue[][] = [];
    for (let row = 0; row < Math.trunc(rows); row++) {
      const line: FormulaValue[] = [];
      for (let column = 0; column < Math.trunc(columns); column++) {
        const random = context?.random?.('RANDARRAY', context?.volatileOccurrence, row * columns + column);
        if (random === undefined || isFormulaError(random)) return random ?? createFormulaError('#BLOCKED!', 'RANDARRAY requires a calculation entropy context');
        const value = min + random * (max - min);
        line.push(whole ? Math.floor(random * (Math.floor(max) - Math.ceil(min) + 1)) + Math.ceil(min) : normalizeExcelPrecision(value));
      }
      result.push(line);
    }
    return result;
  },
};
