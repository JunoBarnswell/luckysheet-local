import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { findLookupIndex, type LookupMatchMode } from './lookup-engine';

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
      if (Array.isArray(row)) list.push(...row);
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

function compareValues(left: FormulaValue, right: FormulaValue): number {
  if (isFormulaError(left) || isFormulaError(right)) return 0;
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  return String(left ?? '').localeCompare(String(right ?? ''), undefined, { numeric: true, sensitivity: 'base' });
}

function matrixHeight(matrix: FormulaValue[][]): number {
  return matrix.length;
}

function matrixWidth(matrix: FormulaValue[][]): number {
  return Math.max(0, ...matrix.map((row) => row.length));
}

export const dynamicArrayFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  FILTER: (args) => {
    if (args.length < 2) return createFormulaError('#VALUE!', 'FILTER requires array and include');
    const array = to2DArray(args[0]);
    const include = to2DArray(args[1]);
    const ifEmpty = args[2];
    const rows = matrixHeight(array);
    const cols = matrixWidth(array);
    const includeRows = matrixHeight(include);
    const includeCols = matrixWidth(include);
    if (includeRows !== rows && includeRows !== 1) {
      return createFormulaError('#VALUE!', 'FILTER include height mismatch');
    }
    const result: FormulaValue[][] = [];
    for (let row = 0; row < rows; row++) {
      let pass = false;
      if (includeRows === 1) {
        for (let column = 0; column < includeCols; column++) {
          if (isTruthy(include[0]?.[column] ?? false)) pass = true;
        }
      } else {
        for (let column = 0; column < includeCols; column++) {
          if (isTruthy(include[row]?.[column] ?? false)) pass = true;
        }
      }
      if (pass) result.push([...(array[row] ?? Array.from({ length: cols }, () => null))]);
    }
    if (result.length === 0) {
      if (ifEmpty !== undefined) return to2DArray(ifEmpty);
      return createFormulaError('#CALC!', 'FILTER returned no results');
    }
    return result;
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
        if (!exactlyOnce || seen.get(key) === 1) columns.push(colValues);
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

  SORT: (args) => {
    if (args.length < 1) return createFormulaError('#VALUE!', 'SORT requires an array');
    const array = to2DArray(args[0]);
    if (array.length === 0) return [[]];
    const sortIndex = Math.max(1, Number(args[1] ?? 1));
    const sortOrder = Number(args[2] ?? 1) >= 0 ? 1 : -1;
    const byCol = args[3] === true || args[3] === 1;
    const column = sortIndex - 1;

    if (byCol) {
      const width = matrixWidth(array);
      if (column >= width) return createFormulaError('#VALUE!', 'SORT sort_index out of bounds');
      const indices = Array.from({ length: width }, (_, index) => index);
      indices.sort((left, right) => sortOrder * compareValues(array[0]?.[left] ?? null, array[0]?.[right] ?? null));
      return array.map((row) => indices.map((index) => row[index] ?? null));
    }

    const sorted = [...array].sort((left, right) => sortOrder * compareValues(left[column] ?? null, right[column] ?? null));
    return sorted;
  },

  SEQUENCE: (args) => {
    const rows = Number(args[0]);
    const columns = Number(args[1] ?? 1);
    const start = Number(args[2] ?? 1);
    const step = Number(args[3] ?? 1);
    if (!Number.isFinite(rows) || rows < 1) return createFormulaError('#VALUE!', 'SEQUENCE rows must be >= 1');
    if (!Number.isFinite(columns) || columns < 1) return createFormulaError('#VALUE!', 'SEQUENCE columns must be >= 1');
    const sizeError = validateGeneratedArraySize('SEQUENCE', rows, columns);
    if (sizeError) return sizeError;
    const result: FormulaValue[][] = [];
    let value = start;
    for (let row = 0; row < rows; row++) {
      const line: FormulaValue[] = [];
      for (let column = 0; column < columns; column++) {
        line.push(value);
        value += step;
      }
      result.push(line);
    }
    return result;
  },

  XMATCH: (args) => {
    const lookupValue = args[0];
    const lookupArray = to1DArray(args[1]);
    const matchMode = Number(args[2] ?? 0);
    const searchMode = Number(args[3] ?? 1);
    if (lookupArray.length === 0) return createFormulaError('#N/A', 'XMATCH lookup array is empty');

    if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return createFormulaError('#VALUE!', 'Invalid XMATCH mode');
    const index = findLookupIndex(lookupValue, lookupArray, matchMode as LookupMatchMode, searchMode);
    if (index >= 0) return index + 1;
    return createFormulaError('#N/A', 'Value not found in XMATCH');
  },

  HSTACK: (args) => {
    if (args.length === 0) return createFormulaError('#VALUE!', 'HSTACK requires arrays');
    const matrices = args.map((arg) => to2DArray(arg));
    const height = Math.max(...matrices.map((matrix) => matrix.length));
    const result: FormulaValue[][] = [];
    for (let row = 0; row < height; row++) {
      const line: FormulaValue[] = [];
      for (const matrix of matrices) line.push(...(matrix[row] ?? []));
      result.push(line);
    }
    return result;
  },

  VSTACK: (args) => {
    if (args.length === 0) return createFormulaError('#VALUE!', 'VSTACK requires arrays');
    const matrices = args.map((arg) => to2DArray(arg));
    const width = Math.max(...matrices.map((matrix) => matrixWidth(matrix)));
    const result: FormulaValue[][] = [];
    for (const matrix of matrices) {
      for (const row of matrix) {
        const padded = [...row];
        while (padded.length < width) padded.push(null);
        result.push(padded);
      }
    }
    return result;
  },

  TAKE: (args) => {
    if (args.length < 2) return createFormulaError('#VALUE!', 'TAKE requires array and rows');
    const array = to2DArray(args[0]);
    const rows = Number(args[1]);
    const columns = args[2] === undefined ? undefined : Number(args[2]);
    if (!Number.isFinite(rows)) return createFormulaError('#VALUE!', 'TAKE rows must be numeric');
    if (rows >= 0) {
      const sliced = array.slice(0, rows);
      if (columns === undefined) return sliced;
      return sliced.map((row) => row.slice(0, columns));
    }
    const start = Math.max(0, array.length + rows);
    const sliced = array.slice(start);
    if (columns === undefined) return sliced;
    if (columns >= 0) return sliced.map((row) => row.slice(0, columns));
    return sliced.map((row) => row.slice(columns));
  },

  DROP: (args) => {
    if (args.length < 2) return createFormulaError('#VALUE!', 'DROP requires array and rows');
    const array = to2DArray(args[0]);
    const rows = Number(args[1]);
    const columns = args[2] === undefined ? undefined : Number(args[2]);
    if (!Number.isFinite(rows)) return createFormulaError('#VALUE!', 'DROP rows must be numeric');
    const rowSlice = rows >= 0 ? array.slice(rows) : array.slice(0, Math.max(0, array.length + rows));
    if (columns === undefined) return rowSlice;
    if (columns >= 0) return rowSlice.map((row) => row.slice(columns));
    return rowSlice.map((row) => row.slice(0, Math.max(0, row.length + columns)));
  },

  SORTBY: (args) => {
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
        sortOrder = Number(args[argIndex] ?? 1) >= 0 ? 1 : -1;
        argIndex += 1;
      }
      sortKeys.push({ values: byArray, order: sortOrder });
    }
    indices.sort((left, right) => {
      for (const key of sortKeys) {
        const delta = key.order * compareValues(key.values[left] ?? null, key.values[right] ?? null);
        if (delta !== 0) return delta;
      }
      return left - right;
    });
    return indices.map((index) => array[index] ?? []);
  },

  RANDARRAY: (args) => {
    const rows = Number(args[0]);
    const columns = Number(args[1] ?? 1);
    const min = Number(args[2] ?? 0);
    const max = Number(args[3] ?? 1);
    const whole = args[4] === true || args[4] === 1;
    if (!Number.isFinite(rows) || rows < 1) return createFormulaError('#VALUE!', 'RANDARRAY rows must be >= 1');
    if (!Number.isFinite(columns) || columns < 1) return createFormulaError('#VALUE!', 'RANDARRAY columns must be >= 1');
    const sizeError = validateGeneratedArraySize('RANDARRAY', rows, columns);
    if (sizeError) return sizeError;
    const result: FormulaValue[][] = [];
    for (let row = 0; row < rows; row++) {
      const line: FormulaValue[] = [];
      for (let column = 0; column < columns; column++) {
        const value = min + Math.random() * (max - min);
        line.push(whole ? Math.floor(value) : value);
      }
      result.push(line);
    }
    return result;
  },
};
