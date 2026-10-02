import type { FormulaEvaluationContext, FormulaEvaluationValue } from '../evaluator';
import { createRangeView, operandRanges, sameRangeShape, type FormulaRangeView } from '../range-view';
import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import { createFormulaError, isFormulaError, type FormulaError, type FormulaValue } from '../values';
import { findLookupIndex, type LookupMatchMode, type LookupVector } from './lookup-engine';
import { CONDITIONAL_FUNCTIONS, evaluateConditional } from './conditional';

export const RANGE_CONSUMER_FUNCTIONS = new Set([...CONDITIONAL_FUNCTIONS, 'INDEX', 'ROWS', 'COLUMNS', 'MATCH', 'XMATCH', 'XLOOKUP', 'VLOOKUP', 'HLOOKUP', 'COUNTBLANK', 'MEDIAN', 'LARGE', 'SMALL', 'SUMPRODUCT']);

function scalar(value: FormulaEvaluationValue | undefined, context?: FormulaEvaluationContext): FormulaValue {
  if (value === undefined) return null;
  const view = createRangeView(value, context);
  return isFormulaError(view) ? view : view.rows === 1 && view.columns === 1 ? view.read(0, 0) : createFormulaError('#VALUE!', 'Argument must be scalar');
}

function integer(value: FormulaEvaluationValue | undefined, fallback: number, context?: FormulaEvaluationContext): number | FormulaError {
  const result = value === undefined ? fallback : coerceExcelNumber(scalar(value, context));
  return isFormulaError(result) ? result : Math.trunc(result);
}

function vector(view: FormulaRangeView): LookupVector | FormulaError {
  if (view.rows !== 1 && view.columns !== 1) return createFormulaError('#VALUE!', 'Lookup array must be one row or one column');
  return { length: view.rows === 1 ? view.columns : view.rows, at: (index) => view.read(view.rows === 1 ? 0 : index, view.rows === 1 ? index : 0) };
}

export function evaluateRangeFunction(name: string, args: readonly FormulaEvaluationValue[], context?: FormulaEvaluationContext): FormulaEvaluationValue {
  if (CONDITIONAL_FUNCTIONS.has(name)) return evaluateConditional(name, args, context);
  if (name === 'SUMPRODUCT') {
    if (!args.length) return createFormulaError('#VALUE!', 'SUMPRODUCT requires arguments');
    const views: FormulaRangeView[] = [];
    for (const arg of args) { const view = createRangeView(arg, context); if (isFormulaError(view)) return view; views.push(view); }
    if (views.some((view) => !sameRangeShape(view, views[0]!))) return createFormulaError('#VALUE!', 'SUMPRODUCT arrays must have identical shape');
    // The first operand supplies all nonzero candidates; blanks multiply to zero.
    // Errors in every operand still propagate, including positions with a zero.
    for (const view of views) for (const { value } of view.entries()) if (isFormulaError(value)) return value;
    let total = 0;
    for (const entry of views[0]!.entries()) {
      if (typeof entry.value !== 'number') continue;
      let product = entry.value;
      for (const view of views.slice(1)) { const value = view.read(entry.row, entry.column); product *= typeof value === 'number' ? value : 0; }
      total += product;
    }
    return Number.isFinite(total) ? normalizeExcelPrecision(total) : createFormulaError('#NUM!', 'SUMPRODUCT result is not finite');
  }
  const firstIndex = ['MATCH', 'XMATCH', 'XLOOKUP', 'VLOOKUP', 'HLOOKUP'].includes(name) ? 1 : 0;
  const ranges = args[firstIndex] === undefined ? undefined : operandRanges(args[firstIndex]!);
  const area = name === 'INDEX' ? integer(args[3], 1, context) : 1;
  if (isFormulaError(area)) return area;
  if (area < 1 || (ranges ? area > ranges.length : area !== 1)) return createFormulaError('#REF!', 'INDEX area is out of bounds');
  const source = ranges && name === 'INDEX' ? { kind: 'range' as const, range: ranges[area - 1]! } : args[firstIndex] ?? null;
  const view = createRangeView(source, context);
  if (isFormulaError(view)) return view;
  if (name === 'ROWS') return view.rows;
  if (name === 'COLUMNS') return view.columns;
  if (name === 'COUNTBLANK') {
    let nonblank = 0;
    for (const { value } of view.entries()) if (value !== null && value !== '') nonblank++;
    return view.rows * view.columns - nonblank;
  }
  if (name === 'MEDIAN' || name === 'LARGE' || name === 'SMALL') {
    const numbers: number[] = [];
    const operands = name === 'MEDIAN' ? args : [args[0]!];
    for (const operand of operands) {
      const data = createRangeView(operand, context);
      if (isFormulaError(data)) return data;
      for (const { value } of data.entries()) { if (isFormulaError(value)) return value; if (typeof value === 'number') numbers.push(value); }
    }
    if (!numbers.length) return createFormulaError('#NUM!', 'No numeric values');
    numbers.sort((left, right) => left - right);
    if (name === 'MEDIAN') { const middle = Math.floor(numbers.length / 2); return normalizeExcelPrecision(numbers.length % 2 ? numbers[middle]! : (numbers[middle - 1]! + numbers[middle]!) / 2); }
    const k = integer(args[1], 0, context);
    if (isFormulaError(k)) return k;
    return k < 1 || k > numbers.length ? createFormulaError('#NUM!', 'Order statistic is out of bounds') : numbers[name === 'LARGE' ? numbers.length - k : k - 1]!;
  }
  if (name === 'INDEX') {
    let row = integer(args[1], 1, context);
    let column = integer(args[2], 1, context);
    if (isFormulaError(row) || isFormulaError(column)) return isFormulaError(row) ? row : column;
    if (args[2] === undefined && view.rows === 1) { column = row; row = 1; }
    if (row < 0 || column < 0 || row > view.rows || column > view.columns) return createFormulaError('#REF!', 'INDEX is out of bounds');
    if (view.reference) {
      const range = view.reference;
      return { kind: 'range', range: { kind: 'range', start: { sheetId: range.start.sheetId, row: range.start.row + (row ? row - 1 : 0), column: range.start.column + (column ? column - 1 : 0) }, end: { sheetId: range.start.sheetId, row: row ? range.start.row + row - 1 : range.end.row, column: column ? range.start.column + column - 1 : range.end.column } } };
    }
    if (row && column) return view.read(row - 1, column - 1);
    return Array.from({ length: row ? 1 : view.rows }, (_, r) => Array.from({ length: column ? 1 : view.columns }, (_, c) => view.read(row ? row - 1 : r, column ? column - 1 : c)));
  }
  const lookup = scalar(args[0], context);
  if (isFormulaError(lookup)) return lookup;
  let lookupVector: LookupVector | FormulaError;
  let mode: number | FormulaError;
  let search: number | FormulaError = 1;
  let resultView: FormulaRangeView | undefined;
  let resultIndex: number | FormulaError = 1;
  if (name === 'VLOOKUP' || name === 'HLOOKUP') {
    const horizontal = name === 'HLOOKUP';
    lookupVector = { length: horizontal ? view.columns : view.rows, at: (index) => view.read(horizontal ? 0 : index, horizontal ? index : 0) };
    resultIndex = integer(args[2], 0, context);
    if (isFormulaError(resultIndex)) return resultIndex;
    if (resultIndex < 1 || resultIndex > (horizontal ? view.rows : view.columns)) return createFormulaError('#REF!', 'Lookup result index is out of bounds');
    const approximate = args[3] === undefined ? true : scalar(args[3], context);
    if (isFormulaError(approximate)) return approximate;
    mode = approximate === false || approximate === 0 || String(approximate).toUpperCase() === 'FALSE' ? 0 : -1;
  } else {
    lookupVector = vector(view);
    mode = integer(args[name === 'XLOOKUP' ? 4 : 2], name === 'MATCH' ? 1 : 0, context);
    search = name === 'MATCH' ? 1 : integer(args[name === 'XLOOKUP' ? 5 : 3], 1, context);
    if (name === 'MATCH' && !isFormulaError(mode)) mode = mode === 1 ? -1 : mode === -1 ? 1 : mode;
    if (name === 'XLOOKUP') {
      const result = createRangeView(args[2] ?? null, context);
      if (isFormulaError(result)) return result;
      resultView = result;
      if (view.rows === 1 ? result.columns !== view.columns : result.rows !== view.rows) return createFormulaError('#VALUE!', 'XLOOKUP return array must align with the lookup vector');
    }
  }
  if (isFormulaError(lookupVector)) return lookupVector;
  if (isFormulaError(mode)) return mode;
  if (isFormulaError(search)) return search;
  if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search) || (mode === 2 && Math.abs(search) === 2)) return createFormulaError('#VALUE!', 'Invalid lookup mode');
  const index = findLookupIndex(lookup, lookupVector, mode as LookupMatchMode, search, context?.collationContext);
  if (index < 0) return name === 'XLOOKUP' && args[3] !== undefined ? args[3] : createFormulaError('#N/A', 'Lookup value was not found');
  if (name === 'MATCH' || name === 'XMATCH') return index + 1;
  if (name === 'VLOOKUP') return view.read(index, (resultIndex as number) - 1);
  if (name === 'HLOOKUP') return view.read((resultIndex as number) - 1, index);
  if (view.rows === 1) return resultView!.rows === 1 ? resultView!.read(0, index) : Array.from({ length: resultView!.rows }, (_, row) => [resultView!.read(row, index)]);
  return resultView!.columns === 1 ? resultView!.read(index, 0) : [Array.from({ length: resultView!.columns }, (_, column) => resultView!.read(index, column))];
}
