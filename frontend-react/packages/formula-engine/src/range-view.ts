import type { CellAddress } from './ast';
import type { FormulaEvaluationContext, FormulaEvaluationValue } from './evaluator';
import type { RangeDependency } from './range-index';
import { createFormulaError, isFormulaError, type FormulaError, type FormulaValue } from './values';

export interface RangeEntry {
  readonly row: number;
  readonly column: number;
  readonly address?: CellAddress;
  readonly value: FormulaValue;
}

/** A consumption view over canonical references or an already evaluated array. */
export interface FormulaRangeView {
  readonly rows: number;
  readonly columns: number;
  readonly reference?: RangeDependency;
  read(row: number, column: number): FormulaValue;
  entries(): Iterable<RangeEntry>;
}

export function operandRanges(value: FormulaEvaluationValue): readonly RangeDependency[] | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value) || !('kind' in value)) return undefined;
  return value.kind === 'range' ? [value.range] : value.kind === 'reference' && 'ranges' in value ? value.ranges : undefined;
}

export function createRangeView(value: FormulaEvaluationValue, context?: FormulaEvaluationContext): FormulaRangeView | FormulaError {
  const ranges = operandRanges(value);
  if (ranges) {
    if (!context || ranges.length !== 1) return createFormulaError('#VALUE!', 'This function requires one rectangular reference');
    return referenceView(ranges[0]!, context);
  }
  if (isFormulaError(value)) return value;
  if (typeof value === 'object' && value !== null && !Array.isArray(value)) return createFormulaError('#VALUE!', 'A callable or unresolved reference cannot be consumed as an array');
  const matrix = Array.isArray(value) ? value : [[value]];
  const rows = matrix.length;
  const columns = matrix[0]?.length ?? 0;
  if (matrix.some((row) => row.length !== columns)) return createFormulaError('#VALUE!', 'Array must be rectangular');
  return {
    rows, columns,
    read: (row, column) => matrix[row]?.[column] ?? null,
    *entries() {
      for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) yield { row, column, value: matrix[row]![column]! };
    },
  };
}

function referenceView(range: RangeDependency, context: FormulaEvaluationContext): FormulaRangeView {
  const rows = range.end.row - range.start.row + 1;
  const columns = range.end.column - range.start.column + 1;
  return {
    rows, columns, reference: range,
    read: (row, column) => context.readCell({ sheetId: range.start.sheetId, row: range.start.row + row, column: range.start.column + column }),
    *entries() {
      if (context.readSparseRangeCells) {
        for (const cell of context.readSparseRangeCells(range)) yield { row: cell.address.row - range.start.row, column: cell.address.column - range.start.column, ...cell };
      } else {
        for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
          const address = { sheetId: range.start.sheetId, row: range.start.row + row, column: range.start.column + column };
          yield { row, column, address, value: context.readCell(address) };
        }
      }
    },
  };
}

/** SUMIF/AVERAGEIF extend the target from its top-left canonical cell. */
export function projectRangeView(view: FormulaRangeView, rows: number, columns: number, context?: FormulaEvaluationContext): FormulaRangeView {
  if (view.reference && context) return referenceView({ kind: 'range', start: view.reference.start, end: { sheetId: view.reference.start.sheetId, row: view.reference.start.row + rows - 1, column: view.reference.start.column + columns - 1 } }, context);
  return { rows, columns, read: view.read, *entries() { for (const entry of view.entries()) if (entry.row < rows && entry.column < columns) yield entry; } };
}

export function sameRangeShape(left: FormulaRangeView, right: FormulaRangeView): boolean {
  return left.rows === right.rows && left.columns === right.columns;
}
