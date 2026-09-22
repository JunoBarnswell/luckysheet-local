import type { CellAddress } from './ast';
import { createFormulaError, isArrayValue, isFormulaError, type ArrayValue, type FormulaValue } from './values';
import type { SpillRange, SpillState } from './spill';

export interface SpillResolveInput {
  sheetId: string;
  anchor: { row: number; column: number };
  values: ArrayValue;
  rowCount: number;
  columnCount: number;
  isOccupied: (row: number, column: number) => boolean;
}

export interface ResolvedSpill {
  sheetId: string;
  anchor: { row: number; column: number };
  range: SpillRange['range'];
  values: SpillRange['values'];
  state: SpillState;
  blocker?: { row: number; column: number };
}

export function isSpillMatrix(value: FormulaValue): value is ArrayValue {
  if (!isArrayValue(value)) return false;
  const rows = value.length;
  const columns = Math.max(0, ...value.map((row) => row.length));
  return rows * columns > 1;
}

export function resolveSpill(input: SpillResolveInput): ResolvedSpill {
  const height = input.values.length;
  const width = Math.max(0, ...input.values.map((row) => row.length));
  const endRow = input.anchor.row + height - 1;
  const endColumn = input.anchor.column + width - 1;
  const range = {
    sheetId: input.sheetId,
    startRow: input.anchor.row,
    endRow: Math.min(endRow, input.rowCount - 1),
    startColumn: input.anchor.column,
    endColumn: Math.min(endColumn, input.columnCount - 1),
  };

  if (endRow >= input.rowCount || endColumn >= input.columnCount) {
    return {
      sheetId: input.sheetId,
      anchor: { ...input.anchor },
      range,
      values: toCoreMatrix(input.values),
      state: 'spill-error',
    };
  }

  for (let row = input.anchor.row; row <= endRow; row++) {
    for (let column = input.anchor.column; column <= endColumn; column++) {
      if (row === input.anchor.row && column === input.anchor.column) continue;
      if (input.isOccupied(row, column)) {
        return {
          sheetId: input.sheetId,
          anchor: { ...input.anchor },
          range,
          values: toCoreMatrix(input.values),
          state: 'blocked',
          blocker: { row, column },
        };
      }
    }
  }

  return {
    sheetId: input.sheetId,
    anchor: { ...input.anchor },
    range,
    values: toCoreMatrix(input.values),
    state: 'ok',
  };
}

export function spillValueAt(spill: ResolvedSpill | SpillRange, row: number, column: number): FormulaValue | undefined {
  const relRow = row - spill.anchor.row;
  const relColumn = column - spill.anchor.column;
  if (relRow < 0 || relColumn < 0) return undefined;
  if (row < spill.range.startRow || row > spill.range.endRow || column < spill.range.startColumn || column > spill.range.endColumn) {
    return undefined;
  }
  if (spill.state === 'blocked' && row === spill.anchor.row && column === spill.anchor.column) {
    return createFormulaError('#SPILL!', 'Spill range is not blank');
  }
  const raw = spill.values[relRow]?.[relColumn];
  return raw === undefined ? undefined : fromCoreValue(raw);
}

export function isSpillChild(spill: SpillRange, row: number, column: number): boolean {
  if (row === spill.anchor.row && column === spill.anchor.column) return false;
  return row >= spill.range.startRow
    && row <= spill.range.endRow
    && column >= spill.range.startColumn
    && column <= spill.range.endColumn;
}

export function anchorDisplayValue(spill: ResolvedSpill | SpillRange, matrix: ArrayValue): FormulaValue {
  if (spill.state === 'blocked') return createFormulaError('#SPILL!', 'Spill range is not blank');
  return matrix[0]?.[0] ?? null;
}

function toCoreMatrix(matrix: ArrayValue): SpillRange['values'] {
  return matrix.map((row) =>
    row.map((value) => {
      if (isFormulaError(value)) return { kind: 'error' as const, code: value.code, message: value.message };
      return value as SpillRange['values'][number][number];
    }),
  );
}

function fromCoreValue(value: SpillRange['values'][number][number]): FormulaValue {
  if (typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'error') {
    return createFormulaError((value as { code: import('./values').FormulaErrorCode }).code, (value as { message?: string }).message ?? '');
  }
  return value as FormulaValue;
}

export function spillKey(address: CellAddress): string {
  return `${address.sheetId}:${address.row}:${address.column}`;
}
