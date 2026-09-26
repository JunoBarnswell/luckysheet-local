/** Canonical worksheet coordinate bounds and axis-reference mapping semantics. */
export const MAX_ROW_INDEX = 1_048_575;
export const MAX_COLUMN_INDEX = 16_383;

export interface StructuralShift {
  axis: 'row' | 'column';
  at: number;
  count: number;
  op: 'insert' | 'delete';
}

export interface CellShiftSelection {
  readonly startRow: number;
  readonly endRow: number;
  readonly startColumn: number;
  readonly endColumn: number;
}

export type CellShiftPointTransformResult =
  | { readonly kind: 'mapped'; readonly row: number; readonly column: number }
  | { readonly kind: 'deleted' }
  | { readonly kind: 'out-of-bounds'; readonly row: number; readonly column: number };

export type PointTransformResult =
  | { readonly kind: 'mapped'; readonly position: number }
  | { readonly kind: 'deleted' }
  | { readonly kind: 'out-of-bounds'; readonly position: number };

export type IntervalTransformResult =
  | { readonly kind: 'mapped'; readonly start: number; readonly end: number }
  | { readonly kind: 'deleted' }
  | { readonly kind: 'out-of-bounds'; readonly start: number; readonly end: number };

export class ReferenceTransformDomain {
  static createRowPermutationMap(startRow: number, sourceRowsByTarget: readonly number[]): readonly number[] {
    if (!Number.isSafeInteger(startRow) || startRow < 0
      || startRow > MAX_ROW_INDEX
      || !Array.isArray(sourceRowsByTarget)
      || sourceRowsByTarget.length < 1
      || sourceRowsByTarget.length > MAX_ROW_INDEX - startRow + 1) {
      throw new Error('Reference transform row-permutation inputs are invalid');
    }
    const count = sourceRowsByTarget.length;
    const targetRowsBySource = new Array<number>(count);
    for (let targetOffset = 0; targetOffset < count; targetOffset += 1) {
      const sourceRow = sourceRowsByTarget[targetOffset];
      if (!Number.isSafeInteger(sourceRow) || sourceRow! < startRow || sourceRow! >= startRow + count) {
        throw new Error('Reference transform row-permutation source is outside its range');
      }
      const sourceOffset = sourceRow! - startRow;
      if (targetRowsBySource[sourceOffset] !== undefined) {
        throw new Error('Reference transform row-permutation contains a duplicate source');
      }
      targetRowsBySource[sourceOffset] = startRow + targetOffset;
    }
    if (targetRowsBySource.some((targetRow) => targetRow === undefined)) {
      throw new Error('Reference transform row-permutation must contain every source');
    }
    return Object.freeze(targetRowsBySource);
  }

  static mapPermutationIndex(position: number, startRow: number, targetRowsBySource: readonly number[]): number {
    if (!Number.isSafeInteger(position) || position < 0 || position > MAX_ROW_INDEX
      || !Number.isSafeInteger(startRow) || startRow < 0 || startRow > MAX_ROW_INDEX
      || !Array.isArray(targetRowsBySource) || targetRowsBySource.length < 1
      || targetRowsBySource.length > MAX_ROW_INDEX - startRow + 1) {
      throw new Error('Reference transform row-permutation mapping inputs are invalid');
    }
    const endExclusive = startRow + targetRowsBySource.length;
    if (position < startRow || position >= endExclusive) return position;
    const mapped = targetRowsBySource[position - startRow];
    if (!Number.isSafeInteger(mapped) || mapped! < startRow || mapped! >= endExclusive) {
      throw new Error('Reference transform row-permutation mapping is invalid');
    }
    return mapped!;
  }

  static mapCellShiftIndex(position: number, start: number, end: number, op: 'insert' | 'delete', maximum: number): number {
    if (!Number.isSafeInteger(position) || position < 0
      || !Number.isSafeInteger(start) || start < 0
      || !Number.isSafeInteger(end) || end < start
      || !Number.isSafeInteger(maximum) || maximum < 0 || end > maximum
      || (op !== 'insert' && op !== 'delete')) {
      throw new Error('Reference transform cell-shift index inputs are invalid');
    }
    if (position > maximum) return position;
    if (op === 'delete' && position >= start && position <= end) return -1;
    if (position < start) return position;
    return position + (op === 'insert' ? 1 : -1) * (end - start + 1);
  }

  static mapCellShiftPoint(
    row: number,
    column: number,
    selection: CellShiftSelection,
    axis: 'row' | 'column',
    op: 'insert' | 'delete',
  ): CellShiftPointTransformResult {
    if (!Number.isSafeInteger(row) || row < 0
      || !Number.isSafeInteger(column) || column < 0
      || !selection
      || !Number.isSafeInteger(selection.startRow) || selection.startRow < 0
      || !Number.isSafeInteger(selection.endRow) || selection.endRow < selection.startRow || selection.endRow > MAX_ROW_INDEX
      || !Number.isSafeInteger(selection.startColumn) || selection.startColumn < 0
      || !Number.isSafeInteger(selection.endColumn) || selection.endColumn < selection.startColumn || selection.endColumn > MAX_COLUMN_INDEX
      || (axis !== 'row' && axis !== 'column')
      || (op !== 'insert' && op !== 'delete')) {
      throw new Error('Reference transform cell-shift inputs are invalid');
    }
    const maximumRow = MAX_ROW_INDEX;
    const maximumColumn = MAX_COLUMN_INDEX;
    if (row > maximumRow || column > maximumColumn) return { kind: 'out-of-bounds', row, column };

    const start = axis === 'row' ? selection.startRow : selection.startColumn;
    const end = axis === 'row' ? selection.endRow : selection.endColumn;
    const position = axis === 'row' ? row : column;
    const inBand = axis === 'row'
      ? column >= selection.startColumn && column <= selection.endColumn && row >= start
      : row >= selection.startRow && row <= selection.endRow && column >= start;
    if (!inBand) return { kind: 'mapped', row, column };
    const maximum = axis === 'row' ? maximumRow : maximumColumn;
    const mappedPosition = this.mapCellShiftIndex(position, start, end, op, maximum);
    if (mappedPosition === -1) return { kind: 'deleted' };
    if (mappedPosition > maximum) {
      return {
        kind: 'out-of-bounds',
        row: axis === 'row' ? mappedPosition : row,
        column: axis === 'column' ? mappedPosition : column,
      };
    }
    return axis === 'row'
      ? { kind: 'mapped', row: mappedPosition, column }
      : { kind: 'mapped', row, column: mappedPosition };
  }

  static mapPoint(position: number, at: number, count: number, direction: 1 | -1, maximum: number): PointTransformResult {
    if (!Number.isSafeInteger(position) || position < 0
      || !Number.isSafeInteger(at) || at < 0
      || !Number.isSafeInteger(count) || count < 1
      || (direction !== 1 && direction !== -1)
      || !Number.isSafeInteger(maximum) || maximum < 0 || at > maximum || count > maximum + 1) {
      throw new Error('Reference transform point inputs are invalid');
    }
    if (position > maximum) return { kind: 'out-of-bounds', position };
    if (direction === -1 && position >= at && position - at < count) return { kind: 'deleted' };
    const mapped = direction === 1
      ? position >= at ? position + count : position
      : position >= at ? position - count : position;
    return Number.isSafeInteger(mapped) && mapped <= maximum
      ? { kind: 'mapped', position: mapped }
      : { kind: 'out-of-bounds', position: mapped };
  }

  static mapInterval(
    start: number,
    end: number,
    shift: StructuralShift,
  ): IntervalTransformResult {
    if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(end) || end < 0
      || !shift || (shift.axis !== 'row' && shift.axis !== 'column')
      || !Number.isSafeInteger(shift.at) || shift.at < 0
      || !Number.isSafeInteger(shift.count) || shift.count < 1
      || (shift.op !== 'insert' && shift.op !== 'delete')) {
      throw new Error('Reference transform interval inputs are invalid');
    }
    const maximum = shift.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
    if (start > maximum || end > maximum || shift.at > maximum || shift.count > maximum + 1
      || (shift.op === 'delete' && shift.count - 1 > maximum - shift.at)) {
      throw new Error('Reference transform interval inputs exceed worksheet bounds');
    }
    const low = Math.min(start, end);
    const high = Math.max(start, end);
    let interval: { readonly start: number; readonly end: number } | undefined;
    if (shift.op === 'insert') {
      if (shift.at <= low) interval = { start: low + shift.count, end: high + shift.count };
      else if (shift.at <= high) interval = { start: low, end: high + shift.count };
      else interval = { start: low, end: high };
    } else {
      const deletedEnd = shift.at + shift.count - 1;
      if (high < shift.at) interval = { start: low, end: high };
      else if (low > deletedEnd) interval = { start: low - shift.count, end: high - shift.count };
      else {
        const nextStart = low < shift.at ? low : shift.at;
        const nextEnd = high > deletedEnd ? high - shift.count : shift.at - 1;
        if (nextStart <= nextEnd) interval = { start: nextStart, end: nextEnd };
      }
    }
    if (!interval) return { kind: 'deleted' };
    return Number.isSafeInteger(interval.start) && Number.isSafeInteger(interval.end)
      && interval.start >= 0 && interval.end <= maximum
      ? { kind: 'mapped', start: interval.start, end: interval.end }
      : { kind: 'out-of-bounds', start: interval.start, end: interval.end };
  }
}
