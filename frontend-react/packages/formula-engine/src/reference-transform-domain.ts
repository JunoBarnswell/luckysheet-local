/** Canonical worksheet coordinate bounds and axis-reference mapping semantics. */
export const MAX_ROW_INDEX = 1_048_575;
export const MAX_COLUMN_INDEX = 16_383;

export interface StructuralShift {
  axis: 'row' | 'column';
  at: number;
  count: number;
  op: 'insert' | 'delete';
}

export type PointTransformResult =
  | { readonly kind: 'mapped'; readonly position: number }
  | { readonly kind: 'deleted' }
  | { readonly kind: 'out-of-bounds'; readonly position: number };

export type IntervalTransformResult =
  | { readonly kind: 'mapped'; readonly start: number; readonly end: number }
  | { readonly kind: 'deleted' }
  | { readonly kind: 'out-of-bounds'; readonly start: number; readonly end: number };

export class ReferenceTransformDomain {
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
