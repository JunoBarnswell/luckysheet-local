import type { RangeRef } from '@react-sheets/core-model';

/** The only clear semantics accepted by the worksheet range command. */
export type ClearFamily = 'contents' | 'formats' | 'all' | 'comments-and-notes' | 'hyperlinks';

export interface ClearRangeParams {
  sheetId: string;
  range: RangeRef;
  family: ClearFamily;
}

export interface ClearRangePlan {
  params: ClearRangeParams;
  range: RangeRef;
}

function normalizeRange(range: RangeRef): RangeRef {
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new Error('Clear range is invalid');
  }
  return {
    ...range,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

export function createClearRangePlan(sheet: { id: string }, input: ClearRangeParams): ClearRangePlan {
  const range = normalizeRange(input.range);
  if (range.sheetId !== sheet.id || input.sheetId !== sheet.id) throw new Error('Clear range targets another worksheet');
  return { params: { ...input, range }, range };
}
