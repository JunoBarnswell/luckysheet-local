import type { RangeRef } from '@react-sheets/core-model';

export function parseRangeInput(input: string, sheetId: string): Omit<RangeRef, 'sheetId'> | undefined {
  const parseCell = (value: string) => {
    const match = /^([A-Z]+)(\d+)$/.exec(value.trim().toUpperCase());
    if (!match?.[1] || !match[2]) return undefined;
    const column = [...match[1]].reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1;
    const row = Number(match[2]) - 1;
    return Number.isInteger(row) && row >= 0 && column >= 0 ? { row, column } : undefined;
  };
  const [startText, endText = startText] = input.split(':');
  const start = parseCell(startText ?? '');
  const end = parseCell(endText ?? '');
  if (!start || !end) return undefined;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  };
}
