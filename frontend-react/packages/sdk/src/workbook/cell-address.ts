import { MAX_SHEET_COLUMN_COUNT, MAX_SHEET_ROW_COUNT, parseAddress } from '@react-sheets/core-model';

export function parseCellAddress(input: string): { row: number; column: number } {
  if (typeof input !== 'string' || !/^\$?[A-Z]+\$?[1-9]\d*$/i.test(input.trim())) throw new Error(`Invalid cell address: ${String(input)}`);
  const parsed = parseAddress(input.trim().replaceAll('$', ''));
  if (!parsed || !Number.isSafeInteger(parsed.row) || !Number.isSafeInteger(parsed.column)
    || parsed.row < 0 || parsed.row >= MAX_SHEET_ROW_COUNT || parsed.column < 0 || parsed.column >= MAX_SHEET_COLUMN_COUNT) throw new Error(`Cell address exceeds Excel bounds: ${input}`);
  return parsed;
}
