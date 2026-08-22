import type { CellAddress, ParsedCellReference } from './ast';
import { FormulaReferenceError } from './errors';

const CELL_REFERENCE_PATTERN = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/;

export function tryParseCellReferenceText(text: string): ParsedCellReference | undefined {
  const match = CELL_REFERENCE_PATTERN.exec(text);
  if (!match) return undefined;

  const [, absoluteColumnMarker, columnText, absoluteRowMarker, rowText] = match;
  if (!columnText || !rowText) return undefined;

  const rowNumber = Number(rowText);
  const columnNumber = columnNameToIndex(columnText);
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 1 || columnNumber === undefined) return undefined;

  return {
    row: rowNumber - 1,
    column: columnNumber,
    absoluteRow: absoluteRowMarker === '$',
    absoluteColumn: absoluteColumnMarker === '$',
  };
}

export function parseCellReferenceText(text: string): ParsedCellReference {
  const reference = tryParseCellReferenceText(text);
  if (!reference) throw new FormulaReferenceError(`Invalid cell reference: ${text}`);
  return reference;
}

export function parseCellAddress(input: string, defaultSheetId = 'Sheet1'): CellAddress {
  const text = input.trim();
  if (!text) throw new FormulaReferenceError('Cell address cannot be empty');

  const separatorIndex = text.lastIndexOf('!');
  const rawSheetId = separatorIndex < 0 ? undefined : text.slice(0, separatorIndex);
  const cellText = separatorIndex < 0 ? text : text.slice(separatorIndex + 1);
  const sheetId = rawSheetId === undefined ? defaultSheetId : unquoteSheetId(rawSheetId);
  if (!sheetId) throw new FormulaReferenceError(`Invalid sheet in cell address: ${input}`);

  const reference = parseCellReferenceText(cellText);
  return { sheetId, row: reference.row, column: reference.column };
}

export function assertCellAddress(address: CellAddress): void {
  if (!address.sheetId || !Number.isSafeInteger(address.row) || address.row < 0) {
    throw new FormulaReferenceError(`Invalid cell row or sheet: ${JSON.stringify(address)}`);
  }
  if (!Number.isSafeInteger(address.column) || address.column < 0) {
    throw new FormulaReferenceError(`Invalid cell column: ${JSON.stringify(address)}`);
  }
}

export function cellAddressKey(address: CellAddress): string {
  assertCellAddress(address);
  return JSON.stringify([address.sheetId, address.row, address.column]);
}

export function compareCellAddresses(left: CellAddress, right: CellAddress): number {
  const sheetComparison = left.sheetId.localeCompare(right.sheetId);
  if (sheetComparison !== 0) return sheetComparison;
  if (left.row !== right.row) return left.row - right.row;
  return left.column - right.column;
}

export function formatCellAddress(address: CellAddress, includeSheetId = false): string {
  assertCellAddress(address);
  const cell = `${columnIndexToName(address.column)}${address.row + 1}`;
  return includeSheetId ? `${quoteSheetId(address.sheetId)}!${cell}` : cell;
}

export function columnNameToIndex(columnName: string): number | undefined {
  if (!/^[A-Za-z]+$/.test(columnName)) return undefined;
  let index = 0;
  for (const character of columnName.toUpperCase()) {
    index = index * 26 + character.charCodeAt(0) - 64;
    if (!Number.isSafeInteger(index)) return undefined;
  }
  return index - 1;
}

function columnIndexToName(columnIndex: number): string {
  let remaining = columnIndex + 1;
  let name = '';
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return name;
}

function unquoteSheetId(sheetId: string): string {
  const trimmed = sheetId.trim();
  if (trimmed.length >= 2 && trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replaceAll("''", "'");
  }
  if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replaceAll('""', '"');
  }
  return trimmed;
}

function quoteSheetId(sheetId: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetId) ? sheetId : `'${sheetId.replaceAll("'", "''")}'`;
}
