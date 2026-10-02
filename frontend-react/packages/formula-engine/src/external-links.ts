import type { ExternalReferenceNode, CellAddress } from './ast';
import type { FormulaSheetIdentity } from './sheet-reference';
import { createFormulaError, type FormulaValue, type FormulaError, isFormulaError } from './values';
import type { RangeDependency } from './range-index';

export type ExternalLinkState = 'connected' | 'refreshing' | 'stale' | 'denied' | 'unavailable' | 'broken';
/** Transient subject-bound calculation data. Never persisted with an authored workbook. */
export interface ExternalCalculationLink {
  readonly id: string;
  readonly token: string;
  readonly sourceUnitId: string;
  readonly subject: string;
  readonly sourceRevision: number;
  readonly accessRevision: number;
  readonly state: ExternalLinkState;
  readonly error?: { code: string; message: string };
  readonly sheets: readonly (FormulaSheetIdentity & { rowCount: number; columnCount: number })[];
  readonly blockedRanges?: readonly { sheetId: string; startRow: number; endRow: number; startColumn: number; endColumn: number }[];
  readonly cells: readonly { address: CellAddress; value: FormulaValue }[];
}

export function assertCalculationBlockedRanges(value: unknown): asserts value is NonNullable<ExternalCalculationLink['blockedRanges']> {
  if (!Array.isArray(value) || value.some(range => !range || typeof range.sheetId !== 'string' || !range.sheetId || !['startRow', 'endRow', 'startColumn', 'endColumn'].every(key => Number.isSafeInteger(range[key]) && range[key] >= 0) || range.endRow < range.startRow || range.endColumn < range.startColumn || range.endRow > 1048575 || range.endColumn > 16383)) throw new Error('CALCULATION_BLOCKED_RANGE_INVALID');
}

function validExternalValue(value: unknown, depth = 0): boolean {
  if (value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value)) return true;
  if (isFormulaError(value)) return typeof value.code === 'string' && typeof value.message === 'string';
  return depth < 4 && Array.isArray(value) && value.length <= 100000 && value.every(row => Array.isArray(row) && row.length <= 16384 && row.every(cell => validExternalValue(cell, depth + 1)));
}

export function assertExternalCalculationLink(value: ExternalCalculationLink): void {
  if (!value || !value.id || !value.token || !value.sourceUnitId || !value.subject
    || !Number.isSafeInteger(value.sourceRevision) || value.sourceRevision < 0
    || !Number.isSafeInteger(value.accessRevision) || value.accessRevision < 0
    || !['connected', 'refreshing', 'stale', 'denied', 'unavailable', 'broken'].includes(value.state)
    || !Array.isArray(value.sheets) || !Array.isArray(value.cells) || value.cells.length > 100000) throw new Error('EXTERNAL_LINK_CONTEXT_INVALID');
  const sheets = new Map(value.sheets.map(sheet => [sheet.id, sheet]));
  if (sheets.size !== value.sheets.length || value.sheets.some(sheet => !sheet.id || !sheet.name
    || !Number.isSafeInteger(sheet.rowCount) || sheet.rowCount < 1 || !Number.isSafeInteger(sheet.columnCount) || sheet.columnCount < 1)) throw new Error('EXTERNAL_LINK_SHEET_INVALID');
  if (value.blockedRanges !== undefined) assertCalculationBlockedRanges(value.blockedRanges);
  const cells = new Set<string>();
  for (const cell of value.cells) {
    const sheet = sheets.get(cell.address.sheetId);
    const key = JSON.stringify(cell.address);
    if (!sheet || !Number.isSafeInteger(cell.address.row) || cell.address.row < 0 || cell.address.row >= sheet.rowCount
      || !Number.isSafeInteger(cell.address.column) || cell.address.column < 0 || cell.address.column >= sheet.columnCount
      || cells.has(key) || !validExternalValue(cell.value)) throw new Error('EXTERNAL_LINK_CELL_INVALID');
    cells.add(key);
  }
  if (!['connected', 'stale', 'refreshing'].includes(value.state) && value.cells.length) throw new Error('EXTERNAL_LINK_REVOKED_CACHE');
}

export function externalSheetKey(linkId: string, sheetId: string): string {
  return `external:${encodeURIComponent(linkId)}:${encodeURIComponent(sheetId)}`;
}

export function externalReferenceRange(node: ExternalReferenceNode, link?: ExternalCalculationLink): RangeDependency | FormulaError {
  if (!link || link.state === 'broken') return createFormulaError('#REF!', link?.error?.message ?? 'External link is not bound');
  if (link.state === 'denied') return createFormulaError('#BLOCKED!', link.error?.message ?? 'External source access denied');
  if (link.state === 'unavailable' || link.state === 'refreshing' && !link.cells.length) return createFormulaError('#N/A', link.error?.message ?? 'External source is unavailable');
  const sheet = link.sheets.find(sheet => sheet.id === node.qualifier.sheetId || sheet.name.toUpperCase() === node.qualifier.sheetId?.toUpperCase());
  if (!sheet) return createFormulaError('#REF!', 'External source worksheet is missing');
  const sheetId = externalSheetKey(link.id, sheet.id);
  const ref = node.reference;
  let start: CellAddress, end: CellAddress;
  switch (ref.type) {
    case 'cell-reference': start = { sheetId, row: ref.reference.row, column: ref.reference.column }; end = start; break;
    case 'range-reference': start = { sheetId, row: Math.min(ref.start.reference.row, ref.end.reference.row), column: Math.min(ref.start.reference.column, ref.end.reference.column) }; end = { sheetId, row: Math.max(ref.start.reference.row, ref.end.reference.row), column: Math.max(ref.start.reference.column, ref.end.reference.column) }; break;
    case 'whole-column-reference': start = { sheetId, row: 0, column: ref.startColumn }; end = { sheetId, row: sheet.rowCount - 1, column: ref.endColumn }; break;
    case 'whole-row-reference': start = { sheetId, row: ref.startRow, column: 0 }; end = { sheetId, row: ref.endRow, column: sheet.columnCount - 1 }; break;
    default: return createFormulaError('#REF!', 'External structured table binding is unavailable');
  }
  if (end.row >= sheet.rowCount || end.column >= sheet.columnCount) return createFormulaError('#REF!', 'External reference exceeds source extent');
  if (link.blockedRanges?.some(range => range.sheetId === sheet.id && range.startRow <= end.row && start.row <= range.endRow && range.startColumn <= end.column && start.column <= range.endColumn)) return createFormulaError('#BLOCKED!', 'External range contains hidden inputs');
  return { kind: 'range', start, end };
}
