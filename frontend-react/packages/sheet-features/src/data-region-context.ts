import type { RangeRef, SheetTableModel, WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import { resolveFilterOwner, type FilterOwner } from './sheet-table-features';

export type DataRegionOwner =
  | { kind: 'worksheet' }
  | { kind: 'sheet-table'; tableId: string };

export type HeaderResolution =
  | { kind: 'present'; row: number }
  | { kind: 'absent' };

export interface DataRegionContext {
  schema: 'DataRegionContext';
  version: 1;
  sheetId: string;
  selection: RangeRef;
  currentRegion: RangeRef;
  range: RangeRef;
  usedRange: RangeRef;
  owner: DataRegionOwner;
  header: HeaderResolution;
  activeColumn: number;
  visibleRows: readonly number[];
  searchScope: 'selection' | 'current-region' | 'sheet' | 'workbook';
}

export interface DataRegionResolveInput {
  selection: RangeRef;
  activeRow: number;
  activeColumn: number;
  searchScope?: DataRegionContext['searchScope'];
}

function cloneRange(range: RangeRef): RangeRef {
  return structuredClone(range);
}

function contains(range: RangeRef, row: number, column: number): boolean {
  return range.startRow <= row && row <= range.endRow && range.startColumn <= column && column <= range.endColumn;
}

function cellHasContent(sheet: WorksheetModel, row: number, column: number): boolean {
  const cell = sheet.cells.get(row, column);
  return Boolean(cell && (cell.value !== null && cell.value !== undefined || cell.formula !== undefined || cell.formulaValue !== undefined));
}

function usedRange(sheet: WorksheetModel): RangeRef {
  let startRow = sheet.rowCount - 1;
  let endRow = 0;
  let startColumn = sheet.columnCount - 1;
  let endColumn = 0;
  let occupied = false;
  sheet.cells.forEach((_cell, row, column) => {
    occupied = true;
    startRow = Math.min(startRow, row);
    endRow = Math.max(endRow, row);
    startColumn = Math.min(startColumn, column);
    endColumn = Math.max(endColumn, column);
  });
  for (const { row, column } of sheet.review.noteEntries()) {
    occupied = true;
    startRow = Math.min(startRow, row);
    endRow = Math.max(endRow, row);
    startColumn = Math.min(startColumn, column);
    endColumn = Math.max(endColumn, column);
  }
  for (const thread of sheet.review.threadEntries()) {
    occupied = true;
    startRow = Math.min(startRow, thread.row);
    endRow = Math.max(endRow, thread.row);
    startColumn = Math.min(startColumn, thread.column);
    endColumn = Math.max(endColumn, thread.column);
  }
  if (!occupied) return { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

function currentRegion(sheet: WorksheetModel, input: DataRegionResolveInput): RangeRef {
  const selection = cloneRange(input.selection);
  if (selection.startRow !== selection.endRow || selection.startColumn !== selection.endColumn) return selection;
  const table = sheet.sheetTables.find((candidate) => contains(candidate.range, input.activeRow, input.activeColumn));
  if (table) return cloneRange(table.range);
  const dataRegion = sheet.dataRegions.find((candidate) => contains(candidate.range, input.activeRow, input.activeColumn));
  if (dataRegion) return cloneRange(dataRegion.range);
  if (!cellHasContent(sheet, input.activeRow, input.activeColumn)) {
    return { sheetId: sheet.id, startRow: input.activeRow, endRow: input.activeRow, startColumn: input.activeColumn, endColumn: input.activeColumn };
  }
  let startRow = input.activeRow;
  let endRow = input.activeRow;
  let startColumn = input.activeColumn;
  let endColumn = input.activeColumn;
  const rowHasContent = (row: number) => {
    for (let column = startColumn; column <= endColumn; column += 1) if (cellHasContent(sheet, row, column)) return true;
    return false;
  };
  const columnHasContent = (column: number) => {
    for (let row = startRow; row <= endRow; row += 1) if (cellHasContent(sheet, row, column)) return true;
    return false;
  };
  let grew = true;
  while (grew) {
    grew = false;
    while (startRow > 0 && rowHasContent(startRow - 1)) { startRow -= 1; grew = true; }
    while (endRow + 1 < sheet.rowCount && rowHasContent(endRow + 1)) { endRow += 1; grew = true; }
    while (startColumn > 0 && columnHasContent(startColumn - 1)) { startColumn -= 1; grew = true; }
    while (endColumn + 1 < sheet.columnCount && columnHasContent(endColumn + 1)) { endColumn += 1; grew = true; }
  }
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

function tableAtRange(sheet: WorksheetModel, range: RangeRef): SheetTableModel | undefined {
  return sheet.sheetTables.find((table) => table.sheetId === sheet.id
    && table.range.startRow === range.startRow && table.range.endRow === range.endRow
    && table.range.startColumn === range.startColumn && table.range.endColumn === range.endColumn);
}

function resolveOwner(sheet: WorksheetModel, range: RangeRef, activeColumn: number): DataRegionOwner {
  const table = tableAtRange(sheet, range) ?? sheet.sheetTables.find((candidate) => contains(candidate.range, range.startRow, range.startColumn));
  if (table) return { kind: 'sheet-table', tableId: table.id };
  const filterOwner: FilterOwner | undefined = resolveFilterOwner(sheet, activeColumn);
  if (filterOwner?.kind === 'table') return { kind: 'sheet-table', tableId: filterOwner.tableId };
  return { kind: 'worksheet' };
}

function resolveHeader(sheet: WorksheetModel, range: RangeRef, owner: DataRegionOwner): HeaderResolution {
  if (owner.kind === 'sheet-table') {
    const table = sheet.sheetTables.find((candidate) => candidate.id === owner.tableId);
    if (table?.hasHeaderRow) return { kind: 'present', row: table.range.startRow };
    return { kind: 'absent' };
  }
  const region = sheet.dataRegions.find((candidate) => candidate.range.startRow === range.startRow
    && candidate.range.endRow === range.endRow && candidate.range.startColumn === range.startColumn
    && candidate.range.endColumn === range.endColumn);
  if (region && region.headerRow >= range.startRow && region.headerRow <= range.endRow) return { kind: 'present', row: region.headerRow };
  let populated = 0;
  let textHeaders = 0;
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const cell = sheet.cells.get(range.startRow, column);
    const value = cell?.formulaValue ?? cell?.value;
    if (value === null || value === undefined || value === '') continue;
    populated += 1;
    if (typeof value === 'string') textHeaders += 1;
  }
  return populated > 0 && populated === textHeaders
    ? { kind: 'present', row: range.startRow }
    : { kind: 'absent' };
}

export function resolveDataRegionContext(workbook: WorkbookModel, input: DataRegionResolveInput): DataRegionContext {
  const sheet = workbook.getSheet(input.selection.sheetId);
  const selection = cloneRange(input.selection);
  const region = currentRegion(sheet, input);
  const owner = resolveOwner(sheet, region, input.activeColumn);
  const header = resolveHeader(sheet, region, owner);
  const visibleRows: number[] = [];
  for (let row = region.startRow; row <= region.endRow; row += 1) {
    if (!sheet.hiddenRows.has(row)) visibleRows.push(row);
  }
  return {
    schema: 'DataRegionContext',
    version: 1,
    sheetId: sheet.id,
    selection,
    currentRegion: cloneRange(region),
    range: cloneRange(region),
    usedRange: usedRange(sheet),
    owner,
    header,
    activeColumn: input.activeColumn,
    visibleRows,
    searchScope: input.searchScope ?? 'current-region',
  };
}

export function isDataRegionContext(value: unknown): value is DataRegionContext {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const context = value as Partial<DataRegionContext>;
  return context.schema === 'DataRegionContext' && context.version === 1
    && typeof context.sheetId === 'string' && typeof context.activeColumn === 'number'
    && Boolean(context.selection && context.currentRegion && context.range && context.usedRange && context.owner && context.header)
    && Array.isArray(context.visibleRows);
}

export function assertDataRegionContextMatches(expected: DataRegionContext, actual: DataRegionContext): void {
  if (!isDataRegionContext(expected)) throw new Error('Invalid DataRegionContext payload');
  if (JSON.stringify(expected.range) !== JSON.stringify(actual.range)) throw new Error('DATA_REGION_CONTEXT_MISMATCH: range changed');
  if (JSON.stringify(expected.owner) !== JSON.stringify(actual.owner)) throw new Error('DATA_REGION_CONTEXT_MISMATCH: owner changed');
  if (JSON.stringify(expected.header) !== JSON.stringify(actual.header)) throw new Error('DATA_REGION_CONTEXT_MISMATCH: header semantics changed');
}

export function filterOwnerFromDataRegionContext(context: DataRegionContext): FilterOwner {
  return context.owner.kind === 'sheet-table' ? { kind: 'table', tableId: context.owner.tableId } : { kind: 'worksheet' };
}
