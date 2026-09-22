import type { CellAddress } from './ast';
import type { TableReferenceSpecifier } from './ast';
import type { RangeDependency } from './range-index';
import { normalizeRange } from './range-index';
import { createFormulaError, isFormulaError, type FormulaError, type FormulaValue } from './values';

export interface SheetTableColumnRef {
  readonly id: string;
  readonly name: string;
}

export interface SheetTableRef {
  readonly id: string;
  readonly sheetId: string;
  readonly name: string;
  readonly range: {
    readonly sheetId: string;
    readonly startRow: number;
    readonly endRow: number;
    readonly startColumn: number;
    readonly endColumn: number;
  };
  readonly hasHeaderRow: boolean;
  readonly hasTotalRow: boolean;
  readonly columns: readonly SheetTableColumnRef[];
}

export interface SheetTableReferenceRequest {
  readonly specifier?: TableReferenceSpecifier;
  readonly columnName?: string;
  readonly columnEndName?: string;
  readonly thisRow: boolean;
}

export function normalizeSheetTables(tables: readonly SheetTableRef[]): Map<string, SheetTableRef> {
  const index = new Map<string, SheetTableRef>();
  for (const table of tables) {
    index.set(table.name.trim().toUpperCase(), table);
  }
  return index;
}

function tableBounds(table: SheetTableRef) {
  const range = table.range;
  const bodyStartRow = range.startRow + (table.hasHeaderRow ? 1 : 0);
  const bodyEndRow = range.endRow - (table.hasTotalRow ? 1 : 0);
  const headerRow = table.hasHeaderRow ? range.startRow : null;
  const totalRow = table.hasTotalRow ? range.endRow : null;
  return { range, bodyStartRow, bodyEndRow, headerRow, totalRow };
}

function columnIndexOf(table: SheetTableRef, columnName: string): number | FormulaError {
  const columnIndex = table.columns.findIndex(
    (column) => column.name.trim().toUpperCase() === columnName.trim().toUpperCase(),
  );
  if (columnIndex < 0) return createFormulaError('#NAME?', `Unknown table column: ${columnName}`);
  return columnIndex;
}

function rangeForRows(
  table: SheetTableRef,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): RangeDependency {
  return normalizeRange(
    { sheetId: table.sheetId, row: startRow, column: startColumn },
    { sheetId: table.sheetId, row: endRow, column: endColumn },
  );
}

function resolveSpecifierRange(
  table: SheetTableRef,
  specifier: TableReferenceSpecifier,
  columnIndex?: number,
  endColumnIndex?: number,
): RangeDependency | CellAddress | FormulaError {
  const { range, bodyStartRow, bodyEndRow, headerRow, totalRow } = tableBounds(table);
  const startColumn = columnIndex === undefined ? range.startColumn : range.startColumn + columnIndex;
  const endColumn = endColumnIndex === undefined
    ? (columnIndex === undefined ? range.endColumn : range.startColumn + columnIndex)
    : range.startColumn + endColumnIndex;

  switch (specifier) {
    case 'all':
      return rangeForRows(table, range.startRow, range.endRow, startColumn, endColumn);
    case 'headers':
      if (headerRow === null) return createFormulaError('#REF!', 'Table has no header row');
      return columnIndex === undefined
        ? rangeForRows(table, headerRow, headerRow, startColumn, endColumn)
        : { sheetId: table.sheetId, row: headerRow, column: startColumn };
    case 'data':
      if (bodyEndRow < bodyStartRow) return createFormulaError('#REF!', 'Table has no data rows');
      return rangeForRows(table, bodyStartRow, bodyEndRow, startColumn, endColumn);
    case 'totals':
      if (totalRow === null) return createFormulaError('#REF!', 'Table has no total row');
      return columnIndex === undefined
        ? rangeForRows(table, totalRow, totalRow, startColumn, endColumn)
        : { sheetId: table.sheetId, row: totalRow, column: startColumn };
  }
}

export function resolveSheetTableReference(
  tableName: string,
  request: SheetTableReferenceRequest,
  currentCell: CellAddress,
  tables: ReadonlyMap<string, SheetTableRef>,
): RangeDependency | CellAddress | FormulaError {
  const table = tables.get(tableName.trim().toUpperCase());
  if (!table) return createFormulaError('#NAME?', `Unknown table: ${tableName}`);

  const columnIndex = request.columnName === undefined
    ? undefined
    : columnIndexOf(table, request.columnName);
  if (isFormulaError(columnIndex)) return columnIndex;
  const endColumnIndex = request.columnEndName === undefined
    ? columnIndex
    : columnIndexOf(table, request.columnEndName);
  if (isFormulaError(endColumnIndex)) return endColumnIndex;
  if (columnIndex !== undefined && endColumnIndex !== undefined && endColumnIndex < columnIndex) {
    return createFormulaError('#REF!', 'Structured table column range is reversed');
  }

  if (request.specifier) {
    return resolveSpecifierRange(table, request.specifier, columnIndex, endColumnIndex);
  }

  if (columnIndex === undefined) {
    return createFormulaError('#NAME?', 'Table reference requires a column or specifier');
  }

  const column = table.range.startColumn + columnIndex;
  const endColumn = table.range.startColumn + (endColumnIndex ?? columnIndex);
  const { bodyStartRow, bodyEndRow } = tableBounds(table);

  if (request.thisRow) {
    if (currentCell.sheetId !== table.sheetId) {
      return createFormulaError('#REF!', 'Structured table row reference requires the same worksheet');
    }
    if (currentCell.row < bodyStartRow || currentCell.row > bodyEndRow) {
      return createFormulaError('#REF!', 'Current row is outside the table data body');
    }
    return endColumn === column
      ? { sheetId: table.sheetId, row: currentCell.row, column }
      : rangeForRows(table, currentCell.row, currentCell.row, column, endColumn);
  }

  if (bodyEndRow < bodyStartRow) {
    return createFormulaError('#REF!', 'Table has no data rows');
  }

  return normalizeRange(
    { sheetId: table.sheetId, row: bodyStartRow, column },
    { sheetId: table.sheetId, row: bodyEndRow, column: endColumn },
  );
}

// Backward-compatible helper for legacy call sites/tests.
export function resolveSheetTableColumnReference(
  tableName: string,
  columnName: string,
  thisRow: boolean,
  currentCell: CellAddress,
  tables: ReadonlyMap<string, SheetTableRef>,
): RangeDependency | CellAddress | FormulaError {
  return resolveSheetTableReference(tableName, { columnName, thisRow }, currentCell, tables);
}
