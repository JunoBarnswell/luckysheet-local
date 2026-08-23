import type { CellAddress } from './ast';
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

export function normalizeSheetTables(tables: readonly SheetTableRef[]): Map<string, SheetTableRef> {
  const index = new Map<string, SheetTableRef>();
  for (const table of tables) {
    index.set(table.name.trim().toUpperCase(), table);
  }
  return index;
}

export function resolveSheetTableReference(
  tableName: string,
  columnName: string,
  thisRow: boolean,
  currentCell: CellAddress,
  tables: ReadonlyMap<string, SheetTableRef>,
): RangeDependency | CellAddress | FormulaError {
  const table = tables.get(tableName.trim().toUpperCase());
  if (!table) return createFormulaError('#NAME?', `Unknown table: ${tableName}`);

  const columnIndex = table.columns.findIndex(
    (column) => column.name.trim().toUpperCase() === columnName.trim().toUpperCase(),
  );
  if (columnIndex < 0) return createFormulaError('#NAME?', `Unknown table column: ${columnName}`);

  const column = table.range.startColumn + columnIndex;
  const bodyStartRow = table.range.startRow + (table.hasHeaderRow ? 1 : 0);
  const bodyEndRow = table.range.endRow - (table.hasTotalRow ? 1 : 0);

  if (thisRow) {
    if (currentCell.sheetId !== table.sheetId) {
      return createFormulaError('#REF!', 'Structured table row reference requires the same worksheet');
    }
    if (currentCell.row < bodyStartRow || currentCell.row > bodyEndRow) {
      return createFormulaError('#REF!', 'Current row is outside the table data body');
    }
    return { sheetId: table.sheetId, row: currentCell.row, column };
  }

  if (bodyEndRow < bodyStartRow) {
    return createFormulaError('#REF!', 'Table has no data rows');
  }

  return normalizeRange(
    { sheetId: table.sheetId, row: bodyStartRow, column },
    { sheetId: table.sheetId, row: bodyEndRow, column },
  );
}
