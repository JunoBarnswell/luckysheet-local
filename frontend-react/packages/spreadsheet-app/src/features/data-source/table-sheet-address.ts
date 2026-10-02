import type { CellData, WorkbookModel, WorksheetModel, TableSheetDefinition, WorkbookTableModel } from '@react-sheets/core-model';
import { compareWorkbookValues, type FormulaEngine, type FormulaValue } from '@react-sheets/formula-engine';

export class TableSheetAddressError extends Error {
  readonly code: 'TABLE_SHEET_ADDRESS_INVALID' | 'UNSUPPORTED_FEATURE';
  constructor(code: TableSheetAddressError['code'], message: string) {
    super(`${code}: ${message}`);
    this.code = code;
  }
}

export interface TableSheetCellAddress {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  readonly header?: CellData;
}

export interface WorkbookViewCell {
  readonly owner: WorksheetModel;
  readonly row: number;
  readonly column: number;
  readonly cell?: CellData;
  readonly writable: boolean;
}

/** Resolves value, formula, protection owner and edit address together. */
export function resolveWorkbookViewCell(
  workbook: WorkbookModel, sheet: WorksheetModel, formula: FormulaEngine,
  row: number, column: number,
  readCell: (sheet: WorksheetModel, row: number, column: number) => CellData | undefined,
): WorkbookViewCell {
  if (sheet.kind === 'table-sheet') {
    const address = resolveTableSheetCellAddress(workbook, sheet, formula, row, column);
    if (!address) return { owner: sheet, row, column, writable: false };
    const owner = workbook.getSheet(address.sheetId);
    return { owner, row: address.row, column: address.column, cell: address.header ?? readCell(owner, address.row, address.column), writable: !address.header };
  }
  if (sheet.kind === 'gantt-sheet' && row > 0 && sheet.ganttSheet) {
    const table = workbook.dataModel.tables.get(sheet.ganttSheet.viewId);
    const range = table?.sourceRange;
    const field = table?.fields[column];
    if (range && field && range.startRow + row <= range.endRow) {
      const owner = workbook.getSheet(range.sheetId);
      const sourceRow = range.startRow + row;
      const sourceColumn = range.startColumn + field.ordinal;
      return { owner, row: sourceRow, column: sourceColumn, cell: readCell(owner, sourceRow, sourceColumn), writable: true };
    }
  }
  return { owner: sheet, row, column, cell: readCell(sheet, row, column), writable: true };
}

interface ProjectionOrder {
  readonly definition: TableSheetDefinition;
  readonly table: WorkbookTableModel;
  readonly source: WorksheetModel;
  readonly inputRevision: number;
  readonly calculationGeneration: number;
  readonly rows: readonly number[];
}

/** Derived row positions only. Values remain in the canonical source. */
const orders = new WeakMap<WorksheetModel, ProjectionOrder>();

/** One address owner for TableSheet display, selection input and commit. */
export function resolveTableSheetCellAddress(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  formula: FormulaEngine,
  row: number,
  column: number,
): TableSheetCellAddress | undefined {
  if (sheet.kind !== 'table-sheet') return undefined;
  const definition = sheet.tableSheet;
  if (!definition) throw new TableSheetAddressError('TABLE_SHEET_ADDRESS_INVALID', `Sheet ${sheet.id} has no view definition`);
  const table = workbook.dataModel.tables.get(definition.viewId);
  if (!table) throw new TableSheetAddressError('TABLE_SHEET_ADDRESS_INVALID', `Table ${definition.viewId} is unavailable`);
  const visible = definition.columns[column];
  if (!visible) return undefined;
  const field = table.fields.find((field) => field.id === visible.fieldId);
  if (!field) throw new TableSheetAddressError('TABLE_SHEET_ADDRESS_INVALID', `Field ${visible.fieldId} is unavailable`);
  if (visible.formula !== undefined || visible.type === 'formula' || visible.type === 'lookup') {
    throw new TableSheetAddressError('UNSUPPORTED_FEATURE', `Field ${field.id} requires a canonical record calculation owner`);
  }
  if (row === 0) return { sheetId: sheet.id, row, column, header: { value: visible.caption } };
  const range = table.sourceRange;
  if (!range) throw new TableSheetAddressError('UNSUPPORTED_FEATURE', `Table ${table.id} requires a block-backed record address resolver`);
  const source = workbook.sheets.get(range.sheetId);
  if (!source) throw new TableSheetAddressError('TABLE_SHEET_ADDRESS_INVALID', `Source worksheet ${range.sheetId} is unavailable`);
  if (row < 1 || row > range.endRow - range.startRow) return undefined;
  const sort = [
    ...definition.grouping.map(({ fieldId }) => ({ fieldId, direction: 'asc' as const })),
    ...(definition.sortState ?? []),
  ];
  if (definition.grouping.some(({ collapsed }) => collapsed)) throw new TableSheetAddressError('UNSUPPORTED_FEATURE', 'Collapsed groups require group identity in the view contract');
  let sourceRow = range.startRow + row;
  if (sort.length) {
    let order = orders.get(sheet);
    if (!order || order.definition !== definition || order.table !== table || order.source !== source
      || order.inputRevision !== source.cells.revision || order.calculationGeneration !== formula.getCalculationGeneration()) {
      const rows = Array.from({ length: range.endRow - range.startRow }, (_, index) => range.startRow + index + 1);
      const sortFields = sort.map((entry) => {
        const field = table.fields.find((field) => field.id === entry.fieldId);
        if (!field) throw new TableSheetAddressError('TABLE_SHEET_ADDRESS_INVALID', `Sort field ${entry.fieldId} is unavailable`);
        return { column: range.startColumn + field.ordinal, direction: entry.direction };
      });
      const value = (sourceRow: number, sourceColumn: number): FormulaValue | undefined => {
        const cell = source.cells.get(sourceRow, sourceColumn);
        return formula.getSpillValueAt(source.id, sourceRow, sourceColumn) ?? (cell?.formula ? formula.getCellResult({ sheetId: source.id, row: sourceRow, column: sourceColumn })?.value : cell?.value);
      };
      const collation = formula.getCollationContext();
      rows.sort((left, right) => {
        for (const criterion of sortFields) {
          const difference = compareWorkbookValues(value(left, criterion.column), value(right, criterion.column), collation);
          if (difference) return criterion.direction === 'asc' ? difference : -difference;
        }
        return left - right;
      });
      order = { definition, table, source, inputRevision: source.cells.revision, calculationGeneration: formula.getCalculationGeneration(), rows };
      orders.set(sheet, order);
    }
    sourceRow = order.rows[row - 1]!;
  }
  return { sheetId: source.id, row: sourceRow, column: range.startColumn + field.ordinal };
}
