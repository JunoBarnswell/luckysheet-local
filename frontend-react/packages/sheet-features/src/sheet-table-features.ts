import type {
  CellData,
  CellStyle,
  FilterModel,
  RangeRef,
  SheetTableModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { normalizeRangeRef } from './data-features';

type SheetTableColumn = SheetTableModel['columns'][number];
export type TotalsFunction = NonNullable<SheetTableColumn['totalsFunction']>;

const TABLE_HEADER_STYLE: Partial<CellStyle> = {
  background: '#4472C4',
  textColor: '#FFFFFF',
  bold: true,
};

const TABLE_TOTAL_STYLE: Partial<CellStyle> = {
  background: '#D9E1F2',
  bold: true,
};

const TABLE_BAND_LIGHT = '#FFFFFF';
const TABLE_BAND_DARK = '#D9E1F2';

export function isPointInRange(range: RangeRef, row: number, column: number): boolean {
  const normalized = normalizeRangeRef(range);
  return row >= normalized.startRow
    && row <= normalized.endRow
    && column >= normalized.startColumn
    && column <= normalized.endColumn;
}

export function findSheetTableAt(sheet: WorksheetModel, row: number, column: number): SheetTableModel | undefined {
  return sheet.sheetTables.find((table) => table.sheetId === sheet.id && isPointInRange(table.range, row, column));
}

export function tableBodyBounds(table: SheetTableModel): { startRow: number; endRow: number } {
  const range = normalizeRangeRef(table.range);
  return {
    startRow: range.startRow + (table.hasHeaderRow ? 1 : 0),
    endRow: range.endRow - (table.hasTotalRow ? 1 : 0),
  };
}

export function computeBandedCellStyle(sheet: WorksheetModel, row: number, column: number): Partial<CellStyle> | undefined {
  const rule = sheet.bandedRule;
  if (!rule || !isPointInRange(rule.range, row, column)) return undefined;
  const bandIndex = row - rule.range.startRow;
  return { background: bandIndex % 2 === 0 ? rule.firstColor : rule.secondColor };
}

export function computeSheetTableCellStyle(table: SheetTableModel, row: number, column: number): Partial<CellStyle> | undefined {
  const range = normalizeRangeRef(table.range);
  if (!isPointInRange(range, row, column)) return undefined;

  if (table.hasHeaderRow && row === range.startRow) return TABLE_HEADER_STYLE;
  if (table.hasTotalRow && row === range.endRow) return TABLE_TOTAL_STYLE;

  const { startRow, endRow } = tableBodyBounds(table);
  if (row < startRow || row > endRow) return { background: TABLE_BAND_LIGHT };

  if (table.showBandedRows) {
    const bandIndex = row - startRow;
    return { background: bandIndex % 2 === 0 ? TABLE_BAND_LIGHT : TABLE_BAND_DARK };
  }

  if (table.showBandedColumns) {
    const bandIndex = column - range.startColumn;
    return { background: bandIndex % 2 === 0 ? TABLE_BAND_LIGHT : TABLE_BAND_DARK };
  }

  return { background: TABLE_BAND_LIGHT };
}

export function mergePresentationStyles(...styles: Array<Partial<CellStyle> | undefined>): Partial<CellStyle> | undefined {
  const merged: Partial<CellStyle> = {};
  for (const style of styles) {
    if (!style) continue;
    Object.assign(merged, style);
  }
  return Object.keys(merged).length > 0 ? merged : undefined;
}

export function createFilterModelForTable(table: SheetTableModel): FilterModel {
  return {
    sheetId: table.sheetId,
    range: normalizeRangeRef(table.range),
    criteria: {},
  };
}

export function tableFilterColumns(table: SheetTableModel): number[] {
  if (!table.showFilterButton) return [];
  const range = normalizeRangeRef(table.range);
  const columns: number[] = [];
  for (let column = range.startColumn; column <= range.endColumn; column++) columns.push(column);
  return columns;
}

export function resolveActiveFilterColumns(sheet: WorksheetModel): number[] {
  if (!sheet.filter) return [];
  const fromCriteria = Object.keys(sheet.filter.criteria)
    .map(Number)
    .filter((column) => Number.isFinite(column));
  if (fromCriteria.length > 0) return fromCriteria.sort((left, right) => left - right);

  const range = normalizeRangeRef(sheet.filter.range);
  const columns: number[] = [];
  for (let column = range.startColumn; column <= range.endColumn; column++) columns.push(column);
  return columns;
}

export interface FilterButtonCell {
  row: number;
  column: number;
}

function rangesEqual(left: RangeRef, right: RangeRef): boolean {
  const a = normalizeRangeRef(left);
  const b = normalizeRangeRef(right);
  return a.startRow === b.startRow
    && a.endRow === b.endRow
    && a.startColumn === b.startColumn
    && a.endColumn === b.endColumn;
}

export function findSheetTableForFilter(sheet: WorksheetModel): SheetTableModel | undefined {
  if (!sheet.filter) return undefined;
  const filterRange = normalizeRangeRef(sheet.filter.range);
  return sheet.sheetTables.find(
    (table) => table.sheetId === sheet.id
      && table.showFilterButton
      && rangesEqual(table.range, filterRange),
  );
}

/** 表筛选漏斗绘制在表头行；普通区域筛选仍走列头字母条 */
export function resolveFilterButtonCells(sheet: WorksheetModel): FilterButtonCell[] {
  const table = findSheetTableForFilter(sheet);
  if (!table) return [];
  const range = normalizeRangeRef(table.range);
  const headerRow = table.hasHeaderRow ? range.startRow : range.startRow;
  const buttons: FilterButtonCell[] = [];
  for (let column = range.startColumn; column <= range.endColumn; column++) {
    buttons.push({ row: headerRow, column });
  }
  return buttons;
}

export function subtotalCodeForTotalsFunction(totalsFunction: SheetTableColumn['totalsFunction']): number | null {
  switch (totalsFunction ?? 'sum') {
    case 'sum':
      return 109;
    case 'average':
      return 101;
    case 'count':
      return 103;
    case 'min':
      return 105;
    case 'max':
      return 104;
    case 'none':
    default:
      return null;
  }
}

export function buildTotalRowFormula(
  tableName: string,
  columnName: string,
  totalsFunction?: SheetTableColumn['totalsFunction'],
): string | null {
  const code = subtotalCodeForTotalsFunction(totalsFunction);
  if (code === null) return null;
  return `=SUBTOTAL(${code},${tableName}[${columnName}])`;
}

export function defaultTotalsFunction(columnIndex: number): TotalsFunction {
  return columnIndex === 0 ? 'none' : 'sum';
}

export interface TotalRowTogglePlan {
  nextTable: SheetTableModel;
  totalRow: number;
  startColumn: number;
  endColumn: number;
  values: CellData[][];
  clearTotalRow: boolean;
}

export function planTotalRowToggle(table: SheetTableModel, enabled: boolean): TotalRowTogglePlan {
  const range = normalizeRangeRef(table.range);
  if (enabled) {
    const expandedRange = { ...range, endRow: range.endRow + 1 };
    const totalRow = expandedRange.endRow;
    const rowValues: CellData[] = table.columns.map((column, columnIndex) => {
      const totalsFunction = column.totalsFunction ?? defaultTotalsFunction(columnIndex);
      if (columnIndex === 0 && totalsFunction === 'none') return { value: 'Total' };
      const formula = buildTotalRowFormula(table.name, column.name, totalsFunction);
      return formula ? { value: null, formula } : { value: null };
    });
    return {
      nextTable: { ...table, range: expandedRange, hasTotalRow: true },
      totalRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
      values: [rowValues],
      clearTotalRow: false,
    };
  }

  const totalRow = range.endRow;
  const shrunkRange = { ...range, endRow: range.endRow - 1 };
  return {
    nextTable: { ...table, range: shrunkRange, hasTotalRow: false },
    totalRow,
    startColumn: range.startColumn,
    endColumn: range.endColumn,
    values: [],
    clearTotalRow: true,
  };
}

export function snapshotTotalRowCells(
  sheet: WorksheetModel,
  totalRow: number,
  startColumn: number,
  endColumn: number,
): Array<{ row: number; column: number; previous?: CellData }> {
  const snapshots: Array<{ row: number; column: number; previous?: CellData }> = [];
  for (let column = startColumn; column <= endColumn; column++) {
    snapshots.push({ row: totalRow, column, previous: sheet.cells.get(totalRow, column) });
  }
  return snapshots;
}
