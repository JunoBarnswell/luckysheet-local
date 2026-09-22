import type {
  CellData,
  CellStyle,
  AutoFilterModel,
  RangeRef,
  SheetTableModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { normalizeRangeRef } from './data-features';

type SheetTableColumn = SheetTableModel['columns'][number];
export type TotalsFunction = NonNullable<SheetTableColumn['totalsFunction']>;

export interface SheetTableCreationRequest {
  sheetId: string;
  range: RangeRef;
  name: string;
  hasHeaderRow: boolean;
  styleName?: string;
  existingNames?: Iterable<string>;
  nextId: (prefix: string) => string;
  readCell?: (row: number, column: number) => unknown;
}

export interface SheetTableCreationPlan {
  table: SheetTableModel;
}

export interface SheetTableAutoExpansionPlan {
  previous: SheetTableModel;
  next: SheetTableModel;
}

function canonicalColumnName(raw: unknown, fallback: string): string {
  const value = typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean'
    ? String(raw).trim()
    : '';
  return value || fallback;
}

function uniqueColumnName(raw: string, used: Set<string>): string {
  let candidate = raw;
  let suffix = 2;
  while (used.has(candidate.toLocaleLowerCase())) candidate = `${raw}${suffix++}`;
  used.add(candidate.toLocaleLowerCase());
  return candidate;
}

/**
 * Build the complete table descriptor before the first workbook mutation.
 * The first row is only a header source when hasHeaderRow is true; otherwise
 * it remains ordinary table body data and receives generated ColumnN names.
 */
export function planSheetTableCreation(request: SheetTableCreationRequest, sheet?: WorksheetModel): SheetTableCreationPlan {
  const range = normalizeRangeRef(request.range);
  if (range.sheetId !== request.sheetId) throw new Error('Create Table range must target the active worksheet');
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new Error('Create Table range is invalid');
  }
  if (sheet && (range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount)) {
    throw new Error('Create Table range exceeds worksheet bounds');
  }
  const name = request.name.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) throw new Error(`Invalid Sheet Table name: ${request.name}`);
  const usedNames = new Set([...request.existingNames ?? []].map((entry) => entry.trim().toLocaleLowerCase()));
  if (usedNames.has(name.toLocaleLowerCase())) throw new Error(`Sheet Table already exists: ${name}`);
  const usedColumns = new Set<string>();
  const columns: SheetTableColumn[] = [];
  for (let offset = 0; offset <= range.endColumn - range.startColumn; offset += 1) {
    const source = request.hasHeaderRow && request.readCell
      ? request.readCell(range.startRow, range.startColumn + offset)
      : undefined;
    const raw = canonicalColumnName(source, `Column${offset + 1}`);
    columns.push({ id: request.nextId('col'), name: uniqueColumnName(raw, usedColumns), totalsFunction: defaultTotalsFunction(offset) });
  }
  return {
    table: validateSheetTableModel({
      id: request.nextId('sheet-table'),
      sheetId: request.sheetId,
      name,
      range,
      hasHeaderRow: request.hasHeaderRow,
      hasTotalRow: false,
      showBandedRows: true,
      showBandedColumns: false,
      showFirstColumn: false,
      showLastColumn: false,
      showFilterButton: request.hasHeaderRow,
      autoExpand: 'both',
      columns,
      ...(request.styleName?.trim() ? { styleName: request.styleName.trim() } : {}),
    }, sheet),
  };
}

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

export function validateSheetTableModel(table: SheetTableModel, sheet?: WorksheetModel): SheetTableModel {
  const range = normalizeRangeRef(table.range);
  if (!table.id.trim() || !table.name.trim()) throw new Error('Sheet Table id and name are required');
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(table.name)) throw new Error(`Invalid Sheet Table name: ${table.name}`);
  if (range.sheetId !== table.sheetId) throw new Error('Sheet Table range must target its sheetId');
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new Error('Sheet Table range is invalid');
  }
  if (table.columns.length !== range.endColumn - range.startColumn + 1) {
    throw new Error('Sheet Table columns must match the range width');
  }
  if (!['none', 'rows', 'columns', 'both'].includes(table.autoExpand)) throw new Error('Sheet Table auto-expand mode is invalid');
  if (!table.hasHeaderRow && table.showFilterButton) throw new Error('A Sheet Table without a header row cannot show filter buttons');
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const column of table.columns) {
    if (!column.id.trim() || !column.name.trim()) throw new Error('Sheet Table columns require id and name');
    if (ids.has(column.id) || names.has(column.name.toLocaleLowerCase())) throw new Error('Sheet Table columns must be unique');
    ids.add(column.id);
    names.add(column.name.toLocaleLowerCase());
  }
  if (table.hasHeaderRow && table.hasTotalRow && range.endRow - range.startRow < 2) {
    throw new Error('Sheet Table with header and total rows requires at least one body row');
  }
  if (table.autoFilter) {
    validateFilterModelOwnership(table.autoFilter, table.sheetId, range, 'Table AutoFilter');
  }
  if (sheet && range.endRow >= sheet.rowCount) throw new Error('Sheet Table exceeds worksheet rows');
  const normalized = { ...structuredClone(table), range };
  if (normalized.showFilterButton && normalized.hasHeaderRow && !normalized.autoFilter) {
    normalized.autoFilter = createAutoFilterModelForTable(normalized);
  }
  return normalized;
}

/**
 * Plan automatic table growth for a contiguous write immediately below or to
 * the right of a table. The plan is metadata-only; callers apply it in the
 * same command transaction as the cell write.
 */
export function planSheetTableAutoExpansion(
  sheet: WorksheetModel,
  writeRange: RangeRef,
  nextId: (prefix: string) => string,
): SheetTableAutoExpansionPlan[] {
  const write = normalizeRangeRef(writeRange);
  if (write.sheetId !== sheet.id) throw new Error('Table auto-expansion write must target its worksheet');
  const plans: SheetTableAutoExpansionPlan[] = [];
  for (const table of sheet.sheetTables.filter((entry) => entry.sheetId === sheet.id)) {
    const source = normalizeRangeRef(table.range);
    const allowRows = table.autoExpand === 'rows' || table.autoExpand === 'both';
    const allowColumns = table.autoExpand === 'columns' || table.autoExpand === 'both';
    const expandsRows = allowRows
      && write.startColumn >= source.startColumn && write.endColumn <= source.endColumn
      && write.startRow <= source.endRow + 1 && write.endRow > source.endRow;
    const expandsColumns = allowColumns
      && write.startRow >= source.startRow && write.endRow <= source.endRow
      && write.startColumn <= source.endColumn + 1 && write.endColumn > source.endColumn;
    if (!expandsRows && !expandsColumns) continue;

    if (expandsRows && table.hasTotalRow) {
      throw new Error(`Cannot auto-expand Sheet Table ${table.name} with a total row; resize it explicitly`);
    }
    const next = structuredClone(table);
    if (expandsRows) next.range.endRow = Math.max(next.range.endRow, write.endRow);
    if (expandsColumns) {
      const oldEnd = next.range.endColumn;
      next.range.endColumn = Math.max(next.range.endColumn, write.endColumn);
      const usedNames = new Set(next.columns.map((column) => column.name.toLocaleLowerCase()));
      for (let column = oldEnd + 1; column <= next.range.endColumn; column += 1) {
        const base = `Column${next.columns.length + 1}`;
        const name = uniqueColumnName(base, usedNames);
        next.columns.push({ id: nextId('col'), name, totalsFunction: defaultTotalsFunction(next.columns.length) });
      }
    }
    if (next.range.endRow >= sheet.rowCount || next.range.endColumn >= sheet.columnCount) {
      throw new Error(`Sheet Table ${table.name} cannot expand beyond worksheet bounds`);
    }
    if (sheet.sheetTables.some((other) => other.id !== table.id && rangesOverlap(next.range, other.range))) {
      throw new Error(`Sheet Table ${table.name} cannot auto-expand into another table`);
    }
    const previousFilter = next.autoFilter;
    if (next.showFilterButton && next.hasHeaderRow) {
      const expandedFilter = createAutoFilterModelForTable(next);
      if (previousFilter) {
        for (const [key, column] of Object.entries(previousFilter.columns)) {
          const columnIndex = Number(key);
          if (Number.isInteger(columnIndex) && expandedFilter.columns[columnIndex]) {
            expandedFilter.columns[columnIndex] = { ...expandedFilter.columns[columnIndex], ...structuredClone(column) };
          }
        }
        if (previousFilter.sortState) expandedFilter.sortState = structuredClone(previousFilter.sortState);
        if (previousFilter.preservedXml !== undefined) expandedFilter.preservedXml = structuredClone(previousFilter.preservedXml);
      }
      next.autoFilter = expandedFilter;
    } else {
      next.autoFilter = undefined;
    }
    plans.push({ previous: structuredClone(table), next: validateSheetTableModel(next, sheet) });
  }
  return plans;
}

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

  if (table.showFirstColumn && column === range.startColumn) return { background: TABLE_BAND_DARK, bold: true };
  if (table.showLastColumn && column === range.endColumn) return { background: TABLE_BAND_DARK, bold: true };

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

export function createAutoFilterModelForTable(table: SheetTableModel): AutoFilterModel {
  const range = normalizeRangeRef(table.range);
  const columns: AutoFilterModel['columns'] = {};
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    columns[column] = { column, showButton: table.showFilterButton, hiddenButton: false };
  }
  return {
    sheetId: table.sheetId,
    range,
    columns,
  };
}

export type FilterOwner = { kind: 'worksheet' } | { kind: 'table'; tableId: string };

export interface ResolvedAutoFilter {
  owner: FilterOwner;
  autoFilter: AutoFilterModel;
}

export function resolveAutoFilters(sheet: WorksheetModel): ResolvedAutoFilter[] {
  const resolved: ResolvedAutoFilter[] = sheet.autoFilter
    ? [{ owner: { kind: 'worksheet' }, autoFilter: sheet.autoFilter }]
    : [];
  for (const table of sheet.sheetTables.filter((entry) => entry.sheetId === sheet.id && entry.autoFilter)) {
    const autoFilter = table.autoFilter!;
    validateFilterModelOwnership(autoFilter, sheet.id, table.range, 'Table AutoFilter');
    if (resolved.some((entry) => rangesOverlap(entry.autoFilter.range, autoFilter.range))) {
      throw new Error('AutoFilter ranges cannot overlap');
    }
    resolved.push({ owner: { kind: 'table', tableId: table.id }, autoFilter });
  }
  return resolved;
}

export function resolveActiveAutoFilter(sheet: WorksheetModel, column?: number): AutoFilterModel | undefined {
  return (column === undefined
    ? resolveAutoFilters(sheet)[0]
    : resolveAutoFilters(sheet).find((entry) => column >= entry.autoFilter.range.startColumn && column <= entry.autoFilter.range.endColumn))?.autoFilter;
}

export function resolveFilterOwner(sheet: WorksheetModel, column?: number): FilterOwner | undefined {
  return (column === undefined
    ? resolveAutoFilters(sheet)[0]
    : resolveAutoFilters(sheet).find((entry) => column >= entry.autoFilter.range.startColumn && column <= entry.autoFilter.range.endColumn))?.owner;
}

export function validateFilterModelOwnership(
  filter: AutoFilterModel,
  sheetId: string,
  ownerRange: RangeRef,
  label: string,
): AutoFilterModel {
  const range = normalizeRangeRef(filter.range);
  if (filter.sheetId !== sheetId || range.sheetId !== sheetId) throw new Error(`${label} must target its worksheet`);
  if (!rangesEqual(range, normalizeRangeRef(ownerRange))) throw new Error(`${label} range must equal its owner range`);
  for (const [key, column] of Object.entries(filter.columns)) {
    if (Number(key) !== column.column || column.column < range.startColumn || column.column > range.endColumn) {
      throw new Error(`${label} column is outside its range`);
    }
  }
  return { ...structuredClone(filter), range };
}

export function validateFilterOwnership(
  sheet: WorksheetModel,
  candidate: AutoFilterModel,
  owner: FilterOwner,
): AutoFilterModel {
  const normalized = normalizeRangeRef(candidate.range);
  if (normalized.sheetId !== sheet.id || candidate.sheetId !== sheet.id) throw new Error('AutoFilter must target its worksheet');
  if (owner.kind === 'worksheet') {
    if (sheet.sheetTables.some((table) => table.sheetId === sheet.id && table.autoFilter && rangesOverlap(normalized, table.autoFilter.range))) {
      throw new Error('Worksheet AutoFilter cannot overlap a Table AutoFilter');
    }
  } else {
    const table = sheet.sheetTables.find((entry) => entry.id === owner.tableId && entry.sheetId === sheet.id);
    if (!table) throw new Error(`Sheet Table not found: ${owner.tableId}`);
    validateFilterModelOwnership(candidate, sheet.id, table.range, 'Table AutoFilter');
    if (sheet.autoFilter && rangesOverlap(normalized, sheet.autoFilter.range)) {
      throw new Error('Table AutoFilter cannot overlap a Worksheet AutoFilter');
    }
    if (sheet.sheetTables.some((entry) => entry.id !== table.id && entry.sheetId === sheet.id && entry.autoFilter && rangesOverlap(normalized, entry.autoFilter.range))) {
      throw new Error('Table AutoFilter cannot overlap another Table AutoFilter');
    }
  }
  return { ...structuredClone(candidate), range: normalized };
}

export function tableFilterColumns(table: SheetTableModel): number[] {
  if (!table.showFilterButton) return [];
  const range = normalizeRangeRef(table.range);
  const columns: number[] = [];
  for (let column = range.startColumn; column <= range.endColumn; column++) columns.push(column);
  return columns;
}

export function resolveActiveFilterColumns(sheet: WorksheetModel): number[] {
  const fromCriteria = resolveAutoFilters(sheet).flatMap(({ autoFilter }) => Object.values(autoFilter.columns)
    .filter((column) => Boolean(column.criterion))
    .map((column) => column.column)
    .map(Number)
    .filter((column) => Number.isFinite(column)));
  return [...new Set(fromCriteria)].sort((left, right) => left - right);
}

export function resolveFilterRangeColumns(sheet: WorksheetModel): number[] {
  const columns = new Set<number>();
  for (const { autoFilter } of resolveAutoFilters(sheet)) {
    const range = normalizeRangeRef(autoFilter.range);
    for (let column = range.startColumn; column <= range.endColumn; column++) columns.add(column);
  }
  return [...columns].sort((left, right) => left - right);
}

export interface FilterButtonState {
  column: number;
  available: boolean;
  active: boolean;
  sorted: boolean;
  hiddenButton: boolean;
  showButton: boolean;
}

export function resolveFilterButtonStates(sheet: WorksheetModel): FilterButtonState[] {
  const states = new Map<number, FilterButtonState>();
  for (const { autoFilter } of resolveAutoFilters(sheet)) {
    const sorted = new Set(autoFilter.sortState?.conditions.flatMap((condition) => {
      const result: number[] = [];
      for (let column = condition.ref.startColumn; column <= condition.ref.endColumn; column += 1) result.push(column);
      return result;
    }) ?? []);
    const range = normalizeRangeRef(autoFilter.range);
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const entry = autoFilter.columns[column] ?? { column, showButton: true, hiddenButton: false };
      states.set(column, { column, available: true, active: Boolean(entry.criterion), sorted: sorted.has(column), hiddenButton: entry.hiddenButton, showButton: entry.showButton });
    }
  }
  return [...states.values()].sort((left, right) => left.column - right.column);
}

export interface FilterButtonCell {
  row: number;
  column: number;
  active: boolean;
  sorted: boolean;
}

function rangesEqual(left: RangeRef, right: RangeRef): boolean {
  const a = normalizeRangeRef(left);
  const b = normalizeRangeRef(right);
  return a.startRow === b.startRow
    && a.endRow === b.endRow
    && a.startColumn === b.startColumn
    && a.endColumn === b.endColumn;
}

function rangesOverlap(left: RangeRef, right: RangeRef): boolean {
  const a = normalizeRangeRef(left);
  const b = normalizeRangeRef(right);
  return a.startRow <= b.endRow && b.startRow <= a.endRow
    && a.startColumn <= b.endColumn && b.startColumn <= a.endColumn;
}

export function findSheetTableForFilter(sheet: WorksheetModel): SheetTableModel | undefined {
  const autoFilter = resolveActiveAutoFilter(sheet);
  if (!autoFilter) return undefined;
  const filterRange = normalizeRangeRef(autoFilter.range);
  return sheet.sheetTables.find(
    (table) => table.sheetId === sheet.id
      && table.showFilterButton
      && rangesEqual(table.range, filterRange),
  );
}

/** 表筛选漏斗绘制在表头行；普通区域筛选仍走列头字母条 */
export function resolveFilterButtonCells(sheet: WorksheetModel): FilterButtonCell[] {
  const buttons: FilterButtonCell[] = [];
  for (const { owner, autoFilter } of resolveAutoFilters(sheet)) {
    const range = normalizeRangeRef(autoFilter.range);
    const table = owner.kind === 'table' ? sheet.sheetTables.find((entry) => entry.id === owner.tableId) : undefined;
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const entry = autoFilter.columns[column];
      if (entry?.showButton === false || entry?.hiddenButton === true) continue;
      const sorted = autoFilter.sortState?.conditions.some((condition) => column >= condition.ref.startColumn && column <= condition.ref.endColumn) ?? false;
      buttons.push({ row: table?.hasHeaderRow ? range.startRow : range.startRow, column, active: Boolean(entry?.criterion), sorted });
    }
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
  const escapedTable = tableName.replace(/]/g, ']]');
  const escapedColumn = columnName.replace(/]/g, ']]');
  return `=SUBTOTAL(${code},${escapedTable}[${escapedColumn}])`;
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
  if (enabled && table.hasTotalRow) throw new Error('Total row is already enabled');
  if (!enabled && !table.hasTotalRow) throw new Error('Total row is already disabled');
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
  if (table.hasHeaderRow && shrunkRange.endRow < shrunkRange.startRow + 1) {
    throw new Error('Cannot disable the only body row while retaining the table header');
  }
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
