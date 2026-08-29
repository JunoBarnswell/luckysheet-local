import type { RangeRef, TableScalar, WorksheetModel } from '@react-sheets/core-model';
import { normalizeRangeRef } from '@react-sheets/sheet-features';

export type RangePreprocessMode = 'none' | 'all' | 'sample';

export interface RangePreprocessOptions {
  sheet: WorksheetModel;
  seed: RangeRef;
  mode: RangePreprocessMode;
  sampleRowLimit?: number;
  readValue?: (row: number, column: number) => TableScalar;
}

export interface RangePreprocessResult {
  range: RangeRef;
  headerRow: number | null;
  sampleRows: TableScalar[][];
  source: 'selection' | 'sheet-table' | 'data-region' | 'continuous-region';
}

/**
 * Resolve one explicit data-region owner before connector/Table/Pivot creation.
 * The algorithm operates on occupied coordinates and never scans logical empty
 * rows outside the worksheet used range.
 */
export function preprocessRange(options: RangePreprocessOptions): RangePreprocessResult {
  const { sheet } = options;
  const seed = normalizeRangeRef(options.seed);
  if (seed.sheetId !== sheet.id) throw new Error('RANGE_PREPROCESS_SHEET_MISMATCH: seed targets another worksheet');
  if (seed.endRow >= sheet.rowCount || seed.endColumn >= sheet.columnCount) throw new Error('RANGE_PREPROCESS_OUT_OF_BOUNDS: seed exceeds worksheet bounds');
  const readValue = options.readValue ?? ((row, column) => {
    const cell = sheet.cells.get(row, column);
    return (cell?.formulaValue ?? cell?.value ?? null) as TableScalar;
  });

  const table = sheet.sheetTables.find((candidate) => intersects(candidate.range, seed));
  if (table) return resultFor(table.range, 'sheet-table', options, readValue);
  const region = sheet.dataRegions.find((candidate) => intersects(candidate.range, seed));
  if (region) return resultFor(region.range, 'data-region', options, readValue);
  if (options.mode === 'none' || seed.startRow !== seed.endRow || seed.startColumn !== seed.endColumn) {
    return resultFor(seed, 'selection', options, readValue);
  }

  const used = sheet.usedRange;
  const rows = new Map<number, Set<number>>();
  const columns = new Map<number, Set<number>>();
  sheet.cells.forEachInRange(used.startRow, used.endRow, used.startColumn, used.endColumn, (cell, row, column) => {
    const value = cell.formulaValue ?? cell.value;
    if (value === null || value === undefined || value === '') return;
    const rowColumns = rows.get(row) ?? new Set<number>();
    rowColumns.add(column);
    rows.set(row, rowColumns);
    const columnRows = columns.get(column) ?? new Set<number>();
    columnRows.add(row);
    columns.set(column, columnRows);
  });
  if (!rows.get(seed.startRow)?.has(seed.startColumn)) {
    throw new Error('RANGE_PREPROCESS_EMPTY: the selected cell is not inside a populated region');
  }

  const range = { ...seed };
  let changed = true;
  while (changed) {
    changed = false;
    if (range.startRow > used.startRow && rowIntersects(rows.get(range.startRow - 1), range.startColumn, range.endColumn)) {
      range.startRow -= 1; changed = true;
    }
    if (range.endRow < used.endRow && rowIntersects(rows.get(range.endRow + 1), range.startColumn, range.endColumn)) {
      range.endRow += 1; changed = true;
    }
    if (range.startColumn > used.startColumn && columnIntersects(columns.get(range.startColumn - 1), range.startRow, range.endRow)) {
      range.startColumn -= 1; changed = true;
    }
    if (range.endColumn < used.endColumn && columnIntersects(columns.get(range.endColumn + 1), range.startRow, range.endRow)) {
      range.endColumn += 1; changed = true;
    }
  }
  return resultFor(range, 'continuous-region', options, readValue);
}

function resultFor(
  rangeInput: RangeRef,
  source: RangePreprocessResult['source'],
  options: RangePreprocessOptions,
  readValue: NonNullable<RangePreprocessOptions['readValue']>,
): RangePreprocessResult {
  const range = normalizeRangeRef(rangeInput);
  const sampleRowLimit = Math.max(1, Math.min(options.sampleRowLimit ?? 32, 256));
  const sampleRows: TableScalar[][] = [];
  if (options.mode !== 'none') {
    const endRow = Math.min(range.endRow, range.startRow + sampleRowLimit - 1);
    for (let row = range.startRow; row <= endRow; row += 1) {
      const values: TableScalar[] = [];
      for (let column = range.startColumn; column <= range.endColumn; column += 1) values.push(readValue(row, column));
      sampleRows.push(values);
    }
  }
  const first = sampleRows[0] ?? [];
  const nonBlankHeaders = first.map((value) => value == null ? '' : String(value).trim());
  const headerRow = nonBlankHeaders.length > 0
    && nonBlankHeaders.every(Boolean)
    && new Set(nonBlankHeaders.map((value) => value.toLocaleLowerCase())).size === nonBlankHeaders.length
    ? range.startRow
    : null;
  return { range, headerRow, sampleRows: options.mode === 'sample' ? sampleRows : [], source };
}

function rowIntersects(columns: ReadonlySet<number> | undefined, start: number, end: number): boolean {
  if (!columns) return false;
  for (const column of columns) if (column >= start && column <= end) return true;
  return false;
}

function columnIntersects(rows: ReadonlySet<number> | undefined, start: number, end: number): boolean {
  if (!rows) return false;
  for (const row of rows) if (row >= start && row <= end) return true;
  return false;
}

function intersects(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
}
