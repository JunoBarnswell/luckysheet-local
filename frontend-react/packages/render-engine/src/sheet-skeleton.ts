import { normalizeCellRange } from './dirty-ranges';
import type { CellAddress, CellRange, Rect, Size } from './types';

export interface SheetSkeletonOptions {
  rowCount: number;
  columnCount: number;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  rowHeights?: readonly number[] | ReadonlyMap<number, number>;
  columnWidths?: readonly number[] | ReadonlyMap<number, number>;
}

function normalizeCount(value: number): number {
  if (!Number.isFinite(value)) throw new Error('Sheet dimensions must be finite');
  return Math.max(0, Math.trunc(value));
}

function normalizeSize(value: number, fallback: number): number {
  const candidate = Number.isFinite(value) ? value : fallback;
  return Math.max(0, candidate);
}

function createSizes(
  count: number,
  defaultSize: number,
  override: readonly number[] | ReadonlyMap<number, number> | undefined,
): number[] {
  const sizes = Array.from({ length: count }, () => defaultSize);
  if (!override) return sizes;
  if (Array.isArray(override)) {
    for (let index = 0; index < Math.min(count, override.length); index += 1) {
      sizes[index] = normalizeSize(override[index] ?? defaultSize, defaultSize);
    }
    return sizes;
  }
  for (const [index, size] of (override as ReadonlyMap<number, number>)) {
    if (index >= 0 && index < count) sizes[index] = normalizeSize(size, defaultSize);
  }
  return sizes;
}

function buildAccumulation(sizes: readonly number[]): number[] {
  const accumulation = [0];
  for (const size of sizes) accumulation.push((accumulation.at(-1) ?? 0) + size);
  return accumulation;
}

export class SheetSkeleton {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly defaultRowHeight: number;
  readonly defaultColumnWidth: number;

  private rowHeights: number[];
  private columnWidths: number[];
  private rowAccumulation: number[];
  private columnAccumulation: number[];

  constructor(options: SheetSkeletonOptions) {
    this.rowCount = normalizeCount(options.rowCount);
    this.columnCount = normalizeCount(options.columnCount);
    this.defaultRowHeight = normalizeSize(options.defaultRowHeight ?? 24, 24);
    this.defaultColumnWidth = normalizeSize(options.defaultColumnWidth ?? 100, 100);
    this.rowHeights = createSizes(this.rowCount, this.defaultRowHeight, options.rowHeights);
    this.columnWidths = createSizes(this.columnCount, this.defaultColumnWidth, options.columnWidths);
    this.rowAccumulation = buildAccumulation(this.rowHeights);
    this.columnAccumulation = buildAccumulation(this.columnWidths);
  }

  get totalWidth(): number {
    return this.columnAccumulation.at(-1) ?? 0;
  }

  get totalHeight(): number {
    return this.rowAccumulation.at(-1) ?? 0;
  }

  get contentSize(): Size {
    return { width: this.totalWidth, height: this.totalHeight };
  }

  getRowHeight(row: number): number {
    return this.rowHeights[row] ?? 0;
  }

  getColumnWidth(column: number): number {
    return this.columnWidths[column] ?? 0;
  }

  getRowTop(row: number): number {
    if (row < 0) return 0;
    if (row >= this.rowCount) return this.totalHeight;
    return this.rowAccumulation[row] ?? 0;
  }

  getColumnLeft(column: number): number {
    if (column < 0) return 0;
    if (column >= this.columnCount) return this.totalWidth;
    return this.columnAccumulation[column] ?? 0;
  }

  getCellRect(row: number, column: number): Rect | null {
    if (!this.isValidRow(row) || !this.isValidColumn(column)) return null;
    return {
      x: this.getColumnLeft(column),
      y: this.getRowTop(row),
      width: this.getColumnWidth(column),
      height: this.getRowHeight(row),
    };
  }

  getRangeRect(range: CellRange): Rect | null {
    const normalized = normalizeCellRange(range);
    if (!normalized || this.rowCount === 0 || this.columnCount === 0) return null;
    const startRow = Math.min(normalized.startRow, this.rowCount - 1);
    const endRow = Math.min(normalized.endRow, this.rowCount - 1);
    const startColumn = Math.min(normalized.startColumn, this.columnCount - 1);
    const endColumn = Math.min(normalized.endColumn, this.columnCount - 1);
    if (startRow > endRow || startColumn > endColumn) return null;
    const x = this.getColumnLeft(startColumn);
    const y = this.getRowTop(startRow);
    return {
      x,
      y,
      width: this.getColumnLeft(endColumn) + this.getColumnWidth(endColumn) - x,
      height: this.getRowTop(endRow) + this.getRowHeight(endRow) - y,
    };
  }

  getVisibleRange(sheetRect: Rect): CellRange | null {
    if (this.rowCount === 0 || this.columnCount === 0 || sheetRect.width <= 0 || sheetRect.height <= 0) return null;
    const left = Math.max(0, sheetRect.x);
    const top = Math.max(0, sheetRect.y);
    const right = Math.min(this.totalWidth, sheetRect.x + sheetRect.width);
    const bottom = Math.min(this.totalHeight, sheetRect.y + sheetRect.height);
    if (right <= left || bottom <= top) return null;

    const startColumn = this.findColumnAt(left);
    const endColumn = this.findColumnAt(Math.max(left, right - Number.EPSILON));
    const startRow = this.findRowAt(top);
    const endRow = this.findRowAt(Math.max(top, bottom - Number.EPSILON));
    if (startColumn < 0 || endColumn < 0 || startRow < 0 || endRow < 0) return null;
    return { startRow, endRow, startColumn, endColumn };
  }

  getCellAtPoint(point: { x: number; y: number }): CellAddress | null {
    if (point.x < 0 || point.y < 0 || point.x >= this.totalWidth || point.y >= this.totalHeight) return null;
    const column = this.findColumnAt(point.x);
    const row = this.findRowAt(point.y);
    return row < 0 || column < 0 ? null : { row, column };
  }

  findRowAt(offset: number): number {
    return this.findIndexAt(offset, this.rowAccumulation, this.rowCount, this.totalHeight);
  }

  findColumnAt(offset: number): number {
    return this.findIndexAt(offset, this.columnAccumulation, this.columnCount, this.totalWidth);
  }

  setRowHeight(row: number, height: number): void {
    if (!this.isValidRow(row)) throw new Error(`Unknown row: ${row}`);
    this.rowHeights[row] = normalizeSize(height, this.defaultRowHeight);
    this.rowAccumulation = buildAccumulation(this.rowHeights);
  }

  setColumnWidth(column: number, width: number): void {
    if (!this.isValidColumn(column)) throw new Error(`Unknown column: ${column}`);
    this.columnWidths[column] = normalizeSize(width, this.defaultColumnWidth);
    this.columnAccumulation = buildAccumulation(this.columnWidths);
  }

  getRowHeights(): number[] {
    return [...this.rowHeights];
  }

  getColumnWidths(): number[] {
    return [...this.columnWidths];
  }

  private isValidRow(row: number): boolean {
    return Number.isInteger(row) && row >= 0 && row < this.rowCount;
  }

  private isValidColumn(column: number): boolean {
    return Number.isInteger(column) && column >= 0 && column < this.columnCount;
  }

  private findIndexAt(offset: number, accumulation: readonly number[], count: number, total: number): number {
    if (count === 0 || !Number.isFinite(offset) || total <= 0) return -1;
    const bounded = Math.min(Math.max(offset, 0), total - Number.EPSILON);
    let low = 0;
    let high = count;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const nextBoundary = accumulation[middle + 1] ?? total;
      if (nextBoundary <= bounded) low = middle + 1;
      else high = middle;
    }
    return Math.min(low, count - 1);
  }
}
