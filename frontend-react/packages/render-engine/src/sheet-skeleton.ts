import { normalizeCellRange } from "./dirty-ranges";
import type { CellAddress, CellRange, Rect, Size } from "./types";

export interface SheetSkeletonOptions {
  rowCount: number;
  columnCount: number;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  rowHeights?: readonly number[] | ReadonlyMap<number, number>;
  columnWidths?: readonly number[] | ReadonlyMap<number, number>;
  /** 隐藏的模型行号集合 */
  hiddenRows?: ReadonlySet<number>;
  /** 隐藏的模型列号集合 */
  hiddenColumns?: ReadonlySet<number>;
  /** 缩放系数(1 = 100%) */
  zoom?: number;
}

function normalizeCount(value: number): number {
  if (!Number.isFinite(value)) throw new Error("Sheet dimensions must be finite");
  return Math.max(0, Math.trunc(value));
}

function normalizeSize(value: number, fallback: number): number {
  const candidate = Number.isFinite(value) ? value : fallback;
  return Math.max(0, candidate);
}

const VIRTUAL_GEOMETRY_THRESHOLD = 100_000;

function readOverrides(input: readonly number[] | ReadonlyMap<number, number> | undefined): Map<number, number> {
  if (!input) return new Map();
  if (input instanceof Map) return new Map(input);
  const overrides = new Map<number, number>();
  input.forEach((value, index) => {
    if (value !== undefined) overrides.set(index, value);
  });
  return overrides;
}

export function columnLabelOf(index: number): string {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

/**
 * 工作表几何骨架。
 * 所有坐标均为"可见布局坐标":隐藏行/列不占据空间,缩放直接体现在尺寸上。
 * 对外 API 一律使用模型行列号,内部通过 visible 映射数组换算。
 */
export class SheetSkeleton {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly defaultRowHeight: number;
  readonly defaultColumnWidth: number;
  readonly zoom: number;

  private rowHeightsByModel: number[];
  private columnWidthsByModel: number[];
  private readonly virtualRows: boolean;
  private readonly virtualColumns: boolean;
  private readonly rowHeightOverrides: Map<number, number>;
  private readonly columnWidthOverrides: Map<number, number>;
  private hiddenRowSet: Set<number>;
  private hiddenColumnSet: Set<number>;

  /** 可见模型行号 → 累计顶部 y */
  private visibleRowTops: number[] = [0];
  private visibleRowModels: number[] = [];
  private visibleColumnLefts: number[] = [0];
  private visibleColumnModels: number[] = [];
  private modelToVisibleRow = new Map<number, number>();
  private modelToVisibleColumn = new Map<number, number>();

  constructor(options: SheetSkeletonOptions) {
    this.rowCount = normalizeCount(options.rowCount);
    this.columnCount = normalizeCount(options.columnCount);
    this.defaultRowHeight = normalizeSize(options.defaultRowHeight ?? 24, 24);
    this.defaultColumnWidth = normalizeSize(options.defaultColumnWidth ?? 100, 100);
    this.zoom = options.zoom && options.zoom > 0 ? options.zoom : 1;
    this.hiddenRowSet = new Set(options.hiddenRows ?? []);
    this.hiddenColumnSet = new Set(options.hiddenColumns ?? []);
    this.virtualRows = this.rowCount > VIRTUAL_GEOMETRY_THRESHOLD;
    this.virtualColumns = this.columnCount > VIRTUAL_GEOMETRY_THRESHOLD;
    this.rowHeightOverrides = readOverrides(options.rowHeights);
    this.columnWidthOverrides = readOverrides(options.columnWidths);

    this.rowHeightsByModel = this.virtualRows ? [] : Array.from({ length: this.rowCount }, (_, row) => this.getVirtualRowHeight(row));
    this.columnWidthsByModel = this.virtualColumns ? [] : Array.from({ length: this.columnCount }, (_, column) => this.getVirtualColumnWidth(column));

    this.rebuildVisibleMappings();
  }

  private rebuildVisibleMappings(): void {
    if (this.virtualRows) {
      this.visibleRowTops = [];
      this.visibleRowModels = [];
      this.modelToVisibleRow = new Map();
    }
    this.visibleRowTops = [0];
    this.visibleRowModels = [];
    this.modelToVisibleRow = new Map();
    if (!this.virtualRows) {
      let top = 0;
      for (let r = 0; r < this.rowCount; r++) {
        if (this.hiddenRowSet.has(r)) continue;
        this.modelToVisibleRow.set(r, this.visibleRowModels.length);
        this.visibleRowModels.push(r);
        top += this.rowHeightsByModel[r]!;
        this.visibleRowTops.push(top);
      }
    }

    this.visibleColumnLefts = [0];
    this.visibleColumnModels = [];
    this.modelToVisibleColumn = new Map();
    if (!this.virtualColumns) {
      let left = 0;
      for (let c = 0; c < this.columnCount; c++) {
        if (this.hiddenColumnSet.has(c)) continue;
        this.modelToVisibleColumn.set(c, this.visibleColumnModels.length);
        this.visibleColumnModels.push(c);
        left += this.columnWidthsByModel[c]!;
        this.visibleColumnLefts.push(left);
      }
    }
  }

  isRowHidden(modelRow: number): boolean {
    return this.hiddenRowSet.has(modelRow);
  }

  isColumnHidden(modelColumn: number): boolean {
    return this.hiddenColumnSet.has(modelColumn);
}

  private getVirtualRowHeight(row: number): number {
    return normalizeSize((this.rowHeightOverrides.get(row) ?? this.defaultRowHeight) * this.zoom, 0);
  }

  private getVirtualColumnWidth(column: number): number {
    return normalizeSize((this.columnWidthOverrides.get(column) ?? this.defaultColumnWidth) * this.zoom, 0);
  }

  private visibleRowExtent(startRow: number, endRow: number): number {
    if (endRow < startRow) return 0;
    if (!this.virtualRows) {
      let total = 0;
      for (let row = startRow; row <= endRow; row++) total += this.isRowHidden(row) ? 0 : this.rowHeightsByModel[row]!;
      return total;
    }
    let total = (endRow - startRow + 1) * this.defaultRowHeight * this.zoom;
    for (const [row, height] of this.rowHeightOverrides) {
      if (row >= startRow && row <= endRow) total += (height * this.zoom) - (this.defaultRowHeight * this.zoom);
    }
    for (const row of this.hiddenRowSet) {
      if (row >= startRow && row <= endRow) total -= this.getVirtualRowHeight(row);
    }
    return Math.max(0, total);
  }

  private visibleColumnExtent(startColumn: number, endColumn: number): number {
    if (endColumn < startColumn) return 0;
    if (!this.virtualColumns) {
      let total = 0;
      for (let column = startColumn; column <= endColumn; column++) total += this.isColumnHidden(column) ? 0 : this.columnWidthsByModel[column]!;
      return total;
    }
    let total = (endColumn - startColumn + 1) * this.defaultColumnWidth * this.zoom;
    for (const [column, width] of this.columnWidthOverrides) {
      if (column >= startColumn && column <= endColumn) total += (width * this.zoom) - (this.defaultColumnWidth * this.zoom);
    }
    for (const column of this.hiddenColumnSet) {
      if (column >= startColumn && column <= endColumn) total -= this.getVirtualColumnWidth(column);
    }
    return Math.max(0, total);
  }

  private virtualRowTop(row: number): number {
    return this.visibleRowExtent(0, row - 1);
  }

  private virtualColumnLeft(column: number): number {
    return this.visibleColumnExtent(0, column - 1);
  }

  get totalWidth(): number {
    return this.virtualColumns
      ? this.visibleColumnExtent(0, this.columnCount - 1)
      : this.visibleColumnLefts.at(-1) ?? 0;
  }

  get totalHeight(): number {
    return this.virtualRows
      ? this.visibleRowExtent(0, this.rowCount - 1)
      : this.visibleRowTops.at(-1) ?? 0;
  }

  get contentSize(): Size {
    return { width: this.totalWidth, height: this.totalHeight };
  }

  getRowHeight(row: number): number {
    return this.virtualRows ? this.getVirtualRowHeight(row) : this.rowHeightsByModel[row] ?? 0;
  }

  getColumnWidth(column: number): number {
    return this.virtualColumns ? this.getVirtualColumnWidth(column) : this.columnWidthsByModel[column] ?? 0;
  }

  getRowTop(modelRow: number): number {
    if (this.virtualRows) return this.isRowHidden(modelRow) ? -1 : this.virtualRowTop(modelRow);
    const visibleIndex = this.modelToVisibleRow.get(modelRow);
    if (visibleIndex === undefined) return -1;
    return this.visibleRowTops[visibleIndex] ?? 0;
  }

  getColumnLeft(modelColumn: number): number {
    if (this.virtualColumns) return this.isColumnHidden(modelColumn) ? -1 : this.virtualColumnLeft(modelColumn);
    const visibleIndex = this.modelToVisibleColumn.get(modelColumn);
    if (visibleIndex === undefined) return -1;
    return this.visibleColumnLefts[visibleIndex] ?? 0;
  }

  /** 可见行数(用于分页/统计) */
  get visibleRowCount(): number {
    return this.virtualRows ? this.rowCount - this.hiddenRowSet.size : this.visibleRowModels.length;
  }

  get visibleColumnCount(): number {
    return this.virtualColumns ? this.columnCount - this.hiddenColumnSet.size : this.visibleColumnModels.length;
  }

  getCellRect(modelRow: number, modelColumn: number): Rect | null {
    if (!this.isValidRow(modelRow) || !this.isValidColumn(modelColumn)) return null;
    if (this.isRowHidden(modelRow) || this.isColumnHidden(modelColumn)) return null;
    return {
      x: this.getColumnLeft(modelColumn),
      y: this.getRowTop(modelRow),
      width: this.getColumnWidth(modelColumn),
      height: this.getRowHeight(modelRow),
    };
  }

  getRangeRect(range: CellRange): Rect | null {
    const normalized = normalizeCellRange(range);
    if (!normalized || this.visibleRowCount === 0 || this.visibleColumnCount === 0) return null;

    if (this.virtualRows || this.virtualColumns) {
      const x = this.getColumnLeft(normalized.startColumn);
      const y = this.getRowTop(normalized.startRow);
      if (x < 0 || y < 0) return null;
      return {
        x,
        y,
        width: this.visibleColumnExtent(normalized.startColumn, normalized.endColumn),
        height: this.visibleRowExtent(normalized.startRow, normalized.endRow),
      };
    }

    let x = 0;
    let width = 0;
    let foundColumn = false;
    for (const modelColumn of this.visibleColumnModels) {
      if (modelColumn >= normalized.startColumn && modelColumn <= normalized.endColumn) {
        foundColumn = true;
        width += this.columnWidthsByModel[modelColumn]!;
      } else if (foundColumn) break;
      else x += this.columnWidthsByModel[modelColumn]!;
    }
    if (!foundColumn) return null;

    let y = 0;
    let height = 0;
    let foundRow = false;
    for (const modelRow of this.visibleRowModels) {
      if (modelRow >= normalized.startRow && modelRow <= normalized.endRow) {
        foundRow = true;
        height += this.rowHeightsByModel[modelRow]!;
      } else if (foundRow) break;
      else y += this.rowHeightsByModel[modelRow]!;
    }
    if (!foundRow) return null;
    return { x, y, width, height };
  }

  getVisibleRange(sheetRect: Rect): CellRange | null {
    if (this.visibleRowCount === 0 || this.visibleColumnCount === 0) return null;
    if (sheetRect.width <= 0 || sheetRect.height <= 0) return null;
    const left = Math.max(0, sheetRect.x);
    const top = Math.max(0, sheetRect.y);
    const right = Math.min(this.totalWidth, sheetRect.x + sheetRect.width);
    const bottom = Math.min(this.totalHeight, sheetRect.y + sheetRect.height);
    if (right <= left || bottom <= top) return null;

    const edgeEpsilon = 1e-7;
    const startColumn = this.findColumnAt(left);
    const endColumn = this.findColumnAt(Math.max(left, right - edgeEpsilon));
    const startRow = this.findRowAt(top);
    const endRow = this.findRowAt(Math.max(top, bottom - edgeEpsilon));
    if (startColumn < 0 || endColumn < 0 || startRow < 0 || endRow < 0) return null;
    return { startRow, endRow, startColumn, endColumn };
  }

  getCellAtPoint(point: { x: number; y: number }): CellAddress | null {
    if (point.x < 0 || point.y < 0 || point.x >= this.totalWidth || point.y >= this.totalHeight) return null;
    const column = this.findColumnAt(point.x);
    const row = this.findRowAt(point.y);
    return row < 0 || column < 0 ? null : { row, column };
  }

  /** 内容偏移 → 可见模型行号 */
  findRowAt(contentOffsetY: number): number {
    if (this.virtualRows) {
      if (!Number.isFinite(contentOffsetY) || contentOffsetY < 0 || contentOffsetY >= this.totalHeight) return -1;
      let low = 0;
      let high = this.rowCount - 1;
      while (low < high) {
        const middle = Math.ceil((low + high) / 2);
        if (this.virtualRowTop(middle) > contentOffsetY) high = middle - 1;
        else low = middle;
      }
      while (low < this.rowCount && this.hiddenRowSet.has(low)) low += 1;
      return low < this.rowCount ? low : -1;
    }
    const visibleIndex = this.findIndexAt(contentOffsetY, this.visibleRowTops, this.totalHeight);
    return visibleIndex < 0 ? -1 : (this.visibleRowModels[visibleIndex] ?? -1);
  }

  /** 内容偏移 → 可见模型列号 */
  findColumnAt(contentOffsetX: number): number {
    if (this.virtualColumns) {
      if (!Number.isFinite(contentOffsetX) || contentOffsetX < 0 || contentOffsetX >= this.totalWidth) return -1;
      let low = 0;
      let high = this.columnCount - 1;
      while (low < high) {
        const middle = Math.ceil((low + high) / 2);
        if (this.virtualColumnLeft(middle) > contentOffsetX) high = middle - 1;
        else low = middle;
      }
      while (low < this.columnCount && this.hiddenColumnSet.has(low)) low += 1;
      return low < this.columnCount ? low : -1;
    }
    const visibleIndex = this.findIndexAt(contentOffsetX, this.visibleColumnLefts, this.totalWidth);
    return visibleIndex < 0 ? -1 : (this.visibleColumnModels[visibleIndex] ?? -1);
  }

  setRowHeight(modelRow: number, heightPx: number, zoom = this.zoom): void {
    if (!this.isValidRow(modelRow)) throw new Error("Unknown row: " + modelRow);
    if (this.virtualRows) this.rowHeightOverrides.set(modelRow, heightPx);
    else this.rowHeightsByModel[modelRow] = normalizeSize(heightPx * zoom, 0);
    this.rebuildVisibleMappings();
  }

  setColumnWidth(modelColumn: number, widthPx: number, zoom = this.zoom): void {
    if (!this.isValidColumn(modelColumn)) throw new Error("Unknown column: " + modelColumn);
    if (this.virtualColumns) this.columnWidthOverrides.set(modelColumn, widthPx);
    else this.columnWidthsByModel[modelColumn] = normalizeSize(widthPx * zoom, 0);
    this.rebuildVisibleMappings();
  }

  /** 返回指定内容偏移处向上/向左最近的可见模型索引(供表头渲染遍历) */
  getVisibleRowModels(): readonly number[] {
    if (this.virtualRows) return [];
    return this.visibleRowModels;
  }

  getVisibleColumnModels(): readonly number[] {
    if (this.virtualColumns) return [];
    return this.visibleColumnModels;
  }

  private isValidRow(row: number): boolean {
    return Number.isInteger(row) && row >= 0 && row < this.rowCount;
  }

  private isValidColumn(column: number): boolean {
    return Number.isInteger(column) && column >= 0 && column < this.columnCount;
  }

  /** 二分查找:返回可见序号(非模型序号) */
  private findIndexAt(offset: number, accumulation: readonly number[], total: number): number {
    const visibleCount = accumulation.length - 1;
    if (visibleCount <= 0 || total <= 0) return -1;
    if (!Number.isFinite(offset) || offset < 0) return 0;
    const bounded = Math.min(offset, total - Number.EPSILON);
    let low = 0;
    let high = visibleCount;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const nextBoundary = accumulation[middle + 1] ?? total;
      if (nextBoundary <= bounded) low = middle + 1;
      else high = middle;
    }
    return Math.min(low, visibleCount - 1);
  }
}
