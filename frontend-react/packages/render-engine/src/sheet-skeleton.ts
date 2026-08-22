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

    this.rowHeightsByModel = new Array(this.rowCount);
    for (let r = 0; r < this.rowCount; r++) {
      const override = options.rowHeights
        ? (Array.isArray(options.rowHeights) ? options.rowHeights[r] : options.rowHeights.get(r))
        : undefined;
      this.rowHeightsByModel[r] = normalizeSize((override ?? this.defaultRowHeight) * this.zoom, 0);
    }
    this.columnWidthsByModel = new Array(this.columnCount);
    for (let c = 0; c < this.columnCount; c++) {
      const override = options.columnWidths
        ? (Array.isArray(options.columnWidths) ? options.columnWidths[c] : options.columnWidths.get(c))
        : undefined;
      this.columnWidthsByModel[c] = normalizeSize((override ?? this.defaultColumnWidth) * this.zoom, 0);
    }

    this.rebuildVisibleMappings();
  }

  private rebuildVisibleMappings(): void {
    this.visibleRowTops = [0];
    this.visibleRowModels = [];
    this.modelToVisibleRow = new Map();
    let top = 0;
    for (let r = 0; r < this.rowCount; r++) {
      if (this.hiddenRowSet.has(r)) continue;
      this.modelToVisibleRow.set(r, this.visibleRowModels.length);
      this.visibleRowModels.push(r);
      top += this.rowHeightsByModel[r]!;
      this.visibleRowTops.push(top);
    }

    this.visibleColumnLefts = [0];
    this.visibleColumnModels = [];
    this.modelToVisibleColumn = new Map();
    let left = 0;
    for (let c = 0; c < this.columnCount; c++) {
      if (this.hiddenColumnSet.has(c)) continue;
      this.modelToVisibleColumn.set(c, this.visibleColumnModels.length);
      this.visibleColumnModels.push(c);
      left += this.columnWidthsByModel[c]!;
      this.visibleColumnLefts.push(left);
    }
  }

  isRowHidden(modelRow: number): boolean {
    return this.hiddenRowSet.has(modelRow);
  }

  isColumnHidden(modelColumn: number): boolean {
    return this.hiddenColumnSet.has(modelColumn);
}

  get totalWidth(): number {
    return this.visibleColumnLefts.at(-1) ?? 0;
  }

  get totalHeight(): number {
    return this.visibleRowTops.at(-1) ?? 0;
  }

  get contentSize(): Size {
    return { width: this.totalWidth, height: this.totalHeight };
  }

  getRowHeight(row: number): number {
    return this.rowHeightsByModel[row] ?? 0;
  }

  getColumnWidth(column: number): number {
    return this.columnWidthsByModel[column] ?? 0;
  }

  getRowTop(modelRow: number): number {
    const visibleIndex = this.modelToVisibleRow.get(modelRow);
    if (visibleIndex === undefined) return -1;
    return this.visibleRowTops[visibleIndex] ?? 0;
  }

  getColumnLeft(modelColumn: number): number {
    const visibleIndex = this.modelToVisibleColumn.get(modelColumn);
    if (visibleIndex === undefined) return -1;
    return this.visibleColumnLefts[visibleIndex] ?? 0;
  }

  /** 可见行数(用于分页/统计) */
  get visibleRowCount(): number {
    return this.visibleRowModels.length;
  }

  get visibleColumnCount(): number {
    return this.visibleColumnModels.length;
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

  /** 内容偏移 → 可见模型行号 */
  findRowAt(contentOffsetY: number): number {
    const visibleIndex = this.findIndexAt(contentOffsetY, this.visibleRowTops, this.totalHeight);
    return visibleIndex < 0 ? -1 : (this.visibleRowModels[visibleIndex] ?? -1);
  }

  /** 内容偏移 → 可见模型列号 */
  findColumnAt(contentOffsetX: number): number {
    const visibleIndex = this.findIndexAt(contentOffsetX, this.visibleColumnLefts, this.totalWidth);
    return visibleIndex < 0 ? -1 : (this.visibleColumnModels[visibleIndex] ?? -1);
  }

  setRowHeight(modelRow: number, heightPx: number, zoom = this.zoom): void {
    if (!this.isValidRow(modelRow)) throw new Error("Unknown row: " + modelRow);
    this.rowHeightsByModel[modelRow] = normalizeSize(heightPx * zoom, 0);
    this.rebuildVisibleMappings();
  }

  setColumnWidth(modelColumn: number, widthPx: number, zoom = this.zoom): void {
    if (!this.isValidColumn(modelColumn)) throw new Error("Unknown column: " + modelColumn);
    this.columnWidthsByModel[modelColumn] = normalizeSize(widthPx * zoom, 0);
    this.rebuildVisibleMappings();
  }

  /** 返回指定内容偏移处向上/向左最近的可见模型索引(供表头渲染遍历) */
  getVisibleRowModels(): readonly number[] {
    return this.visibleRowModels;
  }

  getVisibleColumnModels(): readonly number[] {
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