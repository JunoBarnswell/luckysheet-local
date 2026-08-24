export type LayerId = string;

/** 行标题列宽(px) */
export const ROW_HEADER_WIDTH = 46;
/** 列标题行高(px) */
export const COL_HEADER_HEIGHT = 24;

/** 拖拽命中边界(px) */
export const RESIZE_HIT_TOLERANCE_PX = 4;

export interface Point { x: number; y: number; }
export interface Size { width: number; height: number; }
export interface Rect extends Size, Point {}

export interface ViewportSnapshot extends Size {
  scrollX: number;
  scrollY: number;
  devicePixelRatio: number;
}

export interface CellAddress { row: number; column: number; }

export interface CellRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export type CellValue = string | number | boolean | null | undefined;

export type HorizontalAlignment = "left" | "center" | "right";
export type VerticalAlignment = "top" | "middle" | "bottom";

export interface BorderStyle {
  color: string;
  style: "thin" | "medium" | "thick" | "dashed" | "double";
}

export interface CellBorders {
  top?: BorderStyle;
  right?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
}

export interface CellRenderStyle {
  background?: string;
  textColor?: string;
  fontFamily?: string;
  fontSizePx?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  font?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  padding?: number;
  wrapText?: boolean;
  borders?: CellBorders;
  numberFormat?: string;
  /** 文字旋转角度(度,顺时针) */
  textRotate?: number;
}

/** 条件格式渲染负载 */
export interface ConditionalRenderOverlay {
  dataBar?: { color: string; ratio: number };
  colorScale?: string;
  icon?: "up" | "down" | "flat";
}

export interface MergeInfo {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  isAnchor: boolean;
}

export interface CellRenderData {
  value: CellValue;
  formula?: string;
  displayValue?: string;
  style?: CellRenderStyle;
  error?: string;
  merge?: MergeInfo;
  overlay?: ConditionalRenderOverlay;
  hasComment?: boolean;
  invalid?: boolean;
}

export type CellProvider = (address: CellAddress) => CellRenderData | undefined;

export interface RenderTheme {
  canvasBackground: string;
  gridLine: string;
  cellText: string;
  cellFont: string;
  cellPadding: number;
  defaultCellBackground: string;
  headerBackground: string;
  headerBorder: string;
  headerText: string;
  selectionBorder: string;
  selectionBackground: string;
  editingBorder: string;
  fillHandleSize: number;
  invalidColor: string;
  commentMarkColor: string;
}

export const DEFAULT_RENDER_THEME: RenderTheme = {
  canvasBackground: "#ffffff",
  gridLine: "#e2e8f0",
  cellText: "#1e293b",
  cellFont: '13px Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
  cellPadding: 6,
  defaultCellBackground: "#ffffff",
  headerBackground: "#f8fafc",
  headerBorder: "#cbd5e1",
  headerText: "#475569",
  selectionBorder: "#2563eb",
  selectionBackground: "rgba(37, 99, 235, 0.10)",
  editingBorder: "#2563eb",
  fillHandleSize: 8,
  invalidColor: "#dc2626",
  commentMarkColor: "#dc2626",
};

export interface LayerDefinition {
  id: LayerId;
  zIndex: number;
  scrollable?: boolean;
  pointerEvents?: "none" | "auto";
  opacity?: number;
}

export const DEFAULT_LAYER_DEFINITIONS: readonly LayerDefinition[] = [
  { id: "grid", zIndex: 0, scrollable: true, pointerEvents: "none" },
  { id: "content", zIndex: 1, scrollable: true, pointerEvents: "none" },
  { id: "extensions", zIndex: 2, scrollable: true, pointerEvents: "none" },
  { id: "overlay", zIndex: 3, scrollable: false, pointerEvents: "none" },
  { id: "chrome", zIndex: 4, scrollable: false, pointerEvents: "none" },
];

// ---------- 文档窗格 ----------

export interface PaneLayout {
  kind: 'none' | 'frozen' | 'split';
  xSplit: number;
  ySplit: number;
  startRow: number;
  startColumn: number;
  activePane?: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
  state: 'frozen' | 'frozenSplit' | 'split';
}

export type PaneId = "topLeft" | "topRight" | "bottomLeft" | "main";

/** 单个窗格:屏幕矩形 + 内容滚动偏移 + 可见模型范围 */
export interface RenderPane {
  id: PaneId;
  screenRect: Rect;
  contentOrigin: Point;
  visibleRange: CellRange | null;
}

export interface PaneMap {
  panes: readonly RenderPane[];
  paneAtLocalPoint(point: Point): RenderPane | null;
  paneForCell(cell: CellAddress): RenderPane | null;
}

// ---------- Chrome(表头/选区等非滚动覆盖层) ----------

export interface ChromeSelectionState {
  ranges: readonly CellRange[];
  primary: CellAddress;
  primaryIndex: number;
}

export interface ChromeRemoteCursor {
  actorId: string;
  name: string;
  color: string;
  row: number;
  column: number;
}

export interface ResizePreview {
  axis: "row" | "column";
  index: number;
  sizePx: number;
  label?: string;
}

export interface ChromeFilterButton {
  row: number;
  column: number;
}

export interface ChromeTableOutline {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export interface ChromeOutlineControl {
  axis: 'row' | 'column';
  index: number;
  level: number;
  collapsed: boolean;
  groupId: string;
}

export interface ChromeState {
  selection: ChromeSelectionState;
  editing: CellAddress | null;
  filterColumns: readonly number[];
  filterButtons: readonly ChromeFilterButton[];
  tableOutlines: readonly ChromeTableOutline[];
  outlineControls: readonly ChromeOutlineControl[];
  remoteCursors: readonly ChromeRemoteCursor[];
  resizePreview: ResizePreview | null;
  selectedFloatingId: string | null;
}

export function createEmptyChromeState(): ChromeState {
  return {
    selection: { ranges: [], primary: { row: 0, column: 0 }, primaryIndex: 0 },
    editing: null,
    filterColumns: [],
    filterButtons: [],
    tableOutlines: [],
    outlineControls: [],
    remoteCursors: [],
    resizePreview: null,
    selectedFloatingId: null,
  };
}

// ---------- 浮动对象 ----------

export type FloatingKind = "chart" | "shape" | "image";
export type FloatingHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export interface FloatingDrawable {
  kind: FloatingKind;
  id: string;
  bounds: Rect;
  draw: (context: CanvasRenderingContext2D, rect: Rect) => void;
}

export interface FloatingHit {
  kind: FloatingKind;
  id: string;
  handle?: FloatingHandle;
}

export type HeaderHitKind = "corner" | "row" | "col";

export interface HeaderHit {
  kind: HeaderHitKind;
  index: number;
  /** 命中位置距离可拖拽调整边界的像素数;undefined 表示不在调整热区 */
  resizeBoundaryPx?: number;
}
