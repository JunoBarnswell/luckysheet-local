export type LayerId = string;

export interface Point {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface Rect extends Size, Point {}

export interface ViewportSnapshot extends Size {
  scrollX: number;
  scrollY: number;
  devicePixelRatio: number;
}

export interface CellAddress {
  row: number;
  column: number;
}

export interface CellRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export type CellValue = string | number | boolean | null | undefined;

export type HorizontalAlignment = 'left' | 'center' | 'right';
export type VerticalAlignment = 'top' | 'middle' | 'bottom';

export interface BorderStyle {
  color: string;
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'double';
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
  fontSize?: number;
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
  headerText: string;
  selectionBorder: string;
  selectionBackground: string;
}

export const DEFAULT_RENDER_THEME: RenderTheme = {
  canvasBackground: '#ffffff',
  gridLine: '#e2e8f0',
  cellText: '#1e293b',
  cellFont: '13px Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
  cellPadding: 6,
  defaultCellBackground: '#ffffff',
  headerBackground: '#f8fafc',
  headerText: '#475569',
  selectionBorder: '#2563eb',
  selectionBackground: 'rgba(37, 99, 235, 0.08)',
};

export interface LayerDefinition {
  id: LayerId;
  zIndex: number;
  scrollable?: boolean;
  pointerEvents?: 'none' | 'auto';
  opacity?: number;
}

export const DEFAULT_LAYER_DEFINITIONS: readonly LayerDefinition[] = [
  { id: 'grid', zIndex: 0, scrollable: true, pointerEvents: 'none' },
  { id: 'content', zIndex: 1, scrollable: true, pointerEvents: 'none' },
  { id: 'extensions', zIndex: 2, scrollable: true, pointerEvents: 'none' },
  { id: 'overlay', zIndex: 3, scrollable: false, pointerEvents: 'none' },
];
