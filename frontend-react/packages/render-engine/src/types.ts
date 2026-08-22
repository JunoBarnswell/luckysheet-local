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

export interface CellRenderStyle {
  background?: string;
  textColor?: string;
  font?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  padding?: number;
  wrapText?: boolean;
}

export interface CellRenderData {
  value: CellValue;
  displayValue?: string;
  style?: CellRenderStyle;
  error?: string;
}

export type CellProvider = (address: CellAddress) => CellRenderData | undefined;

export interface RenderTheme {
  canvasBackground: string;
  gridLine: string;
  cellText: string;
  cellFont: string;
  cellPadding: number;
  defaultCellBackground: string;
}

export const DEFAULT_RENDER_THEME: RenderTheme = {
  canvasBackground: '#ffffff',
  gridLine: '#d9dee7',
  cellText: '#1f2937',
  cellFont: '13px Arial, sans-serif',
  cellPadding: 6,
  defaultCellBackground: '#ffffff',
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
  { id: 'overlay', zIndex: 2, scrollable: false, pointerEvents: 'none' },
];
