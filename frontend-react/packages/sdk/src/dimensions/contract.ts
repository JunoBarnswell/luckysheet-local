import type { SdkError } from '../error';
export type DimensionResult = { readonly status: 'applied' | 'cancelled' } | { readonly status: 'rejected'; readonly error: SdkError };
export interface DimensionsActions {
  previewExcelWidth(value: number, defaultMode?: boolean): { valid: boolean; pixels: number | null };
  previewRowHeight(points: number): { valid: boolean; pixels: number | null };
  rowPoints(heightPx: number): number;
  previewPixels(widthPx: number): { widthPx: number; excelWidth: number };
  setExcelWidth(columns: readonly number[], excelWidth: number): void;
  setPixels(columns: readonly number[], widthPx: number): void;
  setHidden(columns: readonly number[], hidden: boolean): void;
  setRowHeightPoints(rows: readonly number[], points: number): void;
  setRowPixels(rows: readonly number[], heightPx: number): void;
  setRowsHidden(rows: readonly number[], hidden: boolean): void;
  setDefaultExcelWidth(excelWidth: number): void;
  cancelAutoFit(): void;
  autoFit(columns: readonly number[]): Promise<DimensionResult>;
  autoFitRows(rows: readonly number[]): Promise<DimensionResult>;
}
