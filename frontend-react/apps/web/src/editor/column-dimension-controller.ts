import { MAX_EXCEL_COLUMN_WIDTH, excelColumnWidthToPixels, pixelsToExcelColumnWidth, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import { DEFAULT_RENDER_THEME, hasMeasurableCellContent, measureCellAutoFit, type CellRenderStyle } from '@react-sheets/render-engine';
import type { CanvasSheetSnapshot, SelectionState, WorkbookSession } from '@react-sheets/spreadsheet-app';

export interface ColumnWidthPreview {
  widthPx: number;
  excelWidth: number;
}

export const MAX_EXCEL_ROW_HEIGHT_POINTS = 409;

export class ColumnDimensionController {
  private autoFitAbort: AbortController | null = null;

  constructor(
    private readonly session: WorkbookSession,
    private readonly getSheet: () => CanvasSheetSnapshot,
    private readonly getSelection: () => SelectionState,
  ) {}

  previewPixels(widthPx: number): ColumnWidthPreview {
    const maximumDigitWidthPx = this.getSheet().maximumDigitWidthPx;
    if (!Number.isFinite(widthPx) || widthPx < 0) throw new Error('Column width must be a finite non-negative pixel value');
    const bounded = Math.max(0, widthPx);
    return { widthPx: Math.round(bounded), excelWidth: pixelsToExcelColumnWidth(bounded, maximumDigitWidthPx) };
  }

  selectedColumns(includeOrdinaryCellRanges = true): number[] {
    const sheet = this.getSheet();
    const selection = this.getSelection();
    const columns = new Set<number>();
    for (const range of selection.ranges) {
      const completeColumn = range.startRow === 0 && range.endRow >= sheet.rowCount - 1;
      if (!completeColumn && !includeOrdinaryCellRanges) continue;
      for (let column = range.startColumn; column <= range.endColumn; column += 1) columns.add(column);
    }
    if (!columns.size) columns.add(selection.activeCell.column);
    return [...columns].sort((left, right) => left - right);
  }

  selectedRows(includeOrdinaryCellRanges = true): number[] {
    const sheet = this.getSheet();
    const selection = this.getSelection();
    const rows = new Set<number>();
    for (const range of selection.ranges) {
      const completeRow = range.startColumn === 0 && range.endColumn >= sheet.columnCount - 1;
      if (!completeRow && !includeOrdinaryCellRanges) continue;
      for (let row = range.startRow; row <= range.endRow; row += 1) rows.add(row);
    }
    if (!rows.size) rows.add(selection.activeCell.row);
    return [...rows].sort((left, right) => left - right);
  }

  columnsForBoundary(boundaryColumn: number): number[] {
    const selected = this.selectedColumns(false);
    return selected.includes(boundaryColumn) ? selected : [boundaryColumn];
  }

  rowsForBoundary(boundaryRow: number): number[] {
    const selected = this.selectedRows(false);
    return selected.includes(boundaryRow) ? selected : [boundaryRow];
  }

  resizeBoundary(boundaryColumn: number, widthPx: number): void {
    this.setPixels(this.columnsForBoundary(boundaryColumn), widthPx);
  }

  setExcelWidth(columns: readonly number[], excelWidth: number): void {
    if (!Number.isFinite(excelWidth) || excelWidth < 0 || excelWidth > MAX_EXCEL_COLUMN_WIDTH) throw new Error('Excel column width must be between 0 and 255');
    if (excelWidth === 0) {
      this.setHidden(columns, true);
      return;
    }
    this.setHidden(columns, false);
    this.setPixels(columns, excelColumnWidthToPixels(excelWidth, this.getSheet().maximumDigitWidthPx));
  }

  setPixels(columns: readonly number[], widthPx: number): void {
    if (!Number.isFinite(widthPx) || widthPx <= 0) throw new Error('Column width must be positive pixels');
    this.session.resizeColumns(columns, Math.max(1, Math.round(widthPx)));
  }

  setHidden(columns: readonly number[], hidden: boolean): void {
    this.session.setColumnsHidden(columns, hidden);
  }

  setRowHeightPoints(rows: readonly number[], points: number): void {
    if (!Number.isFinite(points) || points <= 0 || points > MAX_EXCEL_ROW_HEIGHT_POINTS) throw new Error(`Row height must be between 0 and ${MAX_EXCEL_ROW_HEIGHT_POINTS} points`);
    this.session.resizeRows(rows, pointsToPixels(points));
  }

  setRowPixels(rows: readonly number[], heightPx: number): void {
    if (!Number.isFinite(heightPx) || heightPx <= 0) throw new Error('Row height must be positive pixels');
    this.session.resizeRows(rows, heightPx);
  }

  setRowsHidden(rows: readonly number[], hidden: boolean): void {
    this.session.setRowsHidden(rows, hidden);
  }

  setDefaultExcelWidth(excelWidth: number): void {
    if (!Number.isFinite(excelWidth) || excelWidth <= 0 || excelWidth > MAX_EXCEL_COLUMN_WIDTH) throw new Error('Default Excel column width must be between 0 and 255');
    this.session.setDefaultColumnWidth(excelColumnWidthToPixels(excelWidth, this.getSheet().maximumDigitWidthPx));
  }

  cancelAutoFit(): void {
    this.autoFitAbort?.abort();
    this.autoFitAbort = null;
  }

  async autoFit(columns: readonly number[]): Promise<void> {
    this.cancelAutoFit();
    const controller = new AbortController();
    this.autoFitAbort = controller;
    try {
      const widths = await this.measureColumns([...new Set(columns)], controller.signal);
      if (!controller.signal.aborted) this.session.applyColumnWidths(widths);
    } finally {
      if (this.autoFitAbort === controller) this.autoFitAbort = null;
    }
  }

  async autoFitRows(rows: readonly number[]): Promise<void> {
    this.cancelAutoFit();
    const controller = new AbortController();
    this.autoFitAbort = controller;
    const sheet = this.getSheet();
    try {
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (!context) throw new Error('Canvas text measurement is unavailable');
      const heights: Array<{ row: number; heightPx: number }> = [];
      for (const row of [...new Set(rows)]) {
        if (controller.signal.aborted) throw new DOMException('AutoFit cancelled', 'AbortError');
        if (row < 0 || row >= sheet.rowCount) continue;
        let heightPx = 8;
        for (let column = sheet.usedRange.startColumn; column <= sheet.usedRange.endColumn; column += 1) {
          if (sheet.merges.some((merge) => merge.range.startRow !== merge.range.endRow && row >= merge.range.startRow && row <= merge.range.endRow && column >= merge.range.startColumn && column <= merge.range.endColumn)) continue;
          const cell = sheet.getCell(row, column);
          if (!cell || !hasMeasurableCellContent(cell)) continue;
          const availableWidthPx = sheet.columnWidthsPx[column] ?? sheet.defaultColumnWidthPx;
          heightPx = Math.max(heightPx, measureCellAutoFit(context, { value: cell.value, displayValue: cell.displayValue, formula: cell.formula, style: cell.style }, DEFAULT_RENDER_THEME, availableWidthPx, sheet.filterButtons.some((button) => button.row === row && button.column === column)).heightPx);
        }
        heights.push({ row, heightPx });
        if (heights.length % 250 === 0) await yieldToBrowser();
      }
      if (!controller.signal.aborted) this.session.applyRowHeights(heights);
    } finally {
      if (this.autoFitAbort === controller) this.autoFitAbort = null;
    }
  }

  private async measureColumns(columns: number[], signal: AbortSignal): Promise<Array<{ column: number; widthPx: number }>> {
    const sheet = this.getSheet();
    const bounded = columns.filter((column) => column >= 0 && column < sheet.columnCount);
    if (!bounded.length) return [];
    if (typeof Worker !== 'undefined' && sheet.usedRange.endRow - sheet.usedRange.startRow > 5_000) return this.measureInWorker(sheet, bounded, signal);
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Canvas text measurement is unavailable');
    const filterButtons = new Set(sheet.filterButtons.map((cell) => `${cell.row}:${cell.column}`));
    const maxima = new Map(bounded.map((column) => [column, 8]));
    for (let row = sheet.usedRange.startRow; row <= sheet.usedRange.endRow; row += 1) {
      if (signal.aborted) throw new DOMException('AutoFit cancelled', 'AbortError');
      for (const column of bounded) {
        if (isMultiColumnMerge(sheet, row, column)) continue;
        const cell = sheet.getCell(row, column);
        if (!cell || !hasMeasurableCellContent(cell)) continue;
        const width = measureCellAutoFit(context, { value: cell.value, displayValue: cell.displayValue, formula: cell.formula, style: cell.style }, DEFAULT_RENDER_THEME, undefined, filterButtons.has(`${row}:${column}`)).widthPx;
        maxima.set(column, Math.max(maxima.get(column) ?? 8, width));
      }
      if ((row - sheet.usedRange.startRow) % 1_000 === 0) await yieldToBrowser();
    }
    return [...maxima].map(([column, widthPx]) => ({ column, widthPx: Math.max(8, widthPx) }));
  }

  private async measureInWorker(sheet: CanvasSheetSnapshot, columns: number[], signal: AbortSignal): Promise<Array<{ column: number; widthPx: number }>> {
    const worker = new Worker(new URL('./column-autofit-worker.ts', import.meta.url), { type: 'module' });
    const taskId = `autofit-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const result = new Promise<Array<{ column: number; widthPx: number }>>((resolve, reject) => {
      worker.onmessage = (event: MessageEvent<{ kind: string; taskId: string; widths: Array<{ column: number; widthPx: number }> }>) => {
        if (event.data.kind === 'complete' && event.data.taskId === taskId) resolve(event.data.widths);
      };
      worker.onerror = (event) => reject(new Error(event.message || 'AutoFit worker failed'));
      signal.addEventListener('abort', () => { worker.postMessage({ kind: 'cancel', taskId }); reject(new DOMException('AutoFit cancelled', 'AbortError')); }, { once: true });
    });
    worker.postMessage({ kind: 'start', taskId, columns });
    const filterButtons = new Set(sheet.filterButtons.map((cell) => `${cell.row}:${cell.column}`));
    try {
      for (let startRow = sheet.usedRange.startRow; startRow <= sheet.usedRange.endRow; startRow += 1_000) {
        if (signal.aborted) throw new DOMException('AutoFit cancelled', 'AbortError');
        const cells: Array<{ column: number; value: string; style?: CellRenderStyle; filterButton?: boolean }> = [];
        const endRow = Math.min(sheet.usedRange.endRow, startRow + 999);
        for (let row = startRow; row <= endRow; row += 1) for (const column of columns) {
          if (isMultiColumnMerge(sheet, row, column)) continue;
          const cell = sheet.getCell(row, column);
          if (cell && hasMeasurableCellContent(cell)) cells.push({ column, value: cell.displayValue ?? cell.value, style: cell.style, filterButton: filterButtons.has(`${row}:${column}`) });
        }
        worker.postMessage({ kind: 'chunk', taskId, cells });
        await yieldToBrowser();
      }
      worker.postMessage({ kind: 'finish', taskId });
      return (await result).map((entry) => ({ ...entry, widthPx: Math.max(8, entry.widthPx) }));
    } finally {
      worker.terminate();
    }
  }
}

function isMultiColumnMerge(sheet: CanvasSheetSnapshot, row: number, column: number): boolean {
  return sheet.merges.some((merge) => merge.range.startColumn !== merge.range.endColumn
    && row >= merge.range.startRow && row <= merge.range.endRow
    && column >= merge.range.startColumn && column <= merge.range.endColumn);
}

function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}
