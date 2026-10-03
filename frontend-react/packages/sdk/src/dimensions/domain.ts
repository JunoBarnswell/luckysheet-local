import { MAX_EXCEL_COLUMN_WIDTH, excelColumnWidthToPixels, pixelsToExcelColumnWidth, pointsToPixels, pixelsToPoints } from '@react-sheets/exchange-excel-ooxml';
import { DEFAULT_RENDER_THEME, hasMeasurableCellContent, measureCellAutoFit, type CellRenderData } from '@react-sheets/render-engine';
import type { CanvasSheetSnapshot, WorkbookSession } from '@react-sheets/spreadsheet-app';
import { SdkError } from '../error';
import type { DimensionsActions, DimensionResult } from './contract';
import { autoFitBlockTransferables, createAutoFitBlock, type AutoFitCellInput } from './autofit-protocol';

export interface ColumnWidthPreview {
  widthPx: number;
  excelWidth: number;
}

export const MAX_EXCEL_ROW_HEIGHT_POINTS = 409;
function invalidDimension(message: string): SdkError {
  return new SdkError('INVALID_ARGUMENT', 'dimensions', message, '请提供当前工作表允许的有限行列尺寸。');
}

export class DimensionsDomain {
  private autoFitAbort: AbortController | null = null;
  private disposed = false;
  readonly actions: DimensionsActions;
  private assertActive(): void {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'dimensions', '行列尺寸会话已释放。', '请重新打开工作簿。');
  }
  dispose(): void { this.cancelAutoFit(); this.disposed = true; }
  private reject(cause: unknown): DimensionResult {
    if (cause instanceof DOMException && cause.name === 'AbortError') return { status: 'cancelled' };
    const error = cause instanceof SdkError ? cause : new SdkError('REQUEST_REJECTED', 'dimensions.autofit', cause instanceof Error ? cause.message : 'AutoFit 失败。', '请检查工作簿权限和测量宿主后重试。', { cause });
    if (!this.disposed) this.session.notify(`${error.code}: ${error.message} ${error.recovery}`);
    return { status: 'rejected', error };
  }
  private validateIndices(indices: readonly number[], count: number): number[] {
    this.assertActive();
    if (indices.some(index => !Number.isSafeInteger(index) || index < 0 || index >= count)) throw new SdkError('INVALID_ARGUMENT', 'dimensions', '行列索引超出当前工作表。', '请重新选择有效的行列。');
    return [...new Set(indices)];
  }
  private apply(sheet: CanvasSheetSnapshot, plan: { rows?: Array<{ row: number; heightPx: number; hidden?: boolean }>; columns?: Array<{ column: number; widthPx: number; hidden?: boolean }> }): void {
    this.assertActive();
    if (this.getSheet() !== sheet) throw new SdkError('STALE_OPERATION', 'dimensions.autofit', '测量期间工作表已发生变化，未提交尺寸。', '请在当前工作表重新执行 AutoFit。');
    this.session.runCommand('sheet.dimensions.apply', { sheetId: sheet.id, ...plan });
  }

  constructor(
    private readonly session: WorkbookSession,
    private readonly getSheet: () => CanvasSheetSnapshot,
  ) {
    this.actions = Object.freeze({
      previewExcelWidth: this.previewExcelWidth.bind(this), previewRowHeight: this.previewRowHeight.bind(this), rowPoints: this.rowPoints.bind(this),
      previewPixels: this.previewPixels.bind(this), setExcelWidth: this.setExcelWidth.bind(this),
      setPixels: this.setPixels.bind(this), setHidden: this.setHidden.bind(this),
      setRowHeightPoints: this.setRowHeightPoints.bind(this), setRowPixels: this.setRowPixels.bind(this),
      setRowsHidden: this.setRowsHidden.bind(this), setDefaultExcelWidth: this.setDefaultExcelWidth.bind(this),
      cancelAutoFit: this.cancelAutoFit.bind(this), autoFit: this.autoFit.bind(this), autoFitRows: this.autoFitRows.bind(this),
    });
  }

  previewExcelWidth(value: number, defaultMode = false) {
    this.assertActive();
    const valid = Number.isFinite(value) && value >= (defaultMode ? 1 / 256 : 0) && value <= MAX_EXCEL_COLUMN_WIDTH;
    return { valid, pixels: valid ? excelColumnWidthToPixels(value, this.getSheet().maximumDigitWidthPx) : null };
  }
  previewRowHeight(points: number) {
    this.assertActive();
    const valid = Number.isFinite(points) && points >= 0 && points <= MAX_EXCEL_ROW_HEIGHT_POINTS;
    return { valid, pixels: valid ? Math.round(pointsToPixels(points)) : null };
  }
  rowPoints(heightPx: number): number { this.assertActive(); return pixelsToPoints(heightPx); }
  previewPixels(widthPx: number): ColumnWidthPreview {
    this.assertActive();
    const maximumDigitWidthPx = this.getSheet().maximumDigitWidthPx;
    if (!Number.isFinite(widthPx) || widthPx < 0) throw invalidDimension('Column width must be a finite non-negative pixel value');
    const bounded = Math.max(0, widthPx);
    return { widthPx: Math.round(bounded), excelWidth: pixelsToExcelColumnWidth(bounded, maximumDigitWidthPx) };
  }

  setExcelWidth(columns: readonly number[], excelWidth: number): void {
    if (!Number.isFinite(excelWidth) || excelWidth < 0 || excelWidth > MAX_EXCEL_COLUMN_WIDTH) throw invalidDimension('Excel column width must be between 0 and 255');
    if (excelWidth === 0) {
      this.setHidden(columns, true);
      return;
    }
    const sheet = this.getSheet();
    const indices = this.validateIndices(columns, sheet.columnCount);
    this.apply(sheet, { columns: indices.map(column => ({ column, widthPx: Math.max(1, Math.round(excelColumnWidthToPixels(excelWidth, sheet.maximumDigitWidthPx))), hidden: false })) });
  }

  setPixels(columns: readonly number[], widthPx: number): void {
    if (!Number.isFinite(widthPx) || widthPx <= 0) throw invalidDimension('Column width must be positive pixels');
    const sheet = this.getSheet();
    this.apply(sheet, { columns: this.validateIndices(columns, sheet.columnCount).map(column => ({ column, widthPx: Math.max(1, Math.round(widthPx)) })) });
  }

  setHidden(columns: readonly number[], hidden: boolean): void {
    const sheet = this.getSheet();
    this.session.runCommand('sheet.columns.visibility.set', { sheetId: sheet.id, columns: this.validateIndices(columns, sheet.columnCount), hidden });
  }

  setRowHeightPoints(rows: readonly number[], points: number): void {
    if (!Number.isFinite(points) || points < 0 || points > MAX_EXCEL_ROW_HEIGHT_POINTS) throw invalidDimension(`Row height must be between 0 and ${MAX_EXCEL_ROW_HEIGHT_POINTS} points`);
    if (points === 0) {
      this.setRowsHidden(rows, true);
      return;
    }
    this.setRowPixels(rows, pointsToPixels(points));
  }

  setRowPixels(rows: readonly number[], heightPx: number): void {
    if (!Number.isFinite(heightPx) || heightPx < 0) throw invalidDimension('Row height must be non-negative pixels');
    if (heightPx === 0) {
      this.setRowsHidden(rows, true);
      return;
    }
    const sheet = this.getSheet();
    this.apply(sheet, { rows: this.validateIndices(rows, sheet.rowCount).map(row => ({ row, heightPx, hidden: false })) });
  }

  setRowsHidden(rows: readonly number[], hidden: boolean): void {
    const sheet = this.getSheet();
    this.session.runCommand('sheet.rows.visibility.set', { sheetId: sheet.id, rows: this.validateIndices(rows, sheet.rowCount), hidden });
  }

  setDefaultExcelWidth(excelWidth: number): void {
    if (!Number.isFinite(excelWidth) || excelWidth <= 0 || excelWidth > MAX_EXCEL_COLUMN_WIDTH) throw invalidDimension('Default Excel column width must be between 0 and 255');
    this.assertActive();
    const sheet = this.getSheet();
    this.session.runCommand('sheet.column.defaultWidth.set', { sheetId: sheet.id, widthPx: excelColumnWidthToPixels(excelWidth, sheet.maximumDigitWidthPx) });
  }

  cancelAutoFit(): void {
    this.autoFitAbort?.abort();
    this.autoFitAbort = null;
  }

  async autoFit(columns: readonly number[]): Promise<DimensionResult> {
    this.cancelAutoFit();
    const controller = new AbortController();
    this.autoFitAbort = controller;
    try {
      const sheet = this.getSheet();
      const widths = await this.measureColumns(sheet, this.validateIndices(columns, sheet.columnCount), controller.signal);
      if (controller.signal.aborted) return { status: 'cancelled' };
      this.apply(sheet, { columns: widths });
      return { status: 'applied' };
    } catch (cause) { return this.reject(cause); } finally {
      if (this.autoFitAbort === controller) this.autoFitAbort = null;
    }
  }

  async autoFitRows(rows: readonly number[]): Promise<DimensionResult> {
    this.cancelAutoFit();
    const controller = new AbortController();
    this.autoFitAbort = controller;
    const sheet = this.getSheet();
    try {
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (!context) throw new Error('Canvas text measurement is unavailable');
      const heights: Array<{ row: number; heightPx: number }> = [];
      const requestedRows = this.validateIndices(rows, sheet.rowCount);
      const occupiedByRow = occupiedCellsByRow(sheet, new Set(requestedRows));
      const mergeRows = new MultiColumnMergeSpatialIndex(sheet.merges).intervalsForRows(requestedRows);
      const filterButtons = new Set(sheet.filterButtons.map((button) => `${button.row}:${button.column}`));
      for (const row of requestedRows) {
        if (controller.signal.aborted) throw new DOMException('AutoFit cancelled', 'AbortError');
        let heightPx = 8;
        for (const column of occupiedByRow.get(row) ?? []) {
          if (isCoveredByMultiColumnMerge(mergeRows.get(row), column)) continue;
          const cell = sheet.getCell(row, column);
          if (!cell || !hasMeasurableCellContent(cell)) continue;
          const availableWidthPx = sheet.columnWidthsPx[column] ?? sheet.defaultColumnWidthPx;
          heightPx = Math.max(heightPx, measureCellAutoFit(context, cell as CellRenderData, DEFAULT_RENDER_THEME, availableWidthPx, filterButtons.has(`${row}:${column}`)).heightPx);
        }
        heights.push({ row, heightPx });
        if (heights.length % 250 === 0) await yieldToBrowser();
      }
      if (controller.signal.aborted) return { status: 'cancelled' };
      this.apply(sheet, { rows: heights });
      return { status: 'applied' };
    } catch (cause) { return this.reject(cause); } finally {
      if (this.autoFitAbort === controller) this.autoFitAbort = null;
    }
  }

  private async measureColumns(sheet: CanvasSheetSnapshot, columns: number[], signal: AbortSignal): Promise<Array<{ column: number; widthPx: number }>> {
    const bounded = columns.filter((column) => column >= 0 && column < sheet.columnCount);
    if (!bounded.length) return [];
    const cells = occupiedCellsForColumns(sheet, new Set(bounded));
    if (cells.length > 5_000) {
      if (typeof Worker === 'undefined') throw new SdkError('UNSUPPORTED_FEATURE', 'dimensions.autofit', '大范围 AutoFit 需要 browser Worker。', '请使用支持 Worker 的宿主。');
      return this.measureInWorker(sheet, bounded, cells, signal);
    }
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Canvas text measurement is unavailable');
    const filterButtons = new Set(sheet.filterButtons.map((cell) => `${cell.row}:${cell.column}`));
    const maxima = new Map(bounded.map((column) => [column, 8]));
    const mergeRows = new MultiColumnMergeSpatialIndex(sheet.merges).intervalsForRows([...new Set(cells.map((cell) => cell.row))]);
    for (let index = 0; index < cells.length; index += 1) {
      if (signal.aborted) throw new DOMException('AutoFit cancelled', 'AbortError');
      const { row, column } = cells[index]!;
      if (isCoveredByMultiColumnMerge(mergeRows.get(row), column)) continue;
      const cell = sheet.getCell(row, column);
      if (!cell || !hasMeasurableCellContent(cell)) continue;
      const width = measureCellAutoFit(context, cell as CellRenderData, DEFAULT_RENDER_THEME, undefined, filterButtons.has(`${row}:${column}`)).widthPx;
      maxima.set(column, Math.max(maxima.get(column) ?? 8, width));
      if (index > 0 && index % 1_000 === 0) await yieldToBrowser();
    }
    return [...maxima].map(([column, widthPx]) => ({ column, widthPx: Math.max(8, widthPx) }));
  }

  private async measureInWorker(sheet: CanvasSheetSnapshot, columns: number[], occupiedCells: readonly OccupiedCellAddress[], signal: AbortSignal): Promise<Array<{ column: number; widthPx: number }>> {
    const worker = new Worker(new URL('./autofit-worker.ts', import.meta.url), { type: 'module' });
    const taskId = `autofit-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    let cancelTask: () => void = () => {};
    const result = new Promise<{ widths: Array<{ column: number; widthPx: number }> } | { error: Error }>(resolve => {
      worker.onmessage = (event: MessageEvent<{ kind: string; taskId: string; widths: Array<{ column: number; widthPx: number }> }>) => {
        if (event.data.kind === 'complete' && event.data.taskId === taskId) resolve({ widths: event.data.widths });
      };
      worker.onerror = event => resolve({ error: new Error(event.message || 'AutoFit worker failed') });
      worker.onmessageerror = () => resolve({ error: new Error('AutoFit worker response cannot be decoded') });
      cancelTask = () => { worker.postMessage({ kind: 'cancel', taskId }); resolve({ error: new DOMException('AutoFit cancelled', 'AbortError') }); };
      signal.addEventListener('abort', cancelTask, { once: true });
    });
    worker.postMessage({ kind: 'start', taskId, columns });
    const filterButtons = new Set(sheet.filterButtons.map((cell) => `${cell.row}:${cell.column}`));
    const mergeRows = new MultiColumnMergeSpatialIndex(sheet.merges).intervalsForRows([...new Set(occupiedCells.map((cell) => cell.row))]);
    try {
      for (let start = 0; start < occupiedCells.length; start += 1_000) {
        if (signal.aborted) throw new DOMException('AutoFit cancelled', 'AbortError');
        const cells: AutoFitCellInput[] = [];
        for (const { row, column } of occupiedCells.slice(start, start + 1_000)) {
          if (isCoveredByMultiColumnMerge(mergeRows.get(row), column)) continue;
          const cell = sheet.getCell(row, column);
          if (cell && hasMeasurableCellContent(cell)) cells.push({ column, value: cell.displayValue ?? cell.value, style: cell.style, richText: cell.richText, phonetic: cell.phonetic, filterButton: filterButtons.has(`${row}:${column}`) });
        }
        const block = createAutoFitBlock(cells);
        worker.postMessage({ kind: 'chunk', taskId, block }, autoFitBlockTransferables(block));
        await yieldToBrowser();
      }
      worker.postMessage({ kind: 'finish', taskId });
      const outcome = await result;
      if ('error' in outcome) throw outcome.error;
      return outcome.widths.map((entry) => ({ ...entry, widthPx: Math.max(8, entry.widthPx) }));
    } finally {
      signal.removeEventListener('abort', cancelTask);
      worker.terminate();
    }
  }
}

interface OccupiedCellAddress {
  row: number;
  column: number;
}

function occupiedCellsForColumns(sheet: CanvasSheetSnapshot, columns: ReadonlySet<number>): OccupiedCellAddress[] {
  const cells: OccupiedCellAddress[] = [];
  sheet.forEachOccupiedCell((row, column) => {
    if (columns.has(column)) cells.push({ row, column });
  }, { columns });
  return cells;
}

function occupiedCellsByRow(sheet: CanvasSheetSnapshot, rows: ReadonlySet<number>): Map<number, number[]> {
  const cells = new Map<number, number[]>();
  sheet.forEachOccupiedCell((row, column) => {
    if (!rows.has(row)) return;
    const columns = cells.get(row) ?? [];
    columns.push(column);
    cells.set(row, columns);
  }, { rows });
  return cells;
}

interface MergeColumnInterval {
  startColumn: number;
  endColumn: number;
}

/**
 * Sparse row sweep for AutoFit. It replaces the former requestedRows × merges
 * filter loop, while keeping merge ownership in the snapshot contract.
 */
class MultiColumnMergeSpatialIndex {
  private readonly merges: readonly CanvasSheetSnapshot['merges'][number][];

  constructor(merges: readonly CanvasSheetSnapshot['merges'][number][]) {
    this.merges = merges
      .filter((merge) => merge.range.startColumn !== merge.range.endColumn)
      .slice()
      .sort((left, right) => left.range.startRow - right.range.startRow || left.range.endRow - right.range.endRow);
  }

  intervalsForRows(rows: readonly number[]): ReadonlyMap<number, readonly MergeColumnInterval[]> {
    const result = new Map<number, readonly MergeColumnInterval[]>();
    const orderedRows = [...new Set(rows)].sort((left, right) => left - right);
    const active: CanvasSheetSnapshot['merges'][number][] = [];
    let next = 0;
    for (const row of orderedRows) {
      while (next < this.merges.length && this.merges[next]!.range.startRow <= row) active.push(this.merges[next++]!);
      const intervals = active
        .filter((merge) => merge.range.endRow >= row)
        .map((merge) => ({ startColumn: merge.range.startColumn, endColumn: merge.range.endColumn }))
        .sort((left, right) => left.startColumn - right.startColumn || left.endColumn - right.endColumn);
      if (intervals.length > 0) result.set(row, mergeIntervals(intervals));
    }
    return result;
  }
}

function mergeIntervals(intervals: readonly MergeColumnInterval[]): readonly MergeColumnInterval[] {
  const merged: MergeColumnInterval[] = [];
  for (const interval of intervals) {
    const previous = merged.at(-1);
    if (previous && interval.startColumn <= previous.endColumn + 1) previous.endColumn = Math.max(previous.endColumn, interval.endColumn);
    else merged.push({ ...interval });
  }
  return merged;
}

function isCoveredByMultiColumnMerge(intervals: readonly MergeColumnInterval[] | undefined, column: number): boolean {
  if (!intervals) return false;
  let low = 0;
  let high = intervals.length - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const interval = intervals[middle]!;
    if (column < interval.startColumn) high = middle - 1;
    else if (column > interval.endColumn) low = middle + 1;
    else return true;
  }
  return false;
}

function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}
