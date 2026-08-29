import { sha256Hex, type ChartSeriesModel, type RangeRef, type SheetId, type WorkbookModel, type WorksheetModel } from '@react-sheets/core-model';
import { usedRangeOfSheet } from '../../application-helpers';
import {
  computePrintPages,
  createDefaultPrintLayout,
  getPrintDocument,
  type PageSetup,
  type PaperSize,
  type PrintDocument,
  type PrintLayout,
  type PrintLayoutModel,
  type PrintPageInfo,
  type PrintProjection,
  type PrintProjectionCell,
  type PrintProjectionDrawing,
  type PrintChartProjection,
} from './index';

const MM_TO_PT = 72 / 25.4;

export interface PrintPageSnapshot {
  page: number;
  sheetId: SheetId;
  range: RangeRef;
}

export interface PrintSnapshot {
  layout: PrintLayout;
  model: PrintLayoutModel;
  pages: PrintPageInfo[];
  pageSnapshots: PrintPageSnapshot[];
  printArea: RangeRef;
  pageCount: number;
}

export interface PrintPreviewCommandParams {
  layout: PrintLayout;
  sheetId?: SheetId;
  range?: RangeRef;
}

export interface PrintAreaSetCommandParams {
  sheetId: SheetId;
  range: RangeRef;
}

export type PrintCellReader = (sheet: WorksheetModel, row: number, column: number) => import('@react-sheets/core-model').CellData | undefined;

export interface PrintProjectionOptions {
  /** Canonical resolved-cell reader (formula/spill/data-block aware). */
  readCell?: PrintCellReader;
  /** AssetStore-resolved bytes; print never substitutes an asset id placeholder. */
  assetBytes?: Readonly<Record<string, Uint8Array>>;
  /** Browser preview URLs resolved through the same AssetStore ownership. */
  assetUrls?: Readonly<Record<string, string>>;
  /** Runtime-provided Pivot/table chart data; worksheet charts use readCell. */
  readChart?: (payload: import('@react-sheets/core-model').ChartDrawingPayload) => PrintChartProjection | undefined;
}

function mmToPt(mm: number): number {
  return Math.round(mm * MM_TO_PT);
}

function mapPaperSize(paper: PrintLayout['paper']): PaperSize {
  switch (paper) {
    case 'A3':
      return 'a3';
    case 'Letter':
      return 'letter';
    case 'Legal':
      return 'legal';
    case 'A4':
    default:
      return 'a4';
  }
}

export function printLayoutToPageSetup(layout: PrintLayout): PageSetup {
  return {
    paperSize: mapPaperSize(layout.paper),
    orientation: layout.orientation,
    margins: {
      top: mmToPt(layout.margin.top),
      right: mmToPt(layout.margin.right),
      bottom: mmToPt(layout.margin.bottom),
      left: mmToPt(layout.margin.left),
      header: 36,
      footer: 36,
    },
    scale: layout.scale ?? 100,
    fitToWidth: layout.fitToWidth ? 1 : undefined,
    fitToHeight: layout.fitToHeight ? 1 : undefined,
    printGridlines: layout.printGridlines ?? false,
    printHeadings: layout.printHeadings ?? false,
    centerHorizontally: layout.centerHorizontally ?? false,
    centerVertically: layout.centerVertically ?? false,
    headerText: layout.headerText,
    footerText: layout.footerText,
  };
}

export function pageSetupToPrintLayout(setup: PageSetup): PrintLayout {
  const ptToMm = (pt: number) => Math.round((pt / MM_TO_PT) * 10) / 10;
  const paper: PrintLayout['paper'] =
    setup.paperSize === 'a3' ? 'A3' : setup.paperSize === 'letter' ? 'Letter' : setup.paperSize === 'legal' ? 'Legal' : 'A4';
  return {
    paper,
    orientation: setup.orientation,
    margin: {
      top: ptToMm(setup.margins.top),
      right: ptToMm(setup.margins.right),
      bottom: ptToMm(setup.margins.bottom),
      left: ptToMm(setup.margins.left),
    },
    scale: setup.scale,
    fitToWidth: Boolean(setup.fitToWidth),
    fitToHeight: Boolean(setup.fitToHeight),
    printGridlines: setup.printGridlines,
    printHeadings: setup.printHeadings,
    centerHorizontally: setup.centerHorizontally,
    centerVertically: setup.centerVertically,
    headerText: setup.headerText,
    footerText: setup.footerText,
  };
}

export function resolvePrintArea(sheet: WorksheetModel, selectionRange?: RangeRef): RangeRef {
  const used = usedRangeOfSheet(sheet);
  if (!selectionRange) return used;
  const multiCell =
    selectionRange.startRow !== selectionRange.endRow ||
    selectionRange.startColumn !== selectionRange.endColumn;
  if (multiCell) {
    return {
      sheetId: sheet.id,
      startRow: selectionRange.startRow,
      endRow: selectionRange.endRow,
      startColumn: selectionRange.startColumn,
      endColumn: selectionRange.endColumn,
    };
  }
  if (sheet.cells.count() === 0) {
    return {
      sheetId: sheet.id,
      startRow: selectionRange.startRow,
      endRow: selectionRange.endRow,
      startColumn: selectionRange.startColumn,
      endColumn: selectionRange.endColumn,
    };
  }
  return used;
}

function averageRowHeight(sheet: WorksheetModel, range: RangeRef): number {
  let total = 0;
  let count = 0;
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    total += sheet.rowHeightsPx[row] ?? sheet.defaultRowHeightPx;
    count += 1;
  }
  return count > 0 ? total / count : 28;
}

function averageColumnWidth(sheet: WorksheetModel, range: RangeRef): number {
  let total = 0;
  let count = 0;
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    total += sheet.columnWidthsPx[column] ?? sheet.defaultColumnWidthPx;
    count += 1;
  }
  return count > 0 ? total / count : 80;
}

export function buildPrintLayoutModel(
  unitId: string,
  sheetId: SheetId,
  uiLayout: PrintLayout,
  printArea: RangeRef,
  document?: PrintDocument,
): PrintLayoutModel {
  const base = createDefaultPrintLayout(unitId, sheetId);
  return {
    ...base,
    pageSetup: printLayoutToPageSetup(uiLayout),
    printAreas: [{ sheetId, range: printArea }],
    pageBreaks: document?.pageBreaks ? structuredClone(document.pageBreaks) : [],
    repeatRows: uiLayout.repeatRows
      ? { start: uiLayout.repeatRows.startRow, end: uiLayout.repeatRows.endRow }
      : document?.repeatRows,
    repeatColumns: uiLayout.repeatColumns
      ? { start: uiLayout.repeatColumns.startColumn, end: uiLayout.repeatColumns.endColumn }
      : document?.repeatColumns,
  };
}

export function toPrintPageSnapshots(pages: PrintPageInfo[]): PrintPageSnapshot[] {
  return pages.map((page, index) => ({
    page: index + 1,
    sheetId: page.sheetId,
    range: page.range,
  }));
}

function displayCellValue(cell: import('@react-sheets/core-model').CellData | undefined): string {
  if (!cell) return '';
  if (cell.displayValue !== undefined) return cell.displayValue;
  const value = cell.formulaValue ?? cell.value;
  if (value && typeof value === 'object' && 'kind' in value && (value as { kind?: unknown }).kind === 'error') return (value as { code?: string }).code ?? '#VALUE!';
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

function dimensionOffset(index: number, values: Readonly<Record<number, number>>, fallback: number): number {
  let total = 0;
  for (let current = 0; current < index; current += 1) total += values[current] ?? fallback;
  return total;
}

function dimensionSpan(start: number, end: number, values: Readonly<Record<number, number>>, fallback: number, hidden: ReadonlySet<number>): number {
  let total = 0;
  for (let current = start; current <= end; current += 1) if (!hidden.has(current)) total += values[current] ?? fallback;
  return total;
}

function cellHasPrintContent(cell: import('@react-sheets/core-model').CellData | undefined): boolean {
  return Boolean(cell && (cell.value !== null || cell.formula || cell.style || cell.numberFormat || cell.displayValue !== undefined || cell.presentation));
}

/**
 * Build one page's actual cell/object projection.  Pagination remains a
 * geometry concern; this function is the single value and object read path
 * consumed by every print host so a PDF cannot fall back to metadata-only
 * rows.
 */
export function buildPrintProjection(workbook: WorkbookModel, page: PrintPageInfo, options: PrintProjectionOptions = {}): PrintProjection {
  const sheet = workbook.getSheet(page.sheetId);
  const rowHeights = page.sheetId === sheet.id ? sheet.rowHeightsPx : {};
  const columnWidths = page.sheetId === sheet.id ? sheet.columnWidthsPx : {};
  const hiddenRows = new Set(sheet.hiddenRows);
  const hiddenColumns = new Set(sheet.hiddenColumns);
  const baseRows = Array.from({ length: page.range.endRow - page.range.startRow + 1 }, (_, offset) => page.range.startRow + offset).filter((row) => !hiddenRows.has(row));
  const baseColumns = Array.from({ length: page.range.endColumn - page.range.startColumn + 1 }, (_, offset) => page.range.startColumn + offset).filter((column) => !hiddenColumns.has(column));
  const repeatedRows = page.repeatRows
    ? Array.from({ length: page.repeatRows.end - page.repeatRows.start + 1 }, (_, offset) => page.repeatRows!.start + offset).filter((row) => !hiddenRows.has(row))
    : [];
  const repeatedColumns = page.repeatColumns
    ? Array.from({ length: page.repeatColumns.end - page.repeatColumns.start + 1 }, (_, offset) => page.repeatColumns!.start + offset).filter((column) => !hiddenColumns.has(column))
    : [];
  const rows = [...new Set([...repeatedRows, ...baseRows])];
  const columns = [...new Set([...repeatedColumns, ...baseColumns])];
  const rowOffsets = new Map<number, number>();
  const columnOffsets = new Map<number, number>();
  let rowOffsetPx = 0;
  for (const row of rows) { rowOffsets.set(row, rowOffsetPx); rowOffsetPx += rowHeights[row] ?? sheet.defaultRowHeightPx; }
  let columnOffsetPx = 0;
  for (const column of columns) { columnOffsets.set(column, columnOffsetPx); columnOffsetPx += columnWidths[column] ?? sheet.defaultColumnWidthPx; }
  const cells: PrintProjectionCell[] = [];
  for (const row of rows) for (const column of columns) {
    const cell = options.readCell?.(sheet, row, column) ?? sheet.cells.get(row, column);
    if (!cellHasPrintContent(cell)) continue;
    const image = cell!.presentation?.kind === 'image' ? printImageResource(cell!.presentation.asset, options.assetBytes, options.assetUrls) : undefined;
    cells.push({ row, column, value: cell!.value, displayValue: displayCellValue(cell), ...(cell!.formula ? { formula: cell!.formula } : {}), ...(cell!.style ? { style: structuredClone(cell!.style) } : {}), ...(image ? { image } : {}), rowOffsetPx: rowOffsets.get(row)!, columnOffsetPx: columnOffsets.get(column)!, widthPx: columnWidths[column] ?? sheet.defaultColumnWidthPx, heightPx: rowHeights[row] ?? sheet.defaultRowHeightPx });
  }
  const originX = dimensionOffset(page.range.startColumn, columnWidths, sheet.defaultColumnWidthPx);
  const originY = dimensionOffset(page.range.startRow, rowHeights, sheet.defaultRowHeightPx);
  const pageWidth = dimensionSpan(page.range.startColumn, page.range.endColumn, columnWidths, sheet.defaultColumnWidthPx, hiddenColumns);
  const pageHeight = dimensionSpan(page.range.startRow, page.range.endRow, rowHeights, sheet.defaultRowHeightPx, hiddenRows);
  const repeatWidth = repeatedColumns.reduce((total, column) => total + (columnWidths[column] ?? sheet.defaultColumnWidthPx), 0);
  const repeatHeight = repeatedRows.reduce((total, row) => total + (rowHeights[row] ?? sheet.defaultRowHeightPx), 0);
  const drawings: PrintProjectionDrawing[] = [];
  for (const drawing of sheet.drawings) {
    const left = drawing.transform.x;
    const top = drawing.transform.y;
    const right = left + drawing.transform.width;
    const bottom = top + drawing.transform.height;
    if (right <= originX || left >= originX + pageWidth || bottom <= originY || top >= originY + pageHeight) continue;
    const payload = sheet.drawingPayloads.get(drawing.payloadId);
    if (!payload) throw new Error(`Print projection drawing payload is missing: ${drawing.payloadId}`);
    const drawingRow = drawing.anchor.row;
    const drawingColumn = drawing.anchor.column;
    const repeatedRow = drawingRow !== undefined && page.repeatRows && drawingRow >= page.repeatRows.start && drawingRow <= page.repeatRows.end;
    const repeatedColumn = drawingColumn !== undefined && page.repeatColumns && drawingColumn >= page.repeatColumns.start && drawingColumn <= page.repeatColumns.end;
    const image = payload.kind === 'image' ? printImageResource(payload.asset, options.assetBytes, options.assetUrls) : undefined;
    const chart = payload.kind === 'chart' ? resolvePrintChart(payload, sheet, options.readCell, options.readChart) : undefined;
    drawings.push({ drawing: structuredClone(drawing), payload: structuredClone(payload), ...(image ? { image } : {}), ...(chart ? { chart } : {}), xPx: (repeatedColumn ? 0 : repeatWidth) + left - originX, yPx: (repeatedRow ? 0 : repeatHeight) + top - originY, widthPx: drawing.transform.width, heightPx: drawing.transform.height });
  }
  return { schema: 'PrintProjection', page: structuredClone(page), cells, drawings, visibleRows: rows, visibleColumns: columns, scaleX: page.scaleX ?? 1, scaleY: page.scaleY ?? 1, contentWidthPx: Math.max(page.contentWidthPx ?? 0, columnOffsetPx), contentHeightPx: Math.max(page.contentHeightPx ?? 0, rowOffsetPx) };
}

function printImageResource(
  asset: import('@react-sheets/core-model').AssetRef,
  bytesById: Readonly<Record<string, Uint8Array>> | undefined,
  urlsById: Readonly<Record<string, string>> | undefined,
): { asset: import('@react-sheets/core-model').AssetRef; bytes?: Uint8Array; url?: string } {
  const bytes = bytesById?.[asset.assetId];
  if (bytesById && !bytes) throw new Error(`NATIVE_PRINT_RESOURCE_UNAVAILABLE: ${asset.assetId}`);
  if (bytes && bytes.byteLength !== asset.byteLength) throw new Error(`NATIVE_PRINT_RESOURCE_UNAVAILABLE: ${asset.assetId}`);
  if (bytes && sha256Hex(bytes) !== asset.contentHash) throw new Error(`NATIVE_PRINT_RESOURCE_HASH_MISMATCH: ${asset.assetId}`);
  const url = urlsById?.[asset.assetId];
  return {
    asset: structuredClone(asset),
    ...(bytes ? { bytes: bytes.slice() } : {}),
    ...(url ? { url } : {}),
  };
}

function resolvePrintChart(
  payload: import('@react-sheets/core-model').ChartDrawingPayload,
  sheet: WorksheetModel,
  readCell: PrintCellReader | undefined,
  runtimeReader: ((payload: import('@react-sheets/core-model').ChartDrawingPayload) => PrintChartProjection | undefined) | undefined,
): PrintChartProjection | undefined {
  const runtime = runtimeReader?.(payload);
  if (runtime) return structuredClone(runtime);
  if (payload.source.kind !== 'worksheet-ranges') throw new Error(`NATIVE_PRINT_CHART_SOURCE_UNAVAILABLE: ${payload.chartId}`);
  const source = payload.source.ranges[0];
  if (!source || source.sheetId !== sheet.id) throw new Error(`NATIVE_PRINT_CHART_SOURCE_UNAVAILABLE: ${payload.chartId}`);
  const read = (row: number, column: number) => readCell?.(sheet, row, column) ?? sheet.cells.get(row, column);
  const scalar = (row: number, column: number): string | number | null => {
    const cell = read(row, column);
    const value = cell?.formulaValue ?? cell?.value;
    return typeof value === 'number' && Number.isFinite(value) ? value : typeof value === 'string' ? value : null;
  };
  const categories = payload.categoryRange && payload.categoryRange.sheetId === sheet.id
    ? Array.from({ length: payload.categoryRange.endRow - payload.categoryRange.startRow + 1 }, (_, index) => String(scalar(payload.categoryRange!.startRow + index, payload.categoryRange!.startColumn) ?? ''))
    : Array.from({ length: Math.max(0, source.endRow - source.startRow) }, (_, index) => String(scalar(source.startRow + index + 1, source.startColumn) ?? index + 1));
  const declarations: readonly ChartSeriesModel[] = payload.series ?? Array.from({ length: Math.max(1, source.endColumn - source.startColumn) }, (_, index): ChartSeriesModel => ({ id: `series:${index + 1}`, name: String(scalar(source.startRow, source.startColumn + index + 1) ?? `Series ${index + 1}`), range: { ...source, startColumn: source.startColumn + index + 1, endColumn: source.startColumn + index + 1 } }));
  return {
    categories,
    series: declarations.map((entry, index) => {
      const range = entry.yRange ?? entry.range;
      const values = Array.from({ length: Math.max(0, range.endRow - range.startRow + 1) }, (_, offset) => {
        const value = scalar(range.startRow + offset, range.startColumn);
        return typeof value === 'number' ? value : null;
      });
      return { id: entry.id ?? `series:${index + 1}`, name: entry.name || `Series ${index + 1}`, values, ...(entry.color ? { color: entry.color } : {}) };
    }),
  };
}

export function buildPrintSnapshot(
  workbook: WorkbookModel,
  activeSheetId: SheetId,
  uiLayout?: PrintLayout,
  selectionRange?: RangeRef,
): PrintSnapshot {
  const sheet = workbook.getSheet(activeSheetId);
  const document = getPrintDocument(workbook, activeSheetId);
  const storedLayout = pageSetupToPrintLayout(document.pageSetup);
  const effectiveLayout = uiLayout ?? {
    ...storedLayout,
    repeatRows: document.repeatRows === undefined ? undefined : {
      sheetId: activeSheetId,
      startRow: document.repeatRows.start,
      endRow: document.repeatRows.end,
      startColumn: 0,
      endColumn: Math.max(0, sheet.columnCount - 1),
    },
    repeatColumns: document.repeatColumns === undefined ? undefined : {
      sheetId: activeSheetId,
      startRow: 0,
      endRow: Math.max(0, sheet.rowCount - 1),
      startColumn: document.repeatColumns.start,
      endColumn: document.repeatColumns.end,
    },
  };
  const storedArea = document.printAreas.find((area) => area.sheetId === activeSheetId)?.range;
  const printArea = selectionRange ? resolvePrintArea(sheet, selectionRange) : storedArea ?? resolvePrintArea(sheet);
  const model = buildPrintLayoutModel(workbook.unitId, activeSheetId, effectiveLayout, printArea, document);
  model.rowHeights = { ...sheet.rowHeightsPx };
  model.columnWidths = { ...sheet.columnWidthsPx };
  model.hiddenRows = new Set(sheet.hiddenRows);
  model.hiddenColumns = new Set(sheet.hiddenColumns);
  const rowHeight = averageRowHeight(sheet, printArea);
  const colWidth = averageColumnWidth(sheet, printArea);
  const pages = computePrintPages(model, rowHeight, colWidth);
  const pageSnapshots = toPrintPageSnapshots(pages);
  return {
    layout: effectiveLayout,
    model,
    pages,
    pageSnapshots,
    printArea,
    pageCount: pageSnapshots.length,
  };
}

export function summarizePrintSnapshot(snapshot: PrintSnapshot): string {
  const { printArea, pageCount } = snapshot;
  return `${pageCount} page(s) · rows ${printArea.startRow + 1}-${printArea.endRow + 1} · columns ${printArea.startColumn + 1}-${printArea.endColumn + 1}`;
}
