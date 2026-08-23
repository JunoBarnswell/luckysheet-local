import type {
  CellComment,
  CellNote,
  CellData,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
  DrawingObject,
  DrawingPayload,
  FreezeModel,
  MergeSpan,
  OutlineGroup,
  PivotModel,
  PivotGridProjection,
  PivotResultTree,
  RangeRef,
  SheetTableModel,
  SparklineModel,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  computeBandedCellStyle,
  computeConditionalOverlays,
  computeFilterHiddenRows,
  computeOutlineHiddenColumns,
  computeOutlineHiddenRows,
  computeSheetTableCellStyle,
  findSheetTableAt,
  mergePresentationStyles,
  resolveActiveFilterColumns,
  resolveFilterButtonCells,
  resolveOutlineControls,
  validateDataInput,
  type ConditionalOverlay,
  type FilterButtonCell,
  type OutlineControl,
} from '@react-sheets/sheet-features';
import { FormulaEngine, isFormulaError, isSpillChild, type FormulaValue } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import {
  buildPivotGridProjection,
  computePivotResult,
  getLastValidPivotResult,
  type PivotProjectionSourceState,
} from './features/pivot/engine';
import { cellAddress, columnLabel } from './address';
import { getCellNote } from '@react-sheets/core-model';
import type { DataSourceContentQuery } from './features/data-source';
import {
  findCommentThreadAt,
  getCellHyperlink,
  resolveHyperlinkDisplay,
  threadToCellComment,
} from './features/review';

/** Canvas-friendly cell snapshot — no row-array projection */
export interface CanvasCellSnapshot {
  address: string;
  displayValue?: string;
  formula?: string;
  style?: CellStyle;
  value: string;
  hasComment?: boolean;
  commentText?: string;
  comment?: CellComment;
  note?: CellNote;
  invalid?: boolean;
  hyperlink?: string;
  overlay?: ConditionalOverlay;
}

/** Bounded preview row for print UI only */
export interface PreviewRowSnapshot {
  rowNumber: number;
  cells: Array<{ value: string }>;
  height: number;
}

/** Canvas-friendly sheet snapshot — single getCell path, no SheetView DTO */
export interface CanvasSheetSnapshot {
  id: string;
  name: string;
  columns: string[];
  columnCount: number;
  rowCount: number;
  isEmpty?: boolean;
  occupiedCellCount: number;
  getCell: (row: number, column: number) => CanvasCellSnapshot | undefined;
  usedRange: RangeRef;
  /** Canonical floating-object aggregate. UI never consumes legacy projections. */
  drawings: DrawingObject[];
  /** Read-only payload projection keyed by DrawingObject.payloadId. */
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  pivots: PivotModel[];
  pivotResults: Record<string, PivotResultTree>;
  /** Derived worksheet overlay; never materialized in ordinary cells. */
  pivotProjections: Record<string, PivotGridProjection>;
  sparklines: SparklineModel[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  merges: MergeSpan[];
  freeze: FreezeModel;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  hiddenRows: number[];
  hiddenColumns: number[];
  outlineGroups: OutlineGroup[];
  outlineControls: OutlineControl[];
  filterColumns: number[];
  filterButtons: FilterButtonCell[];
  sheetTables: SheetTableModel[];
  tabColor?: string;
  hidden?: boolean;
  /** Print preview only — bounded slice */
  previewRows: PreviewRowSnapshot[];
}

function toFormulaDisplay(value: FormulaValue): string {
  if (isFormulaError(value)) return value.code;
  if (Array.isArray(value)) {
    return value.length > 0 && Array.isArray(value[0]) ? String(value[0][0]) : String(value[0]);
  }
  return value == null ? '' : String(value);
}

function formatDisplayValue(
  cell: CellData | undefined,
  formula: FormulaEngine,
  sheet: WorksheetModel,
  sheetId: string,
  row: number,
  column: number,
): string {
  if (cell?.formula) {
    return toFormulaDisplay(formula.getCellValue({ sheetId, row, column }));
  }
  const spillValue = formula.getSpillValueAt(sheetId, row, column);
  if (spillValue !== undefined) return toFormulaDisplay(spillValue);
  if (!cell) return '';
  if (cell.value == null) return '';
  if (typeof cell.value === 'number') {
    return formatNumberValue(cell.value, cell.numberFormat ?? cell.style?.numberFormat);
  }
  return String(cell.value);
}

function usedRangeOfSheet(sheet: WorksheetModel): RangeRef {
  let minRow = Number.POSITIVE_INFINITY;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxColumn = 0;
  sheet.cells.forEach((_cell, row, column) => {
    minRow = Math.min(minRow, row);
    minColumn = Math.min(minColumn, column);
    maxRow = Math.max(maxRow, row);
    maxColumn = Math.max(maxColumn, column);
  });
  for (const region of sheet.dataRegions) {
    minRow = Math.min(minRow, region.range.startRow);
    minColumn = Math.min(minColumn, region.range.startColumn);
    maxRow = Math.max(maxRow, region.range.endRow);
    maxColumn = Math.max(maxColumn, region.range.endColumn);
  }
  return {
    sheetId: sheet.id,
    startRow: Number.isFinite(minRow) ? minRow : 0,
    endRow: Number.isFinite(minRow) ? maxRow : 0,
    startColumn: Number.isFinite(minColumn) ? minColumn : 0,
    endColumn: Number.isFinite(minColumn) ? maxColumn : 0,
  };
}

function blockBackedCell(
  sheet: WorksheetModel,
  row: number,
  column: number,
  dataContent: ReadonlyMap<string, DataSourceContentQuery>,
): CellData | undefined {
  const region = sheet.dataRegions.find((candidate) => row > candidate.headerRow
    && row <= candidate.range.endRow
    && column >= candidate.range.startColumn
    && column <= candidate.range.endColumn);
  if (!region) return undefined;
  const query = dataContent.get(region.sourceId);
  if (!query) return { value: '#BLOCK!', style: { textColor: '#b91c1c' } };
  const result = query.peekCellValue(row - region.headerRow - 1, column - region.range.startColumn);
  if (result.value !== undefined) return { value: result.value };
  if (result.state.availability === 'loading') return { value: 'Loading…', style: { textColor: '#64748b', italic: true } };
  return { value: '#BLOCK!', style: { textColor: '#b91c1c', bold: true } };
}

function pivotSourceState(
  pivot: PivotModel,
  dataContent: ReadonlyMap<string, DataSourceContentQuery>,
): PivotProjectionSourceState | undefined {
  if (pivot.source.kind !== 'data-source') return undefined;
  const query = dataContent.get(pivot.source.dataSourceId);
  if (!query) return { availability: 'missing', error: `Data source ${pivot.source.dataSourceId} is unavailable` };
  const states = query.getLoadStates();
  if (states.some((state) => state.availability === 'error')) {
    const current = states.find((state) => state.availability === 'error');
    return { availability: 'error', error: current?.error ?? `Data source ${pivot.source.dataSourceId} failed to load` };
  }
  if (states.some((state) => state.availability === 'missing')) {
    const current = states.find((state) => state.availability === 'missing');
    return { availability: 'missing', error: current?.error ?? `Data source ${pivot.source.dataSourceId} is missing a block` };
  }
  if (states.some((state) => state.availability === 'loading') || states.length === 0) return { availability: 'loading' };
  return { availability: 'ready' };
}

export function buildCanvasSheetSnapshot(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  formula: FormulaEngine,
  showInvalid: boolean,
  cachedPivotResults: Readonly<Record<string, PivotResultTree>> = {},
  dataContent: ReadonlyMap<string, DataSourceContentQuery> = new Map(),
): CanvasSheetSnapshot {
  const overlays = computeConditionalOverlays(sheet);
  const filterHidden = computeFilterHiddenRows(sheet);
  const outlineHiddenRows = computeOutlineHiddenRows(sheet);
  const outlineHiddenColumns = computeOutlineHiddenColumns(sheet);
  const hiddenRows = new Set<number>([...sheet.hiddenRows, ...filterHidden, ...outlineHiddenRows]);
  const hiddenColumns = new Set<number>([...sheet.hiddenColumns, ...outlineHiddenColumns]);
  const filterColumns = resolveActiveFilterColumns(sheet);
  const filterButtons = resolveFilterButtonCells(sheet);
  const outlineControls = resolveOutlineControls(sheet);
  const viewColumns = Array.from({ length: Math.max(26, sheet.columnCount) }, (_, index) => columnLabel(index));
  const usedRange = usedRangeOfSheet(sheet);

  const getCell = (row: number, column: number): CanvasCellSnapshot | undefined => {
    if (row < 0 || row >= sheet.rowCount || column < 0 || column >= sheet.columnCount) return undefined;
    if (hiddenRows.has(row) || hiddenColumns.has(column)) return undefined;
    const modelCell = sheet.cells.get(row, column) ?? blockBackedCell(sheet, row, column, dataContent);
    const value = formatDisplayValue(modelCell, formula, sheet, sheet.id, row, column);
    const key = `${row}:${column}`;
    const overlay = overlays.get(key);
    const table = findSheetTableAt(sheet, row, column);
    const presentation = mergePresentationStyles(
      computeBandedCellStyle(sheet, row, column),
      table ? computeSheetTableCellStyle(table, row, column) : undefined,
    );
    const style = overlay?.style
      ? { ...(modelCell?.style ?? {}), ...presentation, ...overlay.style }
      : presentation
        ? { ...(modelCell?.style ?? {}), ...presentation }
        : modelCell?.style;
    const validation = validateDataInput(sheet, row, column, modelCell?.value ?? null);
    const thread = sheet.commentThreads.find((entry) => entry.row === row && entry.column === column);
    const note = getCellNote(sheet, row, column);
    const comment = thread ? threadToCellComment(thread) : undefined;
    const hyperlinkDetail = getCellHyperlink(sheet, row, column);
    const hyperlink = resolveHyperlinkDisplay(hyperlinkDetail);
    return {
      address: cellAddress(row, column),
      formula: modelCell?.formula,
      style,
      value,
      displayValue: value,
      hasComment: Boolean(comment || note),
      commentText: comment?.text ?? note?.text,
      comment,
      note: note ? structuredClone(note) : undefined,
      invalid: showInvalid && modelCell?.value != null && !validation.valid,
      hyperlink,
      overlay,
    };
  };

  const previewRows: PreviewRowSnapshot[] = [];
  const previewRowLimit = Math.min(Math.max(60, sheet.rowCount), 200);
  for (let row = 0; row < previewRowLimit; row += 1) {
    if (hiddenRows.has(row)) continue;
    const cells: Array<{ value: string }> = [];
    for (let column = 0; column < viewColumns.length; column += 1) {
      cells.push({ value: getCell(row, column)?.value ?? '' });
    }
    previewRows.push({ rowNumber: row + 1, cells, height: sheet.rowHeights[row] ?? 28 });
  }

  const pivotResults: Record<string, PivotResultTree> = {};
  const pivotProjections: Record<string, PivotGridProjection> = {};
  for (const pivot of sheet.pivots) {
    const sourceState = pivotSourceState(pivot, dataContent);
    let cachedResult = cachedPivotResults[pivot.id] ?? getLastValidPivotResult(workbook, pivot.id);
    if (!cachedResult && pivot.source.kind !== 'data-source') {
      try {
        cachedResult = computePivotResult(workbook, pivot);
      } catch {
        // The projection builder emits an explicit error state for malformed
        // synchronous definitions; it must not replace a retained result.
      }
    }
    if (cachedResult) pivotResults[pivot.id] = cachedResult;
    try {
      pivotProjections[pivot.id] = buildPivotGridProjection(workbook, pivot, cachedResult, { sourceState });
      const retained = getLastValidPivotResult(workbook, pivot.id);
      if (!pivotResults[pivot.id] && retained) pivotResults[pivot.id] = retained;
    } catch {
      // Invalid target/source is surfaced by command validation. The snapshot
      // remains renderable for the rest of the worksheet.
    }
  }

  return {
    id: sheet.id,
    name: sheet.name,
    columns: viewColumns,
    columnCount: sheet.columnCount,
    rowCount: sheet.rowCount,
    isEmpty: sheet.cells.count() === 0 && sheet.dataRegions.length === 0,
    occupiedCellCount: sheet.cells.count() + sheet.dataRegions.reduce((count, region) => count + (region.range.endRow - region.range.startRow + 1) * (region.range.endColumn - region.range.startColumn + 1), 0),
    getCell,
    usedRange,
    drawings: structuredClone(sheet.drawings),
    drawingPayloads: new Map(
      [...sheet.drawingPayloads.entries()].map(([payloadId, payload]) => [payloadId, structuredClone(payload)]),
    ),
    pivots: [...sheet.pivots],
    pivotResults,
    pivotProjections,
    sparklines: [...sheet.sparklines],
    conditionalFormats: [...sheet.conditionalFormats],
    dataValidations: [...sheet.dataValidations],
    merges: [...sheet.merges],
    freeze: { ...sheet.freeze },
    rowHeights: { ...sheet.rowHeights },
    columnWidths: { ...sheet.columnWidths },
    hiddenRows: [...hiddenRows].sort((a, b) => a - b),
    hiddenColumns: [...hiddenColumns].sort((a, b) => a - b),
    outlineGroups: sheet.outline ? structuredClone(sheet.outline.groups) : [],
    outlineControls,
    filterColumns,
    filterButtons,
    sheetTables: [...sheet.sheetTables],
    tabColor: sheet.tabColor,
    hidden: sheet.hidden,
    previewRows,
  };
}

export function buildAllSheetSnapshots(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  pivotResults: Readonly<Record<string, PivotResultTree>>,
  dataContent: ReadonlyMap<string, DataSourceContentQuery> = new Map(),
): CanvasSheetSnapshot[] {
  return workbook.getSheets().map((sheet) => buildCanvasSheetSnapshot(workbook, sheet, formula, true, pivotResults, dataContent));
}
