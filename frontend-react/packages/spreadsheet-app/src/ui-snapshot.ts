import type {
  CellComment,
  CellData,
  CellStyle,
  ChartModel,
  ConditionalFormatRule,
  DataValidationRule,
  FreezeModel,
  MergeSpan,
  PivotModel,
  PivotResultTree,
  RangeRef,
  ShapeModel,
  SparklineModel,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { computeConditionalOverlays, computeFilterHiddenRows, validateDataInput, type ConditionalOverlay } from '@react-sheets/sheet-features';
import { FormulaEngine, isFormulaError, type FormulaValue } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import { computePivotResult } from '@react-sheets/pro-features';
import { cellAddress, columnLabel } from './address';

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
  charts: ChartModel[];
  pivots: PivotModel[];
  pivotResults: Record<string, PivotResultTree>;
  shapes: ShapeModel[];
  sparklines: SparklineModel[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  merges: MergeSpan[];
  freeze: FreezeModel;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  hiddenRows: number[];
  filterColumns: number[];
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
  sheetId: string,
  row: number,
  column: number,
): string {
  if (!cell) return '';
  if (cell.formula) {
    return toFormulaDisplay(formula.getCellValue({ sheetId, row, column }));
  }
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
  return {
    sheetId: sheet.id,
    startRow: Number.isFinite(minRow) ? minRow : 0,
    endRow: Number.isFinite(minRow) ? maxRow : 0,
    startColumn: Number.isFinite(minColumn) ? minColumn : 0,
    endColumn: Number.isFinite(minColumn) ? maxColumn : 0,
  };
}

export function buildCanvasSheetSnapshot(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  formula: FormulaEngine,
  showInvalid: boolean,
  cachedPivotResults: Readonly<Record<string, PivotResultTree>> = {},
): CanvasSheetSnapshot {
  const overlays = computeConditionalOverlays(sheet);
  const filterHidden = computeFilterHiddenRows(sheet);
  const hiddenRows = new Set<number>([...sheet.hiddenRows, ...filterHidden]);
  const filterColumns = sheet.filter ? Object.keys(sheet.filter.criteria).map(Number) : [];
  const viewColumns = Array.from({ length: Math.max(26, sheet.columnCount) }, (_, index) => columnLabel(index));
  const usedRange = usedRangeOfSheet(sheet);

  const getCell = (row: number, column: number): CanvasCellSnapshot | undefined => {
    if (row < 0 || row >= sheet.rowCount || column < 0 || column >= sheet.columnCount) return undefined;
    const modelCell = sheet.cells.get(row, column);
    const value = formatDisplayValue(modelCell, formula, sheet.id, row, column);
    const key = `${row}:${column}`;
    const overlay = overlays.get(key);
    const style = overlay?.style ? { ...(modelCell?.style ?? {}), ...overlay.style } : modelCell?.style;
    const validation = validateDataInput(sheet, row, column, modelCell?.value ?? null);
    return {
      address: cellAddress(row, column),
      formula: modelCell?.formula,
      style,
      value,
      displayValue: value,
      hasComment: Boolean(modelCell?.comment),
      commentText: modelCell?.comment?.text,
      comment: modelCell?.comment ? structuredClone(modelCell.comment) : undefined,
      invalid: showInvalid && modelCell?.value != null && !validation.valid,
      hyperlink: modelCell?.hyperlink,
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
  for (const pivot of sheet.pivots) {
    try {
      pivotResults[pivot.id] = cachedPivotResults[pivot.id] ?? computePivotResult(workbook, pivot);
    } catch {
      // invalid pivot must not block sheet render
    }
  }

  return {
    id: sheet.id,
    name: sheet.name,
    columns: viewColumns,
    columnCount: sheet.columnCount,
    rowCount: sheet.rowCount,
    isEmpty: sheet.cells.count() === 0,
    occupiedCellCount: sheet.cells.count(),
    getCell,
    usedRange,
    charts: [...sheet.charts],
    pivots: [...sheet.pivots],
    pivotResults,
    shapes: [...sheet.shapes],
    sparklines: [...sheet.sparklines],
    conditionalFormats: [...sheet.conditionalFormats],
    dataValidations: [...sheet.dataValidations],
    merges: [...sheet.merges],
    freeze: { ...sheet.freeze },
    rowHeights: { ...sheet.rowHeights },
    columnWidths: { ...sheet.columnWidths },
    hiddenRows: [...hiddenRows].sort((a, b) => a - b),
    filterColumns,
    tabColor: sheet.tabColor,
    hidden: sheet.hidden,
    previewRows,
  };
}

export function buildAllSheetSnapshots(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  pivotResults: Readonly<Record<string, PivotResultTree>>,
): CanvasSheetSnapshot[] {
  return workbook.getSheets().map((sheet) => buildCanvasSheetSnapshot(workbook, sheet, formula, true, pivotResults));
}
