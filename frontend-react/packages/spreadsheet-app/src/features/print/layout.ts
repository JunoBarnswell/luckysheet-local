import type { RangeRef, SheetId, WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import type { PrintLayout } from '@react-sheets/pro-features';
import { usedRangeOfSheet } from '../../application-helpers';
import {
  computePrintPages,
  createDefaultPrintLayout,
  type PageSetup,
  type PaperSize,
  type PrintLayoutModel,
  type PrintPageInfo,
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

function mmToPt(mm: number): number {
  return Math.round(mm * MM_TO_PT);
}

function mapPaperSize(paper: PrintLayout['paper']): PaperSize {
  switch (paper) {
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
    printGridlines: false,
    printHeadings: false,
    centerHorizontally: false,
    centerVertically: false,
  };
}

export function pageSetupToPrintLayout(setup: PageSetup): PrintLayout {
  const ptToMm = (pt: number) => Math.round((pt / MM_TO_PT) * 10) / 10;
  const paper: PrintLayout['paper'] =
    setup.paperSize === 'letter' ? 'Letter' : setup.paperSize === 'legal' ? 'Legal' : 'A4';
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
    total += sheet.rowHeights[row] ?? 28;
    count += 1;
  }
  return count > 0 ? total / count : 28;
}

function averageColumnWidth(sheet: WorksheetModel, range: RangeRef): number {
  let total = 0;
  let count = 0;
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    total += sheet.columnWidths[column] ?? 80;
    count += 1;
  }
  return count > 0 ? total / count : 80;
}

export function buildPrintLayoutModel(
  unitId: string,
  sheetId: SheetId,
  uiLayout: PrintLayout,
  printArea: RangeRef,
): PrintLayoutModel {
  const base = createDefaultPrintLayout(unitId, sheetId);
  return {
    ...base,
    pageSetup: printLayoutToPageSetup(uiLayout),
    printAreas: [{ sheetId, range: printArea }],
    repeatRows: uiLayout.repeatRows
      ? { start: uiLayout.repeatRows.startRow, end: uiLayout.repeatRows.endRow }
      : undefined,
  };
}

function computePagesWithFit(
  model: PrintLayoutModel,
  uiLayout: PrintLayout,
  rowHeight: number,
  colWidth: number,
): PrintPageInfo[] {
  const area = model.printAreas[0];
  if (!area) return [];
  const { range } = area;
  const totalRows = range.endRow - range.startRow + 1;
  const totalCols = range.endColumn - range.startColumn + 1;
  const baseRowsPerPage = model.pageSetup.orientation === 'portrait' ? 40 : 28;
  const baseColsPerPage = model.pageSetup.orientation === 'portrait' ? 8 : 12;
  const rowsPerPage = uiLayout.fitToHeight ? totalRows : baseRowsPerPage;
  const colsPerPage = uiLayout.fitToWidth ? totalCols : baseColsPerPage;

  if (!uiLayout.fitToWidth && !uiLayout.fitToHeight) {
    return computePrintPages(model, rowHeight, colWidth);
  }

  const pages: PrintPageInfo[] = [];
  const rowPages = Math.max(1, Math.ceil(totalRows / rowsPerPage));
  const colPages = Math.max(1, Math.ceil(totalCols / colsPerPage));
  let pageIndex = 0;

  for (let rowPage = 0; rowPage < rowPages; rowPage += 1) {
    for (let colPage = 0; colPage < colPages; colPage += 1) {
      const startRow = range.startRow + rowPage * rowsPerPage;
      const endRow = Math.min(range.endRow, startRow + rowsPerPage - 1);
      const startColumn = range.startColumn + colPage * colsPerPage;
      const endColumn = Math.min(range.endColumn, startColumn + colsPerPage - 1);
      pages.push({
        pageIndex: pageIndex++,
        sheetId: area.sheetId,
        range: { sheetId: area.sheetId, startRow, endRow, startColumn, endColumn },
        widthPx: (endColumn - startColumn + 1) * colWidth,
        heightPx: (endRow - startRow + 1) * rowHeight,
      });
    }
  }
  return pages;
}

export function toPrintPageSnapshots(pages: PrintPageInfo[]): PrintPageSnapshot[] {
  return pages.map((page, index) => ({
    page: index + 1,
    sheetId: page.sheetId,
    range: page.range,
  }));
}

export function buildPrintSnapshot(
  workbook: WorkbookModel,
  activeSheetId: SheetId,
  uiLayout: PrintLayout,
  selectionRange?: RangeRef,
): PrintSnapshot {
  const sheet = workbook.getSheet(activeSheetId);
  const printArea = resolvePrintArea(sheet, selectionRange);
  const model = buildPrintLayoutModel(workbook.unitId, activeSheetId, uiLayout, printArea);
  const rowHeight = averageRowHeight(sheet, printArea);
  const colWidth = averageColumnWidth(sheet, printArea);
  const pages = computePagesWithFit(model, uiLayout, rowHeight, colWidth);
  const pageSnapshots = toPrintPageSnapshots(pages);
  return {
    layout: uiLayout,
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
