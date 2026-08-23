import type { RangeRef, SheetId, WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
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

export function buildPrintSnapshot(
  workbook: WorkbookModel,
  activeSheetId: SheetId,
  uiLayout?: PrintLayout,
  selectionRange?: RangeRef,
): PrintSnapshot {
  const sheet = workbook.getSheet(activeSheetId);
  const document = getPrintDocument(workbook, activeSheetId);
  const effectiveLayout = uiLayout ?? pageSetupToPrintLayout(document.pageSetup);
  const storedArea = document.printAreas.find((area) => area.sheetId === activeSheetId)?.range;
  const printArea = selectionRange ? resolvePrintArea(sheet, selectionRange) : storedArea ?? resolvePrintArea(sheet);
  const model = buildPrintLayoutModel(workbook.unitId, activeSheetId, effectiveLayout, printArea, document);
  model.rowHeights = { ...sheet.rowHeights };
  model.columnWidths = { ...sheet.columnWidths };
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
