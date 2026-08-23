import type { RangeRef, SheetId } from '@react-sheets/core-model';

export type PaperSize = 'letter' | 'a4' | 'a3' | 'legal' | 'custom';
export type PageOrientation = 'portrait' | 'landscape';

export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
}

export interface PageSetup {
  paperSize: PaperSize;
  orientation: PageOrientation;
  margins: PageMargins;
  scale: number;
  fitToWidth?: number;
  fitToHeight?: number;
  printGridlines: boolean;
  printHeadings: boolean;
  centerHorizontally: boolean;
  centerVertically: boolean;
  headerText?: string;
  footerText?: string;
}

export const DEFAULT_PAGE_SETUP: PageSetup = {
  paperSize: 'a4',
  orientation: 'portrait',
  margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36 },
  scale: 100,
  printGridlines: false,
  printHeadings: false,
  centerHorizontally: false,
  centerVertically: false,
};

export interface PrintArea {
  sheetId: SheetId;
  range: RangeRef;
}

export interface PrintPageBreak {
  sheetId: SheetId;
  row?: number;
  column?: number;
}

export interface PrintLayoutModel {
  unitId: string;
  pageSetup: PageSetup;
  printAreas: PrintArea[];
  pageBreaks: PrintPageBreak[];
  repeatRows?: { start: number; end: number };
  repeatColumns?: { start: number; end: number };
}

export function createDefaultPrintLayout(unitId: string, sheetId: SheetId): PrintLayoutModel {
  return {
    unitId,
    pageSetup: { ...DEFAULT_PAGE_SETUP, margins: { ...DEFAULT_PAGE_SETUP.margins } },
    printAreas: [{
      sheetId,
      range: { sheetId, startRow: 0, endRow: 999, startColumn: 0, endColumn: 25 },
    }],
    pageBreaks: [],
  };
}

export interface PrintPageInfo {
  pageIndex: number;
  sheetId: SheetId;
  range: RangeRef;
  widthPx: number;
  heightPx: number;
}

/** 独立 print layout pipeline — 分页计算 */
export function computePrintPages(layout: PrintLayoutModel, rowHeight = 20, colWidth = 80): PrintPageInfo[] {
  const pages: PrintPageInfo[] = [];
  const rowsPerPage = layout.pageSetup.orientation === 'portrait' ? 40 : 28;
  const colsPerPage = layout.pageSetup.orientation === 'portrait' ? 8 : 12;

  for (const area of layout.printAreas) {
    const { range } = area;
    const totalRows = range.endRow - range.startRow + 1;
    const totalCols = range.endColumn - range.startColumn + 1;
    const rowPages = Math.ceil(totalRows / rowsPerPage);
    const colPages = Math.ceil(totalCols / colsPerPage);
    let pageIndex = 0;

    for (let rp = 0; rp < rowPages; rp++) {
      for (let cp = 0; cp < colPages; cp++) {
        const startRow = range.startRow + rp * rowsPerPage;
        const endRow = Math.min(range.endRow, startRow + rowsPerPage - 1);
        const startColumn = range.startColumn + cp * colsPerPage;
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
  }
  return pages;
}

export * from './pdf-export';
export { registerPrintCommands } from './commands';
