import { normalizePrintDocumentSnapshot, type PrintDocumentSnapshot, type RangeRef, type SheetId, type WorkbookModel } from '@react-sheets/core-model';

export type PaperSize = PrintDocumentSnapshot['pageSetup']['paperSize'];
export type PageOrientation = PrintDocumentSnapshot['pageSetup']['orientation'];

/** Public print input used by UI and host callers. */
export interface PrintLayout {
  paper: 'A3' | 'A4' | 'Letter' | 'Legal';
  orientation: PageOrientation;
  margin: { top: number; right: number; bottom: number; left: number };
  repeatRows?: RangeRef;
  repeatColumns?: RangeRef;
  scale?: number;
  fitToWidth?: boolean;
  fitToHeight?: boolean;
  printGridlines?: boolean;
  printHeadings?: boolean;
  centerHorizontally?: boolean;
  centerVertically?: boolean;
  headerText?: string;
  footerText?: string;
}

export type PageSetup = PrintDocumentSnapshot['pageSetup'];

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

export type PrintArea = PrintDocumentSnapshot['printAreas'][number];
export type PrintPageBreak = PrintDocumentSnapshot['pageBreaks'][number];
export interface PrintTitleSpan {
  start: number;
  end: number;
}

/** Canonical, serializable workbook print state. */
/** Canonical workbook-owned print state. */
export type PrintDocument = PrintDocumentSnapshot;

export interface PrintLayoutModel {
  unitId: string;
  pageSetup: PageSetup;
  printAreas: PrintArea[];
  pageBreaks: PrintPageBreak[];
  repeatRows?: { start: number; end: number };
  repeatColumns?: { start: number; end: number };
  rowHeights?: Readonly<Record<number, number>>;
  columnWidths?: Readonly<Record<number, number>>;
  hiddenRows?: ReadonlySet<number>;
  hiddenColumns?: ReadonlySet<number>;
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

export function createDefaultPrintDocument(unitId: string, sheetId: SheetId): PrintDocument {
  return {
    schema: 'PrintDocument',
    unitId,
    sheetId,
    pageSetup: structuredClone(DEFAULT_PAGE_SETUP),
    // An empty area means "used range". A persisted explicit area is always
    // represented by a non-empty list, so the default never hard-codes a
    // worksheet size into the print document.
    printAreas: [],
    pageBreaks: [],
  };
}

function assertDocumentTarget(workbook: WorkbookModel, document: PrintDocument): void {
  if (document.unitId !== workbook.unitId) {
    throw new Error(`Print document unit mismatch: expected ${workbook.unitId}, received ${document.unitId}`);
  }
  workbook.getSheet(document.sheetId);
}

export function getPrintDocument(workbook: WorkbookModel, sheetId: SheetId): PrintDocument {
  workbook.getSheet(sheetId);
  const current = workbook.getPrintDocument(sheetId);
  if (current) return structuredClone(current);
  return createDefaultPrintDocument(workbook.unitId, sheetId);
}

export function replacePrintDocument(workbook: WorkbookModel, document: PrintDocument): void {
  assertDocumentTarget(workbook, document);
  workbook.setPrintDocument(normalizePrintDocument(document));
}

export function serializePrintDocuments(workbook: WorkbookModel): PrintDocument[] {
  return workbook.getSheets().map((sheet) => structuredClone(
    workbook.getPrintDocument(sheet.id) ?? createDefaultPrintDocument(workbook.unitId, sheet.id),
  ));
}

export function hydratePrintDocuments(workbook: WorkbookModel, documents: readonly PrintDocument[]): void {
  workbook.clearPrintDocuments();
  for (const document of documents) replacePrintDocument(workbook, document);
}

export function normalizePrintDocument(document: PrintDocument): PrintDocument {
  return normalizePrintDocumentSnapshot(document);
}

export interface PrintPageInfo {
  pageIndex: number;
  sheetId: SheetId;
  range: RangeRef;
  widthPx: number;
  heightPx: number;
  repeatRows?: { start: number; end: number };
  repeatColumns?: { start: number; end: number };
}

export interface PrintPaginationOptions {
  rowHeights?: Readonly<Record<number, number>>;
  columnWidths?: Readonly<Record<number, number>>;
  hiddenRows?: ReadonlySet<number>;
  hiddenColumns?: ReadonlySet<number>;
}

const PAPER_POINTS: Record<PaperSize, { width: number; height: number }> = {
  a3: { width: 841.89, height: 1190.55 },
  a4: { width: 595.28, height: 841.89 },
  letter: { width: 612, height: 792 },
  legal: { width: 612, height: 1008 },
  custom: { width: 595.28, height: 841.89 },
};

function pageCapacity(layout: PrintLayoutModel): { width: number; height: number } {
  const paper = PAPER_POINTS[layout.pageSetup.paperSize];
  const landscape = layout.pageSetup.orientation === 'landscape';
  const width = landscape ? paper.height : paper.width;
  const height = landscape ? paper.width : paper.height;
  return {
    width: Math.max(1, width - layout.pageSetup.margins.left - layout.pageSetup.margins.right),
    height: Math.max(1, height - layout.pageSetup.margins.top - layout.pageSetup.margins.bottom - layout.pageSetup.margins.header - layout.pageSetup.margins.footer),
  };
}

function splitByBreaks(start: number, end: number, breaks: number[]): Array<{ start: number; end: number }> {
  const boundaries = [start, ...breaks.filter((value) => value > start && value <= end).sort((a, b) => a - b), end + 1];
  const parts: Array<{ start: number; end: number }> = [];
  for (let index = 0; index + 1 < boundaries.length; index += 1) {
    const partStart = boundaries[index]!;
    const partEnd = boundaries[index + 1]! - 1;
    if (partStart <= partEnd) parts.push({ start: partStart, end: partEnd });
  }
  return parts;
}

function chunkByCapacity(
  segment: { start: number; end: number },
  capacity: number,
  sizeOf: (index: number) => number,
  fit: boolean,
): Array<{ start: number; end: number }> {
  if (fit) return [segment];
  const parts: Array<{ start: number; end: number }> = [];
  let start = segment.start;
  let total = 0;
  for (let index = segment.start; index <= segment.end; index += 1) {
    const size = Math.max(0.01, sizeOf(index));
    if (index > start && total + size > capacity) {
      parts.push({ start, end: index - 1 });
      start = index;
      total = 0;
    }
    total += size;
  }
  if (start <= segment.end) parts.push({ start, end: segment.end });
  return parts;
}

function trimHidden(segment: { start: number; end: number }, hidden: ReadonlySet<number>): { start: number; end: number } | undefined {
  let start = segment.start;
  let end = segment.end;
  while (start <= end && hidden.has(start)) start += 1;
  while (end >= start && hidden.has(end)) end -= 1;
  return start <= end ? { start, end } : undefined;
}

/** One pagination implementation shared by browser and Node print hosts. */
export function computePrintPages(layout: PrintLayoutModel, rowHeight = 20, colWidth = 80, options: PrintPaginationOptions = {}): PrintPageInfo[] {
  const pages: PrintPageInfo[] = [];
  const capacity = pageCapacity(layout);
  const scale = Math.max(0.01, layout.pageSetup.scale / 100);
  const rows = options.rowHeights ?? layout.rowHeights ?? {};
  const columns = options.columnWidths ?? layout.columnWidths ?? {};
  const hiddenRows = options.hiddenRows ?? layout.hiddenRows ?? new Set<number>();
  const hiddenColumns = options.hiddenColumns ?? layout.hiddenColumns ?? new Set<number>();
  const rowSize = (row: number) => (rows[row] ?? rowHeight) * 0.75 * scale;
  const columnSize = (column: number) => (columns[column] ?? colWidth) * 0.75 * scale;

  for (const area of layout.printAreas) {
    const { range } = area;
    const rowBreaks = layout.pageBreaks.filter((item) => item.sheetId === area.sheetId && item.row !== undefined).map((item) => item.row!);
    const columnBreaks = layout.pageBreaks.filter((item) => item.sheetId === area.sheetId && item.column !== undefined).map((item) => item.column!);
    const rowSegments = splitByBreaks(range.startRow, range.endRow, rowBreaks)
      .map((part) => trimHidden(part, hiddenRows))
      .filter((part): part is { start: number; end: number } => Boolean(part));
    const columnSegments = splitByBreaks(range.startColumn, range.endColumn, columnBreaks)
      .map((part) => trimHidden(part, hiddenColumns))
      .filter((part): part is { start: number; end: number } => Boolean(part));
    for (const rowSegment of rowSegments) {
      const rowPages = chunkByCapacity(rowSegment, capacity.height, rowSize, Boolean(layout.pageSetup.fitToHeight));
      for (const columnSegment of columnSegments) {
        const columnPages = chunkByCapacity(columnSegment, capacity.width, columnSize, Boolean(layout.pageSetup.fitToWidth));
        for (const rowPage of rowPages) for (const columnPage of columnPages) {
          let widthPx = 0;
          let heightPx = 0;
          for (let column = columnPage.start; column <= columnPage.end; column += 1) if (!hiddenColumns.has(column)) widthPx += columns[column] ?? colWidth;
          for (let row = rowPage.start; row <= rowPage.end; row += 1) if (!hiddenRows.has(row)) heightPx += rows[row] ?? rowHeight;
          pages.push({
            pageIndex: pages.length,
            sheetId: area.sheetId,
            range: { sheetId: area.sheetId, startRow: rowPage.start, endRow: rowPage.end, startColumn: columnPage.start, endColumn: columnPage.end },
            widthPx,
            heightPx,
            repeatRows: layout.repeatRows,
            repeatColumns: layout.repeatColumns,
          });
        }
      }
    }
  }
  return pages;
}

export * from './pdf-export';
export * from './commands';
export * from './layout';
