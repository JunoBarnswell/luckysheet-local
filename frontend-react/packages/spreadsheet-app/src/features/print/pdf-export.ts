import type { PrintLayoutModel, PrintPageInfo } from './index';

export interface PdfExportOptions {
  filename?: string;
  title?: string;
  author?: string;
  /** Optional page content supplied by the workbook renderer. */
  pageText?: (page: PrintPageInfo, pageNumber: number) => readonly string[];
}

export type PdfExportResult = Uint8Array | Blob;

/** Host-owned output port. It never invokes window.print or returns JSON. */
export interface PdfExportHook {
  readonly kind: 'browser' | 'node';
  renderPages(pages: PrintPageInfo[], layout: PrintLayoutModel, options: PdfExportOptions): Promise<PdfExportResult>;
}

export class PdfExportService {
  constructor(private readonly hook: PdfExportHook = browserPrintHook) {}

  async export(layout: PrintLayoutModel, pages: PrintPageInfo[], options: PdfExportOptions = {}): Promise<PdfExportResult> {
    return this.hook.renderPages(pages, layout, {
      filename: options.filename ?? 'workbook.pdf',
      title: options.title,
      author: options.author,
      pageText: options.pageText,
    });
  }
}

const PAPER_POINTS: Record<PrintLayoutModel['pageSetup']['paperSize'], { width: number; height: number }> = {
  a3: { width: 841.89, height: 1190.55 },
  a4: { width: 595.28, height: 841.89 },
  letter: { width: 612, height: 792 },
  legal: { width: 612, height: 1008 },
  custom: { width: 595.28, height: 841.89 },
};

function ascii(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length);
  for (let index = 0; index < value.length; index += 1) bytes[index] = value.charCodeAt(index) & 0xff;
  return bytes;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const size = parts.reduce((total, part) => total + part.length, 0);
  const result = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

function pdfText(value: string): string {
  let result = '';
  for (const character of value) {
    const code = character.charCodeAt(0);
    const safe = code >= 32 && code <= 126 ? character : '?';
    result += safe === '(' || safe === ')' || safe === '\\' ? `\\${safe}` : safe;
  }
  return result;
}

function pageSize(layout: PrintLayoutModel): { width: number; height: number } {
  const paper = PAPER_POINTS[layout.pageSetup.paperSize];
  return layout.pageSetup.orientation === 'landscape'
    ? { width: paper.height, height: paper.width }
    : paper;
}

function pageContent(page: PrintPageInfo, pageNumber: number, pageTotal: number, layout: PrintLayoutModel, options: PdfExportOptions): string {
  const size = pageSize(layout);
  const lines = [
    options.title ?? 'Workbook',
    `Page ${pageNumber} of ${pageTotal}`,
    `Sheet ${page.sheetId}`,
    `Rows ${page.range.startRow + 1}-${page.range.endRow + 1}, columns ${page.range.startColumn + 1}-${page.range.endColumn + 1}`,
    ...(options.pageText?.(page, pageNumber) ?? []),
  ];
  const left = Math.max(12, layout.pageSetup.margins.left);
  const top = Math.max(12, layout.pageSetup.margins.top);
  const lineHeight = 14;
  let content = 'q\n';
  content += `0 0 ${size.width.toFixed(2)} ${size.height.toFixed(2)} re S\n`;
  content += 'BT\n/F1 11 Tf\n';
  for (let index = 0; index < lines.length; index += 1) {
    const y = size.height - top - index * lineHeight;
    content += `1 0 0 1 ${left.toFixed(2)} ${y.toFixed(2)} Tm\n(${pdfText(String(lines[index] ?? ''))}) Tj\n`;
  }
  if (layout.pageSetup.footerText) {
    content += `1 0 0 1 ${left.toFixed(2)} ${Math.max(12, layout.pageSetup.margins.bottom / 2).toFixed(2)} Tm\n(${pdfText(layout.pageSetup.footerText)}) Tj\n`;
  }
  content += 'ET\nQ\n';
  return content;
}

/**
 * Deterministic, dependency-free PDF 1.7 writer. The bytes are valid PDF
 * objects and can be opened by browser viewers, Preview, Acrobat, and
 * Poppler. Cell painting is supplied by the optional pageText port; the
 * pagination and document metadata remain identical in browser and Node.
 */
export function createPdfBytes(pages: readonly PrintPageInfo[], layout: PrintLayoutModel, options: PdfExportOptions = {}): Uint8Array {
  const objects: string[] = [];
  const pageObjectIds: number[] = [];
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
  let nextObjectId = 4;
  for (let index = 0; index < pages.length; index += 1) {
    const pageObjectId = nextObjectId++;
    const contentObjectId = nextObjectId++;
    pageObjectIds.push(pageObjectId);
    const content = pageContent(pages[index]!, index + 1, pages.length, layout, options);
    objects[pageObjectId] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageSize(layout).width.toFixed(2)} ${pageSize(layout).height.toFixed(2)}] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentObjectId} 0 R >>`;
    objects[contentObjectId] = `<< /Length ${content.length} >>\nstream\n${content}endstream`;
  }
  objects[2] = `<< /Type /Pages /Kids [${pageObjectIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageObjectIds.length} >>`;
  const infoObjectId = nextObjectId++;
  objects[infoObjectId] = `<< /Title (${pdfText(options.title ?? 'Workbook')}) /Author (${pdfText(options.author ?? 'Spreadsheet')}) >>`;

  const chunks: Uint8Array[] = [new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37, 0x0a, 0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a])];
  const offsets = new Array<number>(nextObjectId).fill(0);
  let length = chunks[0]!.length;
  for (let objectId = 1; objectId < nextObjectId; objectId += 1) {
    const object = objects[objectId];
    if (!object) continue;
    const bytes = ascii(`${objectId} 0 obj\n${object}\nendobj\n`);
    offsets[objectId] = length;
    chunks.push(bytes);
    length += bytes.length;
  }
  const xrefOffset = length;
  let xref = `xref\n0 ${nextObjectId}\n0000000000 65535 f \n`;
  for (let objectId = 1; objectId < nextObjectId; objectId += 1) xref += `${String(offsets[objectId]).padStart(10, '0')} 00000 n \n`;
  xref += `trailer\n<< /Size ${nextObjectId} /Root 1 0 R /Info ${infoObjectId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;
  chunks.push(ascii(xref));
  return concatBytes(chunks);
}

export const nodePrintHook: PdfExportHook = {
  kind: 'node',
  async renderPages(pages, layout, options) {
    return createPdfBytes(pages, layout, options);
  },
};

export const browserPrintHook: PdfExportHook = {
  kind: 'browser',
  async renderPages(pages, layout, options) {
    const bytes = createPdfBytes(pages, layout, options);
    if (typeof Blob === 'undefined') return bytes;
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    return new Blob([buffer], { type: 'application/pdf' });
  },
};
