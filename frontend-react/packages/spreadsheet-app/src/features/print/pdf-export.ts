import type { PrintLayoutModel, PrintPageInfo } from './index';

export interface PdfExportOptions {
  filename?: string;
  title?: string;
  author?: string;
}

export interface PdfExportHook {
  /** 由宿主实现 — 浏览器可用 print-to-pdf，Node 可用 pdfkit */
  renderPages(pages: PrintPageInfo[], layout: PrintLayoutModel, options: PdfExportOptions): Promise<Blob | Buffer>;
}

export class PdfExportService {
  constructor(private readonly hook: PdfExportHook) {}

  async export(layout: PrintLayoutModel, pages: PrintPageInfo[], options: PdfExportOptions = {}): Promise<Blob | Buffer> {
    return this.hook.renderPages(pages, layout, {
      filename: options.filename ?? 'workbook.pdf',
      title: options.title,
      author: options.author,
    });
  }
}

/** 浏览器 fallback — 打开 print dialog */
export const browserPrintHook: PdfExportHook = {
  async renderPages(_pages, _layout, options) {
    if (typeof window !== 'undefined') {
      window.print();
    }
    return new Blob([JSON.stringify({ status: 'print-dialog', filename: options.filename })], { type: 'application/json' });
  },
};
