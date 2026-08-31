import { zlibSync } from 'fflate';
import type { PrintLayoutModel, PrintPageInfo, PrintProjection } from './index';
import { decodePngToRgb, DEFAULT_PNG_RESOURCE_LIMITS, PngDecodeError } from './png-decoder';

export interface PdfExportOptions {
  filename?: string;
  title?: string;
  author?: string;
  /** Optional page content supplied by the workbook renderer. */
  pageText?: (page: PrintPageInfo, pageNumber: number) => readonly string[];
  /** Canonical page projection; when present, cells and objects are painted. */
  pageProjection?: (page: PrintPageInfo, pageNumber: number) => PrintProjection;
}

export type PdfExportResult = Uint8Array | Blob;

/** Resource budget shared by every image format accepted by the PDF writer. */
export const PDF_IMAGE_RESOURCE_LIMITS = DEFAULT_PNG_RESOURCE_LIMITS;

export type PdfResourceErrorCode =
  | 'NATIVE_PRINT_RESOURCE_UNAVAILABLE'
  | 'NATIVE_PRINT_RESOURCE_UNSUPPORTED_FORMAT'
  | 'NATIVE_PRINT_RESOURCE_INVALID'
  | 'NATIVE_PRINT_RESOURCE_LIMIT'
  | 'NATIVE_PRINT_RESOURCE_DIMENSIONS_MISMATCH'
  | 'NATIVE_PRINT_RESOURCE_DECODE_FAILED'
  | 'NATIVE_PRINT_RESOURCE_CONFLICT';

/** Typed, observable failure for a resource that cannot be embedded safely. */
export class PdfResourceError extends Error {
  readonly code: PdfResourceErrorCode;

  constructor(code: PdfResourceErrorCode, message: string, cause?: unknown) {
    super(`${code}: ${message}`, cause === undefined ? undefined : { cause });
    this.name = 'PdfResourceError';
    this.code = code;
  }
}

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
      pageProjection: options.pageProjection,
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
  // Encode every text string as UTF-16BE with a BOM.  This keeps Unicode
  // code points in the PDF byte stream instead of silently replacing them
  // with '?' (the previous WinAnsi literal-string path was lossy).  The
  // selected PDF font remains host-resolved, while the document payload is
  // deterministic and Unicode-preserving for browser and Node output.
  let hex = 'FEFF';
  for (let index = 0; index < value.length; index += 1) hex += value.charCodeAt(index).toString(16).padStart(4, '0').toUpperCase();
  return `<${hex}>`;
}

function pageSize(layout: PrintLayoutModel): { width: number; height: number } {
  const paper = PAPER_POINTS[layout.pageSetup.paperSize];
  return layout.pageSetup.orientation === 'landscape'
    ? { width: paper.height, height: paper.width }
    : paper;
}

function pdfRgb(value: string | undefined, fallback = '#ffffff'): [number, number, number] {
  const normalized = value && /^#[0-9a-f]{6}$/i.test(value) ? value : fallback;
  return [Number.parseInt(normalized.slice(1, 3), 16) / 255, Number.parseInt(normalized.slice(3, 5), 16) / 255, Number.parseInt(normalized.slice(5, 7), 16) / 255];
}

function formatRgb(value: string | undefined, fallback = '#ffffff'): string {
  return pdfRgb(value, fallback).map((entry) => entry.toFixed(4)).join(' ');
}

interface PdfImageResource {
  name: string;
  asset: { contentHash: string; mimeType: string; width: number; height: number };
  bytes: Uint8Array;
  filter: 'DCTDecode' | 'FlateDecode';
  objectId?: number;
}

interface PdfWriterState {
  images: Map<string, PdfImageResource>;
}

function pageContent(page: PrintPageInfo, pageNumber: number, pageTotal: number, layout: PrintLayoutModel, options: PdfExportOptions, state: PdfWriterState): string {
  const size = pageSize(layout);
  const projection = options.pageProjection?.(page, pageNumber);
  const lines = [
    ...(layout.pageSetup.headerText ? [layout.pageSetup.headerText] : []),
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
    content += `1 0 0 1 ${left.toFixed(2)} ${y.toFixed(2)} Tm\n${pdfText(String(lines[index] ?? ''))} Tj\n`;
  }
  if (layout.pageSetup.footerText) {
    content += `1 0 0 1 ${left.toFixed(2)} ${Math.max(12, layout.pageSetup.margins.bottom / 2).toFixed(2)} Tm\n${pdfText(layout.pageSetup.footerText)} Tj\n`;
  }
  content += 'ET\nQ\n';
  if (projection) content += projectionContent(projection, layout, size, state, lines.length * lineHeight + 8);
  return content;
}

function projectionContent(projection: PrintProjection, layout: PrintLayoutModel, size: { width: number; height: number }, state: PdfWriterState, metadataOffset: number): string {
  const pageLeft = Math.max(12, layout.pageSetup.margins.left);
  const pageTop = Math.max(12, layout.pageSetup.margins.top) + metadataOffset;
  const pageRight = Math.max(12, layout.pageSetup.margins.right);
  const pageBottom = Math.max(12, layout.pageSetup.margins.bottom);
  const contentWidth = projection.contentWidthPx * 0.75 * projection.scaleX;
  const contentHeight = projection.contentHeightPx * 0.75 * projection.scaleY;
  const availableWidth = Math.max(1, size.width - pageLeft - pageRight);
  const availableHeight = Math.max(1, size.height - pageTop - pageBottom);
  const originX = pageLeft + (layout.pageSetup.centerHorizontally ? Math.max(0, (availableWidth - contentWidth) / 2) : 0);
  const originTop = pageTop + (layout.pageSetup.centerVertically ? Math.max(0, (availableHeight - contentHeight) / 2) : 0);
  let content = 'q\n';
  for (const cell of projection.cells) {
    const x = originX + cell.columnOffsetPx * 0.75 * projection.scaleX;
    const width = cell.widthPx * 0.75 * projection.scaleX;
    const height = cell.heightPx * 0.75 * projection.scaleY;
    const top = originTop + cell.rowOffsetPx * 0.75 * projection.scaleY;
    const y = size.height - top - height;
    const background = cell.style?.background ?? cell.style?.fill?.foreground;
    if (background) content += `${formatRgb(background)} rg ${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re f\n`;
    if (cell.image) {
      const image = registerPdfImage(state, cell.image);
      content += `q ${width.toFixed(2)} 0 0 ${height.toFixed(2)} ${x.toFixed(2)} ${y.toFixed(2)} cm /${image.name} Do Q\n`;
    }
    content += `0.82 0.84 0.88 RG 0.35 w ${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re S\n`;
    if (cell.displayValue.length > 0) {
      const [red, green, blue] = pdfRgb(cell.style?.textColor, '#1f2937');
      const fontSize = Math.max(5, (cell.style?.fontSizePx ?? 13) * 0.75 * Math.min(projection.scaleX, projection.scaleY));
      const baseline = y + Math.max(fontSize + 1, height - 3);
      content += `BT\n${red.toFixed(4)} ${green.toFixed(4)} ${blue.toFixed(4)} rg\n/F1 ${fontSize.toFixed(2)} Tf\n1 0 0 1 ${(x + 3).toFixed(2)} ${baseline.toFixed(2)} Tm\n${pdfText(cell.displayValue)} Tj\nET\n`;
    }
  }
  for (const entry of projection.drawings) {
    const x = originX + entry.xPx * 0.75 * projection.scaleX;
    const width = entry.widthPx * 0.75 * projection.scaleX;
    const height = entry.heightPx * 0.75 * projection.scaleY;
    const top = originTop + entry.yPx * 0.75 * projection.scaleY;
    const y = size.height - top - height;
    const payload = entry.payload;
    if (payload.kind === 'connector') {
      const points = payload.route.points;
      if (points.length >= 2) {
        const first = points[0]!;
        content += `0.20 0.35 0.65 RG ${(payload.strokeWidth ?? 1.5).toFixed(2)} w ${(originX + (entry.xPx + first.x - entry.drawing.transform.x) * 0.75 * projection.scaleX).toFixed(2)} ${(size.height - originTop - (entry.yPx + first.y - entry.drawing.transform.y) * 0.75 * projection.scaleY).toFixed(2)} m\n`;
        for (const point of points.slice(1)) content += `${(originX + (entry.xPx + point.x - entry.drawing.transform.x) * 0.75 * projection.scaleX).toFixed(2)} ${(size.height - originTop - (entry.yPx + point.y - entry.drawing.transform.y) * 0.75 * projection.scaleY).toFixed(2)} l\n`;
        content += 'S\n';
      }
      continue;
    }
    if (entry.image) {
      const image = registerPdfImage(state, entry.image);
      content += `q ${width.toFixed(2)} 0 0 ${height.toFixed(2)} ${x.toFixed(2)} ${y.toFixed(2)} cm /${image.name} Do Q\n`;
      content += `${formatRgb('#64748b')} RG 0.7 w ${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re S\n`;
      continue;
    }
    if (entry.chart) {
      content += `${formatRgb('#ffffff')} rg ${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re f\n`;
      content += `${formatRgb('#64748b')} RG 0.7 w ${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re S\n`;
      content += pdfChartPath(entry.chart, x, y, width, height);
      continue;
    }
    if (payload.kind === 'image') throw new Error(`NATIVE_PRINT_RESOURCE_UNAVAILABLE: image bytes were not supplied for ${entry.drawing.id}`);
    const fill = payload.kind === 'shape' ? payload.fill : payload.kind === 'textbox' ? payload.textFrame.textColor : '#f8fafc';
    const stroke = payload.kind === 'shape' ? payload.stroke : '#64748b';
    const shapePath = payload.kind === 'shape' ? pdfShapePath(payload.type, x, y, width, height) : `${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re`;
    if (payload.kind !== 'shape' || !['line', 'arrow'].includes(payload.type)) content += `${formatRgb(fill, '#f8fafc')} rg ${shapePath} f\n`;
    content += `${formatRgb(stroke, '#64748b')} RG 0.7 w ${shapePath} S\n`;
    const objectText = payload.kind === 'shape' ? payload.text : payload.kind === 'textbox' ? payload.text : undefined;
    if (objectText) content += `BT\n/F1 8 Tf\n1 0 0 1 ${(x + 3).toFixed(2)} ${(y + Math.min(height - 3, 11)).toFixed(2)} Tm\n${pdfText(objectText)} Tj\nET\n`;
  }
  return `${content}Q\n`;
}

function pdfChartPath(chart: NonNullable<PrintProjection['drawings'][number]['chart']>, x: number, y: number, width: number, height: number): string {
  const plotLeft = x + 24;
  const plotBottom = y + 20;
  const plotWidth = Math.max(1, width - 34);
  const plotHeight = Math.max(1, height - 32);
  const values = chart.series.flatMap((series) => series.values.filter((value): value is number => value !== null && Number.isFinite(value)));
  const minimum = Math.min(0, ...values);
  const maximum = Math.max(1, ...values);
  const scale = (value: number) => plotBottom + ((value - minimum) / Math.max(Number.EPSILON, maximum - minimum)) * plotHeight;
  let content = `${formatRgb('#cbd5e1')} RG 0.35 w ${plotLeft.toFixed(2)} ${plotBottom.toFixed(2)} m ${(plotLeft + plotWidth).toFixed(2)} ${plotBottom.toFixed(2)} l S\n`;
  const categoryCount = Math.max(1, chart.categories.length);
  const slot = plotWidth / categoryCount;
  chart.series.forEach((series, seriesIndex) => {
    const color = series.color ?? ['#2563eb', '#10b981', '#f59e0b', '#ef4444'][seriesIndex % 4]!;
    const barWidth = Math.max(1, slot / Math.max(1, chart.series.length + 1));
    series.values.forEach((value, index) => {
      if (value === null || !Number.isFinite(value)) return;
      const barX = plotLeft + index * slot + seriesIndex * barWidth + 1;
      const top = scale(Math.max(0, value));
      const bottom = scale(Math.min(0, value));
      content += `${formatRgb(color)} rg ${barX.toFixed(2)} ${Math.min(top, bottom).toFixed(2)} ${Math.max(0.5, barWidth - 1).toFixed(2)} ${Math.max(0.5, Math.abs(top - bottom)).toFixed(2)} re f\n`;
    });
  });
  return content;
}

function pdfShapePath(type: import('@react-sheets/core-model').ShapeDrawingType, x: number, y: number, width: number, height: number): string {
  const point = (px: number, py: number): string => `${px.toFixed(2)} ${py.toFixed(2)}`;
  const polygon = (points: readonly [number, number][]): string => `${point(points[0]![0], points[0]![1])} m ${points.slice(1).map(([px, py]) => `${point(px, py)} l`).join(' ')} h`;
  switch (type) {
    case 'triangle': return polygon([[x + width / 2, y + height], [x + width, y], [x, y]]);
    case 'right-triangle': return polygon([[x, y + height], [x + width, y], [x, y]]);
    case 'diamond': return polygon([[x + width / 2, y + height], [x + width, y + height / 2], [x + width / 2, y], [x, y + height / 2]]);
    case 'hexagon': return polygon([[x + width * 0.2, y + height], [x + width * 0.8, y + height], [x + width, y + height / 2], [x + width * 0.8, y], [x + width * 0.2, y], [x, y + height / 2]]);
    case 'parallelogram': return polygon([[x + width * 0.2, y + height], [x + width, y + height], [x + width * 0.8, y], [x, y]]);
    case 'trapezoid': return polygon([[x + width * 0.2, y + height], [x + width * 0.8, y + height], [x + width, y], [x, y]]);
    case 'line':
    case 'arrow': return `${point(x, y + height / 2)} m ${point(x + width, y + height / 2)} l`;
    case 'ellipse': {
      const k = 0.5522848;
      const cx = x + width / 2; const cy = y + height / 2; const rx = width / 2; const ry = height / 2;
      return `${point(cx + rx, cy)} m ${point(cx + rx, cy + k * ry)} ${point(cx + k * rx, cy + ry)} ${point(cx, cy + ry)} c ${point(cx - k * rx, cy + ry)} ${point(cx - rx, cy + k * ry)} ${point(cx - rx, cy)} c ${point(cx - rx, cy - k * ry)} ${point(cx - k * rx, cy - ry)} ${point(cx, cy - ry)} c ${point(cx + k * rx, cy - ry)} ${point(cx + rx, cy - k * ry)} ${point(cx + rx, cy)} c h`;
    }
    default: return `${x.toFixed(2)} ${y.toFixed(2)} ${Math.max(0.1, width).toFixed(2)} ${Math.max(0.1, height).toFixed(2)} re`;
  }
}

function registerPdfImage(state: PdfWriterState, input: { asset: { contentHash: string; mimeType: string; width?: number; height?: number }; bytes?: Uint8Array }): PdfImageResource {
  const contentHash = input.asset.contentHash;
  const width = input.asset.width;
  const height = input.asset.height;
  if (typeof width !== 'number' || typeof height !== 'number' || !Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width < 1 || height < 1) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_INVALID', `image dimensions are required for ${contentHash}`);
  if (width > PDF_IMAGE_RESOURCE_LIMITS.maxWidth || height > PDF_IMAGE_RESOURCE_LIMITS.maxHeight || width * height > PDF_IMAGE_RESOURCE_LIMITS.maxPixels) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_LIMIT', `image dimensions ${width}x${height} exceed the resource budget for ${contentHash}`);
  const mimeType = input.asset.mimeType.trim().toLowerCase();
  const existing = state.images.get(contentHash);
  if (existing) {
    if (existing.asset.mimeType !== mimeType || existing.asset.width !== width || existing.asset.height !== height) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_CONFLICT', `asset metadata changed for ${contentHash}`);
    return existing;
  }
  const bytes = input.bytes;
  if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_UNAVAILABLE', `image bytes were not supplied for ${contentHash}`);
  if (bytes.byteLength > PDF_IMAGE_RESOURCE_LIMITS.maxInputBytes) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_LIMIT', `encoded image exceeds ${PDF_IMAGE_RESOURCE_LIMITS.maxInputBytes} bytes for ${contentHash}`);

  let embeddedBytes: Uint8Array;
  let filter: PdfImageResource['filter'];
  if (mimeType === 'image/jpeg' || mimeType === 'image/jpg') {
    if (bytes[0] !== 0xff || bytes[1] !== 0xd8) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_INVALID', `JPEG signature is invalid for ${contentHash}`);
    embeddedBytes = bytes.slice();
    filter = 'DCTDecode';
  } else if (mimeType === 'image/png') {
    let decoded;
    try {
      decoded = decodePngToRgb(bytes, PDF_IMAGE_RESOURCE_LIMITS);
    } catch (error) {
      if (error instanceof PngDecodeError) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_DECODE_FAILED', `${contentHash} (${error.code})`, error);
      throw new PdfResourceError('NATIVE_PRINT_RESOURCE_DECODE_FAILED', `PNG decode failed for ${contentHash}`, error);
    }
    if (decoded.width !== width || decoded.height !== height) throw new PdfResourceError('NATIVE_PRINT_RESOURCE_DIMENSIONS_MISMATCH', `PNG dimensions ${decoded.width}x${decoded.height} do not match AssetRef ${width}x${height} for ${contentHash}`);
    try {
      embeddedBytes = zlibSync(decoded.rgb);
    } catch (error) {
      throw new PdfResourceError('NATIVE_PRINT_RESOURCE_DECODE_FAILED', `RGB compression failed for ${contentHash}`, error);
    }
    filter = 'FlateDecode';
  } else {
    throw new PdfResourceError('NATIVE_PRINT_RESOURCE_UNSUPPORTED_FORMAT', `PDF image embedding supports JPEG and PNG only, received ${input.asset.mimeType} for ${contentHash}`);
  }
  const image: PdfImageResource = { name: `Im${state.images.size + 1}`, asset: { contentHash, mimeType, width, height }, bytes: embeddedBytes, filter };
  state.images.set(contentHash, image);
  return image;
}

/**
 * Deterministic PDF 1.7 writer. The bytes are valid PDF
 * objects and can be opened by browser viewers, Preview, Acrobat, and
 * Poppler. Cell painting is supplied by the optional pageText port; the
 * pagination and document metadata remain identical in browser and Node.
 */
export function createPdfBytes(pages: readonly PrintPageInfo[], layout: PrintLayoutModel, options: PdfExportOptions = {}): Uint8Array {
  const objects: string[] = [];
  const imageObjects = new Map<number, PdfImageResource>();
  const writerState: PdfWriterState = { images: new Map() };
  const pageObjectIds: number[] = [];
  const pageEntries: Array<{ pageObjectId: number; contentObjectId: number; content: string }> = [];
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  // Use a built-in CID font with UTF-16BE text strings so CJK and other
  // non-ASCII workbook labels survive both Node and browser PDF output.
  objects[3] = '<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniCNS-UTF16-H /DescendantFonts [4 0 R] >>';
  objects[4] = '<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (CNS1) /Supplement 0 >> /DW 1000 >>';
  let nextObjectId = 5;
  for (let index = 0; index < pages.length; index += 1) {
    const pageObjectId = nextObjectId++;
    const contentObjectId = nextObjectId++;
    pageObjectIds.push(pageObjectId);
    const content = pageContent(pages[index]!, index + 1, pages.length, layout, options, writerState);
    pageEntries.push({ pageObjectId, contentObjectId, content });
    objects[contentObjectId] = `<< /Length ${content.length} >>\nstream\n${content}endstream`;
  }
  for (const image of writerState.images.values()) {
    const objectId = nextObjectId++;
    image.objectId = objectId;
    imageObjects.set(objectId, image);
  }
  for (const entry of pageEntries) {
    const xObject = writerState.images.size > 0
      ? ` /XObject << ${[...writerState.images.values()].map((image) => `/${image.name} ${image.objectId} 0 R`).join(' ')} >>`
      : '';
    objects[entry.pageObjectId] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageSize(layout).width.toFixed(2)} ${pageSize(layout).height.toFixed(2)}] /Resources << /Font << /F1 3 0 R >>${xObject} >> /Contents ${entry.contentObjectId} 0 R >>`;
  }
  objects[2] = `<< /Type /Pages /Kids [${pageObjectIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageObjectIds.length} >>`;
  const infoObjectId = nextObjectId++;
  objects[infoObjectId] = `<< /Title ${pdfText(options.title ?? 'Workbook')} /Author ${pdfText(options.author ?? 'Spreadsheet')} >>`;

  const chunks: Uint8Array[] = [new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37, 0x0a, 0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a])];
  const offsets = new Array<number>(nextObjectId).fill(0);
  let length = chunks[0]!.length;
  for (let objectId = 1; objectId < nextObjectId; objectId += 1) {
    const object = objects[objectId];
    const image = imageObjects.get(objectId);
    if (!object && !image) continue;
    const bytes = image
      ? concatBytes([ascii(`${objectId} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${image.asset.width} /Height ${image.asset.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /${image.filter}${image.filter === 'FlateDecode' ? ` /DecodeParms << /Colors 3 /BitsPerComponent 8 /Columns ${image.asset.width} >>` : ''} /Length ${image.bytes.byteLength} >>\nstream\n`), image.bytes, ascii('\nendstream\nendobj\n')])
      : ascii(`${objectId} 0 obj\n${object}\nendobj\n`);
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
