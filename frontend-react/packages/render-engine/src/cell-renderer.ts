import { formatValue } from "@react-sheets/number-format";
import * as bwipjs from "@bwip-js/browser";
import {
  type CellAddress,
  type CellProvider,
  type CellRange,
  type CellRenderData,
  type Rect,
  type RenderPane,
  type RenderTheme,
  type FloatingDrawable,
} from "./types";
import { SheetSkeleton, columnLabelOf } from "./sheet-skeleton";

export interface PaneDrawOptions {
  context: CanvasRenderingContext2D;
  skeleton: SheetSkeleton;
  pane: RenderPane;
  visibleRange: CellRange | null;
  cellProvider: CellProvider;
  theme: RenderTheme;
  /** 内容坐标系下的局部重绘矩形 */
  drawRects?: readonly Rect[];
}

export interface ExtensionsDrawOptions extends PaneDrawOptions {
  floatables: readonly FloatingDrawable[];
}

function shouldDrawRect(rects: readonly Rect[] | undefined, rect: Rect): boolean {
  if (!rects || rects.length === 0) return true;
  return rects.some((candidate) =>
    candidate.x < rect.x + rect.width
    && candidate.x + candidate.width > rect.x
    && candidate.y < rect.y + rect.height
    && candidate.y + candidate.height > rect.y);
}

// ---------------- 网格层 ----------------

export function drawGridLayer(options: PaneDrawOptions): void {
  const { context, skeleton, visibleRange, theme, pane, drawRects } = options;
  const background = { x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height };
  context.fillStyle = theme.canvasBackground;
  context.fillRect(background.x, background.y, background.width, background.height);
  if (!visibleRange) return;

  // 网格线属于整张可见网格，而不是 occupied cells。先收集可见合并区域，
  // 再逐个空白/有值单元格绘制四条边，确保空白单元格仍保持完整网格。
  // 旧实现只在每一行/列的第一个单元格后 break，导致内容区只剩 A 列和首行的线。
  const merges = collectVisibleMerges(options, visibleRange);
  context.strokeStyle = theme.gridLine;
  context.lineWidth = 1;
  context.beginPath();
  for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
    if (skeleton.isRowHidden(row)) continue;
    for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column++) {
      if (skeleton.isColumnHidden(column)) continue;
      const cell = options.cellProvider({ row, column });
      if (cell?.merge) continue;
      const x = skeleton.getColumnLeft(column);
      const y = skeleton.getRowTop(row);
      const width = skeleton.getColumnWidth(column);
      const height = skeleton.getRowHeight(row);
      const rect = { x, y, width, height };
      if (!shouldDrawRect(drawRects, rect)) continue;
      const right = x + width;
      const bottom = y + height;
      context.moveTo(Math.round(x) + 0.5, Math.round(y) + 0.5);
      context.lineTo(Math.round(right) + 0.5, Math.round(y) + 0.5);
      context.moveTo(Math.round(x) + 0.5, Math.round(bottom) + 0.5);
      context.lineTo(Math.round(right) + 0.5, Math.round(bottom) + 0.5);
      context.moveTo(Math.round(x) + 0.5, Math.round(y) + 0.5);
      context.lineTo(Math.round(x) + 0.5, Math.round(bottom) + 0.5);
      context.moveTo(Math.round(right) + 0.5, Math.round(y) + 0.5);
      context.lineTo(Math.round(right) + 0.5, Math.round(bottom) + 0.5);
    }
  }

  // 非锚点单元格被跳过后，合并区域的外框由这里一次性绘制；内部不生成线。
  for (const merge of merges.values()) {
    const rect = {
      x: skeleton.getColumnLeft(merge.startColumn),
      y: skeleton.getRowTop(merge.startRow),
      width: sumWidth(skeleton, merge.startColumn, merge.endColumn),
      height: sumHeight(skeleton, merge.startRow, merge.endRow),
    };
    if (!shouldDrawRect(drawRects, rect)) continue;
    const right = rect.x + rect.width;
    const bottom = rect.y + rect.height;
    context.moveTo(Math.round(rect.x) + 0.5, Math.round(rect.y) + 0.5);
    context.lineTo(Math.round(right) + 0.5, Math.round(rect.y) + 0.5);
    context.moveTo(Math.round(rect.x) + 0.5, Math.round(bottom) + 0.5);
    context.lineTo(Math.round(right) + 0.5, Math.round(bottom) + 0.5);
    context.moveTo(Math.round(rect.x) + 0.5, Math.round(rect.y) + 0.5);
    context.lineTo(Math.round(rect.x) + 0.5, Math.round(bottom) + 0.5);
    context.moveTo(Math.round(right) + 0.5, Math.round(rect.y) + 0.5);
    context.lineTo(Math.round(right) + 0.5, Math.round(bottom) + 0.5);
  }

  context.stroke();
  void background;
}

interface VisibleMerge {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

function collectVisibleMerges(options: PaneDrawOptions, range: CellRange): Map<string, VisibleMerge> {
  const merges = new Map<string, VisibleMerge>();
  const { cellProvider } = options;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      const merge = cellProvider({ row, column })?.merge;
      if (!merge) continue;
      const value = {
        startRow: merge.startRow,
        endRow: merge.endRow,
        startColumn: merge.startColumn,
        endColumn: merge.endColumn,
      };
      const key = `${value.startRow}:${value.startColumn}:${value.endRow}:${value.endColumn}`;
      merges.set(key, value);
    }
  }
  return merges;
}

// ---------------- 内容层 ----------------

export function drawCellLayer(options: PaneDrawOptions): void {
  const { context, skeleton, visibleRange, cellProvider, theme, drawRects } = options;
  if (!visibleRange) return;
  context.textBaseline = "middle";

  for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
    if (skeleton.isRowHidden(row)) continue;
    for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column++) {
      if (skeleton.isColumnHidden(column)) continue;
      const address: CellAddress = { row, column };
      const rect: Rect = {
        x: skeleton.getColumnLeft(column),
        y: skeleton.getRowTop(row),
        width: skeleton.getColumnWidth(column),
        height: skeleton.getRowHeight(row),
      };
      if (!shouldDrawRect(drawRects, rect)) continue;
      const cell = cellProvider(address);
      const merge = cell?.merge;
      const isAnchor = !merge || merge.isAnchor;
      const spanRect: Rect = merge && merge.isAnchor
        ? {
            x: skeleton.getColumnLeft(merge.startColumn),
            y: skeleton.getRowTop(merge.startRow),
            width: sumWidth(skeleton, merge.startColumn, merge.endColumn),
            height: sumHeight(skeleton, merge.startRow, merge.endRow),
          }
        : rect;

      if (cell?.overlay?.colorScale || cell?.style?.background || (cell === undefined && false)) {
        // 背景(含色阶)
      }
      const backgroundColor = cell?.overlay?.colorScale ?? cell?.style?.background;
      if (backgroundColor && isAnchor) {
        context.fillStyle = backgroundColor;
        context.fillRect(spanRect.x, spanRect.y, spanRect.width, spanRect.height);
      }

      if (cell?.overlay?.dataBar && isAnchor) {
        drawDataBar(context, spanRect, cell.overlay.dataBar);
      }

      if (isAnchor && cell) {
        drawCustomBorders(context, spanRect, cell);
      }

      if (!isAnchor) continue;

      if (cell?.hasComment) drawCommentMark(context, spanRect, theme);
      if (cell?.invalid) drawInvalidRing(context, spanRect, theme);

      if (cell) {
        if (cell.presentation?.kind === 'barcode') drawBarcodePresentation(context, spanRect, resolveDisplayText(cell), cell.presentation);
        else if (cell.presentation?.kind === 'image') drawCellImagePresentation(context, spanRect, cell.presentation);
        else drawCellValue(context, skeleton, options, address, cell, spanRect);
        if (cell.editor?.kind === 'checkbox') drawCheckboxEditor(context, spanRect, cell.value === true || String(cell.value).toUpperCase() === 'TRUE');
        if (cell.overlay?.icon) drawTrendIcon(context, spanRect, cell.overlay.icon);
      }
    }
  }
}

const barcodeCanvasCache = new Map<string, HTMLCanvasElement | OffscreenCanvas>();
const cellImageCache = new Map<string, HTMLImageElement>();
const BARCODE_ENCODERS: Record<Extract<NonNullable<CellRenderData['presentation']>, { kind: 'barcode' }>['symbology'], string> = {
  qr: 'qrcode', code128: 'code128', code39: 'code39', code93: 'code93', code49: 'code49', codabar: 'rationalizedCodabar', ean13: 'ean13', ean8: 'ean8', upca: 'upca', 'gs1-128': 'gs1-128', pdf417: 'pdf417', 'data-matrix': 'datamatrix',
};

function drawBarcodePresentation(
  context: CanvasRenderingContext2D,
  rect: Rect,
  value: string,
  presentation: Extract<NonNullable<CellRenderData['presentation']>, { kind: 'barcode' }>,
): void {
  const quiet = Math.max(1, presentation.options.quietZone);
  const width = Math.max(8, Math.floor(rect.width - quiet * 2));
  const height = Math.max(8, Math.floor(rect.height - quiet * 2));
  const cacheKey = `${presentation.symbology}|${value}|${width}|${height}|${presentation.options.foreground}|${presentation.options.background}|${presentation.options.showText}|${presentation.options.labelPosition}|${presentation.options.fontSize ?? ''}|${JSON.stringify(presentation.parameters)}`;
  context.save();
  context.fillStyle = presentation.options.background;
  context.fillRect(rect.x, rect.y, rect.width, rect.height);
  try {
    let canvas = barcodeCanvasCache.get(cacheKey);
    if (!canvas) {
      canvas = typeof OffscreenCanvas !== 'undefined'
        ? new OffscreenCanvas(width, height)
        : document.createElement('canvas');
      const fullAscii = presentation.parameters.symbology === presentation.symbology
        && 'fullAscii' in presentation.parameters && presentation.parameters.fullAscii === true;
      const bcid = presentation.symbology === 'code39' && fullAscii ? 'code39ext'
        : presentation.symbology === 'code93' && fullAscii ? 'code93ext'
          : BARCODE_ENCODERS[presentation.symbology];
      bwipjs.toCanvas(canvas, {
        bcid,
        text: value,
        scale: 1,
        height: Math.max(4, Math.floor(height / 3)),
        includetext: presentation.options.showText && presentation.options.labelPosition !== 'none',
        textyalign: presentation.options.labelPosition === 'above' ? 'above' : 'below',
        textsize: presentation.options.fontSize,
        includecheck: presentation.parameters.symbology === presentation.symbology && 'includeCheckDigit' in presentation.parameters ? presentation.parameters.includeCheckDigit : undefined,
        addontext: presentation.parameters.symbology === presentation.symbology && 'addOnText' in presentation.parameters ? presentation.parameters.addOnText : undefined,
        eclevel: presentation.parameters.symbology === 'qr' && presentation.parameters.symbology === presentation.symbology && 'errorCorrection' in presentation.parameters
          ? ({ low: 'L', medium: 'M', quartile: 'Q', high: 'H' } as const)[presentation.parameters.errorCorrection ?? 'medium']
          : presentation.parameters.symbology === 'pdf417' && presentation.parameters.symbology === presentation.symbology && 'securityLevel' in presentation.parameters
            ? presentation.parameters.securityLevel
            : undefined,
        ratio: presentation.parameters.symbology === presentation.symbology && 'wideNarrowRatio' in presentation.parameters ? presentation.parameters.wideNarrowRatio : undefined,
        parse: presentation.parameters.symbology === presentation.symbology && 'fullAscii' in presentation.parameters ? presentation.parameters.fullAscii : undefined,
        textxalign: 'center',
        backgroundcolor: presentation.options.background.replace('#', ''),
        barcolor: presentation.options.foreground.replace('#', ''),
      } as Parameters<typeof bwipjs.toCanvas>[1]);
      barcodeCanvasCache.set(cacheKey, canvas);
      if (barcodeCanvasCache.size > 256) barcodeCanvasCache.delete(barcodeCanvasCache.keys().next().value!);
    }
    context.drawImage(canvas, rect.x + quiet, rect.y + quiet, width, height);
  } catch {
    context.fillStyle = '#b91c1c';
    context.font = '10px "Microsoft YaHei", "Segoe UI", sans-serif';
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.fillText('Invalid barcode', rect.x + rect.width / 2, rect.y + rect.height / 2, width);
  }
  context.restore();
}

function drawCellImagePresentation(context: CanvasRenderingContext2D, rect: Rect, presentation: Extract<NonNullable<CellRenderData['presentation']>, { kind: 'image' }>): void {
  if (typeof Image === 'undefined') return;
  let image = cellImageCache.get(presentation.src);
  if (!image) {
    image = new Image();
    image.src = presentation.src;
    cellImageCache.set(presentation.src, image);
  }
  if (!image.complete || image.naturalWidth <= 0) return;
  let x = rect.x;
  let y = rect.y;
  let width = rect.width;
  let height = rect.height;
  if (presentation.fit !== 'stretch') {
    const scale = presentation.fit === 'cover' ? Math.max(width / image.naturalWidth, height / image.naturalHeight) : Math.min(width / image.naturalWidth, height / image.naturalHeight);
    width = image.naturalWidth * scale;
    height = image.naturalHeight * scale;
    x += (rect.width - width) / 2;
    y += (rect.height - height) / 2;
  }
  context.save();
  context.beginPath();
  context.rect(rect.x, rect.y, rect.width, rect.height);
  context.clip();
  context.drawImage(image, x, y, width, height);
  context.restore();
}

function drawCheckboxEditor(context: CanvasRenderingContext2D, rect: Rect, checked: boolean): void {
  const size = Math.min(14, Math.max(10, rect.height - 8));
  const x = rect.x + 4;
  const y = rect.y + (rect.height - size) / 2;
  context.save();
  context.fillStyle = '#ffffff';
  context.strokeStyle = checked ? '#217345' : '#94a3b8';
  context.lineWidth = 1;
  context.fillRect(x, y, size, size);
  context.strokeRect(x + 0.5, y + 0.5, size - 1, size - 1);
  if (checked) {
    context.strokeStyle = '#217345';
    context.lineWidth = 2;
    context.beginPath();
    context.moveTo(x + 2.5, y + size / 2);
    context.lineTo(x + size / 2 - 1, y + size - 3);
    context.lineTo(x + size - 2, y + 3);
    context.stroke();
  }
  context.restore();
}

function sumWidth(skeleton: SheetSkeleton, startColumn: number, endColumn: number): number {
  let total = 0;
  for (let c = startColumn; c <= endColumn; c++) total += skeleton.getColumnWidth(c);
  return total;
}

function sumHeight(skeleton: SheetSkeleton, startRow: number, endRow: number): number {
  let total = 0;
  for (let r = startRow; r <= endRow; r++) total += skeleton.getRowHeight(r);
  return total;
}

function drawDataBar(context: CanvasRenderingContext2D, rect: Rect, bar: { color: string; ratio: number }): void {
  const ratio = Math.max(0, Math.min(1, bar.ratio));
  const width = Math.max(0, (rect.width - 4) * ratio);
  context.fillStyle = bar.color;
  context.globalAlpha = 0.55;
  context.fillRect(rect.x + 2, rect.y + 3, width, rect.height - 6);
  context.globalAlpha = 1;
}

function drawTrendIcon(context: CanvasRenderingContext2D, rect: Rect, icon: "up" | "down" | "flat"): void {
  const size = 9;
  const cx = rect.x + rect.width - size - 5;
  const cy = rect.y + rect.height / 2;
  context.beginPath();
  if (icon === "up") {
    context.moveTo(cx, cy - size / 2);
    context.lineTo(cx + size, cy - size / 2);
    context.lineTo(cx + size / 2, cy + size / 2);
    context.closePath();
    context.fillStyle = "#16a34a";
  } else if (icon === "down") {
    context.moveTo(cx, cy + size / 2);
    context.lineTo(cx + size, cy + size / 2);
    context.lineTo(cx + size / 2, cy - size / 2);
    context.closePath();
    context.fillStyle = "#dc2626";
  } else {
    context.rect(cx, cy - 1.25, size, 2.5);
    context.fillStyle = "#64748b";
  }
  context.fill();
}

function drawCommentMark(context: CanvasRenderingContext2D, rect: Rect, theme: RenderTheme): void {
  const size = 7;
  context.beginPath();
  context.moveTo(rect.x + rect.width - size, rect.y);
  context.lineTo(rect.x + rect.width, rect.y);
  context.lineTo(rect.x + rect.width, rect.y + size);
  context.closePath();
  context.fillStyle = theme.commentMarkColor;
  context.fill();
}

function drawInvalidRing(context: CanvasRenderingContext2D, rect: Rect, theme: RenderTheme): void {
  context.save();
  context.strokeStyle = theme.invalidColor;
  context.setLineDash([3, 2]);
  context.strokeRect(rect.x + 1.5, rect.y + 1.5, rect.width - 3, rect.height - 3);
  context.restore();
}

function drawCustomBorders(context: CanvasRenderingContext2D, rect: Rect, cell: CellRenderData): void {
  const borders = cell.style?.borders;
  if (!borders) return;
  context.lineWidth = 1;
  const sides: Array<["top" | "right" | "bottom" | "left", [number, number, number, number]]> = [
    ["top", [rect.x, rect.y + 0.5, rect.x + rect.width, rect.y + 0.5]],
    ["bottom", [rect.x, rect.y + rect.height - 0.5, rect.x + rect.width, rect.y + rect.height - 0.5]],
    ["left", [rect.x + 0.5, rect.y, rect.x + 0.5, rect.y + rect.height]],
    ["right", [rect.x + rect.width - 0.5, rect.y, rect.x + rect.width - 0.5, rect.y + rect.height]],
  ];
  for (const [side, coordinates] of sides) {
    const border = borders[side];
    if (!border) continue;
    context.strokeStyle = border.color;
    context.setLineDash(border.style === "dashed" ? [4, 3] : []);
    context.lineWidth = border.style === "thick" ? 2 : 1;
    context.beginPath();
    context.moveTo(coordinates[0], coordinates[1]);
    context.lineTo(coordinates[2], coordinates[3]);
    context.stroke();
  }
  context.setLineDash([]);
  context.lineWidth = 1;
}

export function resolveDisplayText(cell: CellRenderData): string {
  if (cell.displayValue !== undefined) return cell.displayValue;
  if (typeof cell.value === "number") {
    return formatValue(cell.value, cell.style?.numberFormat);
  }
  if (cell.value == null) return "";
  if (typeof cell.value === "boolean") return cell.value ? "TRUE" : "FALSE";
  return String(cell.value);
}

export function cellRenderFont(style: CellRenderData["style"], theme: RenderTheme): string {
  const size = style?.fontSizePx ?? 13;
  const family = style?.fontFamily ? '"' + style.fontFamily + '", sans-serif' : '"Microsoft YaHei", "Segoe UI", sans-serif';
  const weight = style?.bold ? "700" : "400";
  const slant = style?.italic ? " italic" : "";
  return slant + " " + weight + " " + size + "px " + family;
}

export interface AutoFitMeasurement {
  widthPx: number;
  heightPx: number;
}

/** The sole text geometry calculation shared by CellRenderer and dimension AutoFit. */
export function measureCellAutoFit(
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  cell: CellRenderData,
  theme: RenderTheme,
  availableWidthPx?: number,
  reserveFilterButton = false,
): AutoFitMeasurement {
  const text = resolveDisplayText(cell);
  const style = cell.style;
  const padding = style?.padding ?? theme.cellPadding;
  context.save();
  context.font = cellRenderFont(style, theme);
  const lines = text.split(/\r?\n/);
  const rawWidth = Math.max(0, ...lines.map((line) => context.measureText(line).width));
  const fontSizePx = style?.fontSizePx ?? 13;
  const lineHeight = Math.max(fontSizePx * 1.25, 16);
  const indent = Math.max(0, Math.trunc(style?.indent ?? 0)) * 12;
  let width = rawWidth + padding * 2 + indent + (reserveFilterButton ? 18 : 0) + (style?.borders?.left ? 1 : 0) + (style?.borders?.right ? 1 : 0);
  let lineCount = Math.max(1, lines.length);
  if (style?.wrapText && availableWidthPx && availableWidthPx > padding * 2) {
    lineCount = lines.reduce((count, line) => count + Math.max(1, Math.ceil(context.measureText(line).width / Math.max(1, availableWidthPx - padding * 2))), 0);
    width = Math.min(width, availableWidthPx);
  }
  let height = lineCount * lineHeight + padding * 2 + (style?.borders?.top ? 1 : 0) + (style?.borders?.bottom ? 1 : 0);
  const rotation = Math.abs((style?.textRotate ?? 0) * Math.PI / 180);
  if (rotation > 0) {
    const rotatedWidth = Math.abs(Math.cos(rotation)) * width + Math.abs(Math.sin(rotation)) * height;
    const rotatedHeight = Math.abs(Math.sin(rotation)) * width + Math.abs(Math.cos(rotation)) * height;
    width = rotatedWidth;
    height = rotatedHeight;
  }
  context.restore();
  return { widthPx: Math.ceil(width), heightPx: Math.ceil(height) };
}

function drawCellValue(
  context: CanvasRenderingContext2D,
  skeleton: SheetSkeleton,
  options: PaneDrawOptions,
  address: CellAddress,
  cell: CellRenderData,
  rect: Rect,
): void {
  const { theme } = options;
  const style = cell.style;
  const text = resolveDisplayText(cell);
  if (!text) return;

  const padding = style?.padding ?? theme.cellPadding;
  const indent = Math.max(0, Math.trunc(style?.indent ?? 0)) * 12;
  const hAlign = style?.horizontalAlignment ?? "left";
  const vAlign = style?.verticalAlignment ?? "middle";

  context.save();
  context.font = cellRenderFont(style, theme);
  context.fillStyle = style?.textColor ?? theme.cellText;
  context.textBaseline = "middle";

  const wrap = Boolean(style?.wrapText);
  const maxWidth = rect.width - padding * 2 - indent;
  const measured = context.measureText(text).width;

  if (wrap) {
    drawWrapped(context, text, rect, padding + indent, hAlign, vAlign, maxWidth);
    context.restore();
    return;
  }

  // 溢出:左对齐文本超宽且右侧邻格为空时向右溢出
  let overflowWidth = maxWidth;
  if (hAlign === "left" && measured > maxWidth) {
    const allowed = allowedOverflowWidth(options, address, text, measured, rect);
    overflowWidth = allowed;
  }

  let x: number;
  if (hAlign === "center") x = rect.x + rect.width / 2;
  else if (hAlign === "right") x = rect.x + rect.width - padding - indent;
  else x = rect.x + padding + indent;

  const fontSize = style?.fontSizePx ?? 13;
  let y = vAlign === 'top'
    ? rect.y + padding + fontSize / 2
    : vAlign === 'bottom'
      ? rect.y + rect.height - padding - fontSize / 2
      : rect.y + rect.height / 2;
  const rotate = style?.textRotate ?? 0;
  if (rotate !== 0) {
    const radians = (rotate * Math.PI) / 180;
    context.translate(rect.x + rect.width / 2, rect.y + rect.height / 2);
    context.rotate(-radians);
    x = 0;
    y = 0;
  }

  context.textAlign = hAlign === "center" ? "center" : hAlign;
  context.fillText(text, x, y, overflowWidth);

  if (style?.underline || style?.strikethrough) {
    const textWidth = Math.min(measured, overflowWidth);
    let lineX1 = hAlign === "center" ? x - textWidth / 2 : hAlign === "right" ? x - textWidth : x;
    const lineX2 = lineX1 + textWidth;
    context.strokeStyle = style.textColor ?? theme.cellText;
    context.lineWidth = 1;
    context.beginPath();
    if (style.underline) {
      const lineY = y + 7;
      context.moveTo(lineX1, lineY);
      context.lineTo(lineX2, lineY);
    }
    if (style.strikethrough) {
      context.moveTo(lineX1, y);
      context.lineTo(lineX2, y);
    }
    context.stroke();
  }

  context.restore();
  void skeleton;
}

/** 计算允许的溢出宽度:右侧连续空单元格宽度之和 */
function allowedOverflowWidth(
  options: PaneDrawOptions,
  address: CellAddress,
  text: string,
  measured: number,
  rect: Rect,
): number {
  const { skeleton, cellProvider } = options;
  const firstCellWidth = rect.width - 12; // 自身边界内可用宽
  let available = firstCellWidth;
  let remainingNeed = measured - firstCellWidth;
  if (remainingNeed <= 0) return measured;
  let column = address.column + 1;
  while (remainingNeed > 0 && column < skeleton.columnCount) {
    const neighbor = cellProvider({ row: address.row, column });
    const neighborText = neighbor ? resolveDisplayText(neighbor) : "";
    if (neighborText) break;
    const width = skeleton.getColumnWidth(column);
    available += width;
    remainingNeed -= width;
    column += 1;
  }
  void text;
  return Math.min(measured, available);
}

function drawWrapped(
  context: CanvasRenderingContext2D,
  text: string,
  rect: Rect,
  padding: number,
  hAlign: "left" | "center" | "right",
  vAlign: "top" | "middle" | "bottom",
  maxWidth: number,
): void {
  const hasCjk = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/u.test(text);
  const tokens = hasCjk ? Array.from(text) : text.split(/\s+/);
  const lineHeight = 16;
  const lines: string[] = [];
  let currentLine = "";
  for (const token of tokens) {
    const separator = hasCjk || !currentLine ? "" : " ";
    const attempt = currentLine + separator + token;
    if (context.measureText(attempt).width > maxWidth && currentLine) {
      lines.push(currentLine);
      currentLine = token;
    } else {
      currentLine = attempt;
    }
  }
  if (currentLine) lines.push(currentLine);

  const startY = vAlign === "top"
    ? rect.y + padding + lineHeight / 2
    : vAlign === "bottom"
      ? rect.y + rect.height - padding - ((lines.length - 1) * lineHeight) - lineHeight / 2
      : rect.y + rect.height / 2 - ((lines.length - 1) * lineHeight) / 2;

  context.textAlign = hAlign;
  for (let i = 0; i < lines.length; i++) {
    context.fillText(lines[i]!, hAlign === "center" ? rect.x + rect.width / 2 : hAlign === "right" ? rect.x + rect.width - padding : rect.x + padding, startY + i * lineHeight, maxWidth);
  }
}

// ---------------- 浮动对象层 ----------------

export function drawExtensionsLayer(options: ExtensionsDrawOptions): void {
  const { context, floatables, pane } = options;
  const paneContent = { x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height };
  for (const drawable of floatables) {
    const b = drawable.bounds;
    const intersects = b.x < paneContent.x + paneContent.width
      && b.x + b.width > paneContent.x
      && b.y < paneContent.y + paneContent.height
      && b.y + b.height > paneContent.y;
    if (!intersects) continue;
    drawable.draw(context, b);
  }
}

export { columnLabelOf };
