import { formatValue } from "@react-sheets/number-format";
import * as bwipjs from "@bwip-js/browser";
import {
  type CellAddress,
  type CellValue,
  type CellProvider,
  type CellRange,
  type CellRenderData,
  type Rect,
  type RenderPane,
  type RenderTheme,
  type FloatingDrawable,
} from "./types";
import { SheetSkeleton, columnLabelOf } from "./sheet-skeleton";
import { checkboxStateFromValue, type AssetRef, type CheckboxCellState } from '@react-sheets/core-model';
import { resolveCellContentLayout, resolveTextRotationDegrees, type CellLayoutNeighbor } from './cell-content-layout';

export { cellRenderFont, resolveCellContentLayout, resolveTextRotationDegrees } from './cell-content-layout';

export type AssetUrlResolver = (asset: AssetRef) => Promise<string>;

export interface PaneDrawOptions {
  context: CanvasRenderingContext2D;
  skeleton: SheetSkeleton;
  pane: RenderPane;
  visibleRange: CellRange | null;
  cellProvider: CellProvider;
  theme: RenderTheme;
  /** 内容坐标系下的局部重绘矩形 */
  drawRects?: readonly Rect[];
  resolveAssetUrl?: AssetUrlResolver;
  assetUrlCache?: Map<string, string>;
  assetUrlPending?: Set<string>;
  assetUrlErrors?: Map<string, string>;
  requestRender?: () => void;
  /** Per-frame sparse neighbor index used by overflow text layout. */
  neighborCache?: Map<string, { left: CellLayoutNeighbor[]; right: CellLayoutNeighbor[] }>;
  /** Lazily built row occupancy; one provider pass replaces per-cell scans. */
  rowOccupancyCache?: Map<number, ReadonlySet<number>>;
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

function intersectCellRanges(left: CellRange, right: CellRange): CellRange | null {
  const startRow = Math.max(left.startRow, right.startRow);
  const endRow = Math.min(left.endRow, right.endRow);
  const startColumn = Math.max(left.startColumn, right.startColumn);
  const endColumn = Math.min(left.endColumn, right.endColumn);
  return startRow <= endRow && startColumn <= endColumn
    ? { startRow, endRow, startColumn, endColumn }
    : null;
}

/**
 * Converts incremental content rectangles into the exact model ranges that
 * need painting.  Scanning the entire visible range and filtering every cell
 * afterwards turns a one-row scroll strip into a viewport-sized hot loop.
 */
function resolveDrawRanges(
  skeleton: SheetSkeleton,
  visibleRange: CellRange,
  drawRects: readonly Rect[] | undefined,
): CellRange[] {
  if (!drawRects || drawRects.length === 0) return [visibleRange];
  const ranges: CellRange[] = [];
  for (const rect of drawRects) {
    const range = skeleton.getVisibleRange(rect);
    if (!range) continue;
    const clipped = intersectCellRanges(range, visibleRange);
    if (clipped) ranges.push(clipped);
  }
  return ranges;
}

function forEachCellInRanges(
  skeleton: SheetSkeleton,
  ranges: readonly CellRange[],
  visit: (address: CellAddress) => void,
): void {
  if (ranges.length === 0) return;
  const seen = ranges.length > 1 ? new Set<string>() : null;
  for (const range of ranges) {
    for (let row = range.startRow; row <= range.endRow; row++) {
      if (skeleton.isRowHidden(row)) continue;
      for (let column = range.startColumn; column <= range.endColumn; column++) {
        if (skeleton.isColumnHidden(column)) continue;
        if (seen) {
          const key = `${row}:${column}`;
          if (seen.has(key)) continue;
          seen.add(key);
        }
        visit({ row, column });
      }
    }
  }
}

// ---------------- 网格层 ----------------

export function drawGridLayer(options: PaneDrawOptions): void {
  const { context, skeleton, visibleRange, theme, pane, drawRects } = options;
  const background = { x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height };
  context.fillStyle = theme.canvasBackground;
  // During an incremental scroll the preserved bitmap remains valid. Filling
  // the whole pane here would erase that bitmap before the exposed strip is
  // redrawn, turning every drag frame back into a full repaint.
  const backgrounds = drawRects && drawRects.length > 0 ? drawRects : [background];
  for (const rect of backgrounds) context.fillRect(rect.x, rect.y, rect.width, rect.height);
  if (!visibleRange) return;

  // 网格线属于整张可见网格，而不是 occupied cells。先收集可见合并区域，
  // 再逐个空白/有值单元格绘制四条边，确保空白单元格仍保持完整网格。
  // 旧实现只在每一行/列的第一个单元格后 break，导致内容区只剩 A 列和首行的线。
  const drawRanges = resolveDrawRanges(skeleton, visibleRange, drawRects);
  const merges = collectVisibleMerges(options, drawRanges);
  context.strokeStyle = theme.gridLine;
  context.lineWidth = 1;
  context.beginPath();
  forEachCellInRanges(skeleton, drawRanges, ({ row, column }) => {
    const cell = options.cellProvider({ row, column });
    if (cell?.merge) return;
    const x = skeleton.getColumnLeft(column);
    const y = skeleton.getRowTop(row);
    const width = skeleton.getColumnWidth(column);
    const height = skeleton.getRowHeight(row);
    const rect = { x, y, width, height };
    if (!shouldDrawRect(drawRects, rect)) return;
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
  });

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
}

interface VisibleMerge {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

function collectVisibleMerges(options: PaneDrawOptions, ranges: readonly CellRange[]): Map<string, VisibleMerge> {
  const merges = new Map<string, VisibleMerge>();
  const { cellProvider } = options;
  forEachCellInRanges(options.skeleton, ranges, ({ row, column }) => {
    const merge = cellProvider({ row, column })?.merge;
    if (!merge) return;
    const value = {
      startRow: merge.startRow,
      endRow: merge.endRow,
      startColumn: merge.startColumn,
      endColumn: merge.endColumn,
    };
    const key = `${value.startRow}:${value.startColumn}:${value.endRow}:${value.endColumn}`;
    merges.set(key, value);
  });
  return merges;
}

// ---------------- 内容层 ----------------

export function drawCellLayer(options: PaneDrawOptions): void {
  const { context, skeleton, visibleRange, cellProvider, theme, drawRects } = options;
  if (!visibleRange) return;
  context.textBaseline = "middle";

  const addresses: CellAddress[] = [];
  const queued = new Set<string>();
  const queue = (address: CellAddress) => {
    const key = `${address.row}:${address.column}`;
    if (queued.has(key)) return;
    queued.add(key);
    addresses.push(address);
  };
  forEachCellInRanges(skeleton, resolveDrawRanges(skeleton, visibleRange, drawRects), (address) => {
    queue(address);
    const cell = cellProvider(address);
    if (cell?.merge && !cell.merge.isAnchor) {
      queue({ row: cell.merge.startRow, column: cell.merge.startColumn });
    }
    if (cell?.alignmentSpan && !cell.alignmentSpan.isAnchor) {
      queue({ row: address.row, column: cell.alignmentSpan.startColumn });
    }
  });

  const deferredContent: Array<{ address: CellAddress; cell: CellRenderData; contentRect: Rect; spanRect: Rect }> = [];
  const renderOptions: PaneDrawOptions = {
    ...options,
    neighborCache: options.neighborCache ?? new Map<string, { left: CellLayoutNeighbor[]; right: CellLayoutNeighbor[] }>(),
    rowOccupancyCache: options.rowOccupancyCache ?? new Map<number, ReadonlySet<number>>(),
  };
  for (const address of addresses) {
      const { row, column } = address;
      const rect: Rect = {
        x: skeleton.getColumnLeft(column),
        y: skeleton.getRowTop(row),
        width: skeleton.getColumnWidth(column),
        height: skeleton.getRowHeight(row),
      };
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
      const alignmentSpan = cell?.alignmentSpan;
      const contentRect = alignmentSpan && alignmentSpan.isAnchor
        ? {
            x: skeleton.getColumnLeft(alignmentSpan.startColumn),
            y: rect.y,
            width: sumWidth(skeleton, alignmentSpan.startColumn, alignmentSpan.endColumn),
            height: rect.height,
          }
        : spanRect;
      const paintRect = merge?.isAnchor ? spanRect : alignmentSpan?.isAnchor ? contentRect : rect;
      if (!shouldDrawRect(drawRects, paintRect)) continue;

      if (cell?.overlay?.colorScale || cell?.style?.background || (cell === undefined && false)) {
        // 背景(含色阶)
      }
      const backgroundColor = cell?.overlay?.colorScale ?? cell?.style?.background;
      if (isAnchor && cell?.style?.fill) drawCellFill(context, spanRect, cell.style.fill, backgroundColor);
      else if (backgroundColor && isAnchor) {
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

      if (cell && (!alignmentSpan || alignmentSpan.isAnchor)) {
        deferredContent.push({ address, cell, contentRect, spanRect });
      }
  }
  for (const { address, cell, contentRect, spanRect } of deferredContent) {
    if (cell.presentation?.kind === 'barcode') drawBarcodePresentation(context, contentRect, resolveDisplayText(cell), cell.presentation);
    else if (cell.presentation?.kind === 'image') drawCellImagePresentation(context, contentRect, cell.presentation, options);
    else drawCellValue(context, skeleton, renderOptions, address, cell, contentRect);
    if (cell.editor?.kind === 'checkbox') drawCheckboxEditor(context, spanRect, checkboxStateFromValue(cell.editor, cell.value ?? null));
    if (cell.overlay?.icon) drawTrendIcon(context, spanRect, cell.overlay.icon);
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

function drawCellImagePresentation(context: CanvasRenderingContext2D, rect: Rect, presentation: Extract<NonNullable<CellRenderData['presentation']>, { kind: 'image' }>, options: PaneDrawOptions): void {
  if (typeof Image === 'undefined') return;
  const assetId = presentation.asset.assetId;
  const assetUrl = options.assetUrlCache?.get(assetId);
  if (!assetUrl) {
    if (options.resolveAssetUrl && options.assetUrlPending && !options.assetUrlPending.has(assetId) && !options.assetUrlErrors?.has(assetId)) {
      options.assetUrlPending.add(assetId);
      void options.resolveAssetUrl(presentation.asset)
        .then((url) => options.assetUrlCache?.set(assetId, url))
        .catch((error) => options.assetUrlErrors?.set(assetId, error instanceof Error ? error.message : `ASSET_RESOLVE_FAILED: ${assetId}`))
        .finally(() => {
          options.assetUrlPending?.delete(assetId);
          options.requestRender?.();
        });
    }
    return;
  }
  let image = cellImageCache.get(assetId);
  if (!image) {
    image = new Image();
    image.src = assetUrl;
    cellImageCache.set(assetId, image);
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
  const crop = presentation.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
  const sourceX = image.naturalWidth * crop.left;
  const sourceY = image.naturalHeight * crop.top;
  const sourceWidth = image.naturalWidth * (1 - crop.left - crop.right);
  const sourceHeight = image.naturalHeight * (1 - crop.top - crop.bottom);
  const effects = presentation.effects;
  context.globalAlpha = 1 - (effects?.transparency ?? 0);
  context.filter = `brightness(${1 + (effects?.brightness ?? 0)}) contrast(${1 + (effects?.contrast ?? 0)})`;
  context.drawImage(image, sourceX, sourceY, sourceWidth, sourceHeight, x, y, width, height);
  context.restore();
}

function drawCheckboxEditor(context: CanvasRenderingContext2D, rect: Rect, state: CheckboxCellState | null): void {
  const size = Math.min(14, Math.max(10, rect.height - 8));
  const x = rect.x + 4;
  const y = rect.y + (rect.height - size) / 2;
  context.save();
  context.fillStyle = state === null ? '#fff7ed' : '#ffffff';
  context.strokeStyle = state === null ? '#c2410c' : state === 'checked' ? '#217345' : '#94a3b8';
  context.lineWidth = 1;
  context.fillRect(x, y, size, size);
  context.strokeRect(x + 0.5, y + 0.5, size - 1, size - 1);
  if (state === 'checked') {
    context.strokeStyle = '#217345';
    context.lineWidth = 2;
    context.beginPath();
    context.moveTo(x + 2.5, y + size / 2);
    context.lineTo(x + size / 2 - 1, y + size - 3);
    context.lineTo(x + size - 2, y + 3);
    context.stroke();
  } else if (state === 'indeterminate') {
    context.fillStyle = '#64748b';
    context.fillRect(x + 3, y + size / 2 - 1, size - 6, 2);
  } else if (state === null) {
    context.fillStyle = '#c2410c';
    context.font = `${Math.max(8, size - 2)}px sans-serif`;
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.fillText('?', x + size / 2, y + size / 2 + 0.5);
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

function drawCellFill(
  context: CanvasRenderingContext2D,
  rect: Rect,
  fill: NonNullable<CellRenderData['style']>['fill'],
  fallback: string | undefined,
): void {
  if (!fill) return;
  context.save();
  if (fill.kind === 'solid') {
    context.fillStyle = fill.foreground ?? fallback ?? '#ffffff';
    context.fillRect(rect.x, rect.y, rect.width, rect.height);
    context.restore();
    return;
  }
  if (fill.kind === 'gradient') {
    const degree = ((fill.degree ?? 0) * Math.PI) / 180;
    const length = Math.hypot(rect.width, rect.height);
    const centerX = rect.x + rect.width / 2;
    const centerY = rect.y + rect.height / 2;
    const gradient = fill.gradientType === 'path'
      ? context.createRadialGradient(centerX, centerY, 0, centerX, centerY, Math.max(rect.width, rect.height) / 2)
      : context.createLinearGradient(centerX - Math.cos(degree) * length / 2, centerY - Math.sin(degree) * length / 2, centerX + Math.cos(degree) * length / 2, centerY + Math.sin(degree) * length / 2);
    const stops = [...(fill.stops ?? [])].sort((left, right) => left.position - right.position);
    if (stops.length === 0) {
      gradient.addColorStop(0, fill.foreground ?? fallback ?? '#ffffff');
      gradient.addColorStop(1, fill.background ?? fallback ?? '#ffffff');
    } else {
      for (const stop of stops) gradient.addColorStop(Math.max(0, Math.min(1, stop.position)), stop.color);
    }
    context.fillStyle = gradient;
    context.fillRect(rect.x, rect.y, rect.width, rect.height);
    context.restore();
    return;
  }
  context.fillStyle = fill.background ?? fallback ?? '#ffffff';
  context.fillRect(rect.x, rect.y, rect.width, rect.height);
  context.beginPath();
  context.rect(rect.x, rect.y, rect.width, rect.height);
  context.clip();
  const color = fill.foreground ?? '#94a3b8';
  context.strokeStyle = color;
  context.fillStyle = color;
  context.lineWidth = 1;
  context.beginPath();
  const spacing = Math.max(4, Math.min(10, Math.round(Math.min(rect.width, rect.height) / 2)));
  switch (fill.pattern ?? 'solid') {
    case 'darkDown': case 'lightDown':
      for (let x = rect.x - rect.height; x < rect.x + rect.width; x += spacing) { context.moveTo(x, rect.y); context.lineTo(x + rect.height, rect.y + rect.height); }
      break;
    case 'darkUp': case 'lightUp':
      for (let x = rect.x; x < rect.x + rect.width + rect.height; x += spacing) { context.moveTo(x, rect.y + rect.height); context.lineTo(x - rect.height, rect.y); }
      break;
    case 'darkGrid': case 'lightGrid':
      for (let x = rect.x; x <= rect.x + rect.width; x += spacing) { context.moveTo(x, rect.y); context.lineTo(x, rect.y + rect.height); }
      for (let y = rect.y; y <= rect.y + rect.height; y += spacing) { context.moveTo(rect.x, y); context.lineTo(rect.x + rect.width, y); }
      break;
    case 'darkTrellis': case 'lightTrellis':
      for (let x = rect.x - rect.height; x < rect.x + rect.width; x += spacing) { context.moveTo(x, rect.y); context.lineTo(x + rect.height, rect.y + rect.height); context.moveTo(x, rect.y + rect.height); context.lineTo(x + rect.height, rect.y); }
      break;
    case 'gray0625':
      for (let x = rect.x + 2; x < rect.x + rect.width; x += spacing) for (let y = rect.y + 2; y < rect.y + rect.height; y += spacing) context.fillRect(x, y, 1, 1);
      break;
    case 'darkGray': case 'mediumGray': case 'lightGray':
      context.globalAlpha = fill.pattern === 'darkGray' ? 0.55 : fill.pattern === 'mediumGray' ? 0.35 : 0.2;
      context.fillRect(rect.x, rect.y, rect.width, rect.height);
      break;
    case 'solid': case 'none': case 'gray125':
      break;
  }
  context.stroke();
  context.restore();
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
    context.setLineDash(border.style === "dashed" ? [4, 3] : border.style === "dotted" || border.style === "hair" ? [1, 2] : border.style === "dashDot" ? [5, 2, 1, 2] : border.style === "dashDotDot" ? [5, 2, 1, 2, 1, 2] : []);
    context.lineWidth = border.style === "thick" ? 2 : border.style === "medium" ? 1.5 : 1;
    context.beginPath();
    context.moveTo(coordinates[0], coordinates[1]);
    context.lineTo(coordinates[2], coordinates[3]);
    context.stroke();
  }
  const diagonal = borders.diagonal;
  if (diagonal && (borders.diagonalUp || borders.diagonalDown)) {
    context.strokeStyle = diagonal.color;
    context.setLineDash(diagonal.style === 'dashed' ? [4, 3] : diagonal.style === 'dotted' ? [1, 2] : []);
    context.lineWidth = diagonal.style === 'thick' ? 2 : 1;
    if (borders.diagonalUp) {
      context.beginPath(); context.moveTo(rect.x, rect.y + rect.height); context.lineTo(rect.x + rect.width, rect.y); context.stroke();
    }
    if (borders.diagonalDown) {
      context.beginPath(); context.moveTo(rect.x, rect.y); context.lineTo(rect.x + rect.width, rect.y + rect.height); context.stroke();
    }
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

/**
 * AutoFit consumes the already resolved render projection.  Do not use JS
 * truthiness here: numeric zero and FALSE are visible content, while the
 * canonical empty display is the empty string.
 */
export function hasMeasurableCellContent(
  cell: { value?: unknown; displayValue?: unknown } | undefined,
): boolean {
  if (!cell) return false;
  const displayValue = cell.displayValue ?? cell.value;
  if (displayValue === undefined || displayValue === null) return false;
  return typeof displayValue === 'string' ? displayValue.length > 0 : true;
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
  const layout = resolveCellContentLayout({
    context,
    cell,
    theme,
    text,
    cellRect: { x: 0, y: 0, width: availableWidthPx ?? 0, height: 20 },
    mode: 'display',
    reserveFilterButton,
  });
  if (!cell.phonetic?.visible || cell.phonetic.runs.length === 0) return { widthPx: layout.widthPx, heightPx: layout.heightPx };
  const phoneticText = cell.phonetic.runs.map((run) => run.text).join('');
  const phoneticSize = cell.phonetic.fontSizePx ?? Math.max(7, layout.fontSizePx * 0.55);
  context.save();
  context.font = `${phoneticSize}px ${cell.phonetic.fontFamily ?? cell.style?.fontFamily ?? 'Microsoft YaHei'}, sans-serif`;
  const phoneticWidth = context.measureText(phoneticText).width + theme.cellPadding * 2;
  context.restore();
  return { widthPx: Math.max(layout.widthPx, phoneticWidth), heightPx: layout.heightPx + phoneticSize * 1.25 };
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
  if (style?.unsupportedAlignment?.horizontal || style?.unsupportedAlignment?.vertical) {
    drawUnsupportedAlignment(context, rect, theme);
    return;
  }

  const padding = style?.padding ?? theme.cellPadding;
  const indent = Math.max(0, Math.trunc(style?.indent ?? 0)) * 12;
  const cellRange = cell.merge?.isAnchor
    ? { startColumn: cell.merge.startColumn, endColumn: cell.merge.endColumn }
    : cell.alignmentSpan?.isAnchor
      ? { startColumn: cell.alignmentSpan.startColumn, endColumn: cell.alignmentSpan.endColumn }
      : { startColumn: address.column, endColumn: address.column };
  const layout = resolveCellContentLayout({
    context,
    cell,
    theme,
    text,
    cellRect: rect,
    mode: 'display',
    ...(cell.alignmentSpan?.isAnchor ? { alignmentSpan: rect } : {}),
    cellRange,
    neighborOccupancy: layoutNeighbors(options, address, cellRange),
  });
  const hAlign = layout.horizontalAlignment;
  const vAlign = style?.verticalAlignment ?? "middle";
  const paintRect = layout.displayRect;

  context.save();
  context.font = layout.font;
  context.fillStyle = style?.textColor ?? (cell.hyperlink ? '#0563c1' : theme.cellText);
  context.textBaseline = "middle";
  const phoneticHeight = drawPhoneticGuide(context, cell, paintRect, layout.fontSizePx, style?.fontFamily, style?.textColor ?? theme.cellText);

  const wrap = Boolean(style?.wrapText);
  const measured = layout.rawTextWidthPx;

  if (style?.textOrientation === 'stacked') {
    drawStackedText(context, layout.lines, layout.lineHeightPx, paintRect, padding + indent, vAlign);
    context.restore();
    return;
  }

  if (wrap || layout.lines.length > 1) {
    drawWrapped(context, layout.lines, layout.lineHeightPx, paintRect, padding + indent, hAlign, vAlign, layout.overflowWidthPx);
    context.restore();
    return;
  }

  context.beginPath();
  context.rect(paintRect.x, paintRect.y, paintRect.width, paintRect.height);
  context.clip();
  let x: number;
  if (hAlign === "center" || hAlign === "centerContinuous") x = paintRect.x + paintRect.width / 2;
  else if (hAlign === "right") x = paintRect.x + paintRect.width - padding - indent;
  else x = paintRect.x + padding + indent;

  const fontSize = layout.fontSizePx;
  let y = vAlign === 'top'
    ? paintRect.y + padding + phoneticHeight + fontSize / 2
    : vAlign === 'bottom'
      ? paintRect.y + paintRect.height - padding - fontSize / 2
      : paintRect.y + paintRect.height / 2 + phoneticHeight * 0.25;
  const rotate = resolveTextRotationDegrees(style);
  if (rotate !== 0) {
    const radians = (rotate * Math.PI) / 180;
    context.translate(paintRect.x + paintRect.width / 2, paintRect.y + paintRect.height / 2);
    context.rotate(-radians);
    x = 0;
    y = 0;
  }

  if (cell.richText?.length && rotate === 0 && layout.lines.length === 1 && hAlign !== 'fill' && hAlign !== 'justify' && hAlign !== 'distributed') drawRichTextRuns(context, cell.richText, layout.fontRuns, paintRect, padding, indent, hAlign, y);
  else if (hAlign === 'fill') drawFilledText(context, text, paintRect.x + padding + indent, y, layout.overflowWidthPx);
  else if (hAlign === 'justify' || hAlign === 'distributed') drawDistributedText(context, text, paintRect.x + padding + indent, y, layout.overflowWidthPx, hAlign === 'distributed');
  else {
    const canvasAlign = canvasTextAlign(hAlign);
    context.textAlign = canvasAlign;
    context.fillText(text, x, y);
  }

  if (style?.underline || style?.strikethrough || cell.hyperlink) {
    const textWidth = Math.min(measured, layout.overflowWidthPx);
    const canvasAlign = canvasTextAlign(hAlign);
    let lineX1 = canvasAlign === "center" ? x - textWidth / 2 : canvasAlign === "right" ? x - textWidth : x;
    const lineX2 = lineX1 + textWidth;
    context.strokeStyle = style?.textColor ?? (cell.hyperlink ? '#0563c1' : theme.cellText);
    context.lineWidth = style?.underlineStyle === 'double' || style?.underlineStyle === 'doubleAccounting' ? 0.75 : 1;
    context.beginPath();
    if (style?.underline || cell.hyperlink) {
      const lineY = y + fontSize * 0.45;
      const accounting = style?.underlineStyle === 'singleAccounting' || style?.underlineStyle === 'doubleAccounting';
      context.moveTo(lineX1, lineY);
      context.lineTo(lineX2, lineY);
      if (style?.underlineStyle === 'double' || style?.underlineStyle === 'doubleAccounting') {
        const secondLineY = lineY + (accounting ? 2.5 : 2);
        context.moveTo(lineX1, secondLineY);
        context.lineTo(lineX2, secondLineY);
      }
    }
    if (style?.strikethrough) {
      context.moveTo(lineX1, y);
      context.lineTo(lineX2, y);
    }
    context.stroke();
  }

  context.restore();
  void skeleton;
}

function drawRichTextRuns(
  context: CanvasRenderingContext2D,
  runs: readonly import('@react-sheets/core-model').RichTextRun[],
  measuredRuns: readonly import('./cell-content-layout').CellMeasuredTextRun[],
  rect: Rect,
  padding: number,
  indent: number,
  alignment: import('@react-sheets/core-model').HorizontalAlignment,
  y: number,
): void {
  const totalWidth = measuredRuns.reduce((sum, run) => sum + run.widthPx, 0);
  const start = alignment === 'right'
    ? rect.x + rect.width - padding - indent - totalWidth
    : alignment === 'center' || alignment === 'centerContinuous'
      ? rect.x + rect.width / 2 - totalWidth / 2
      : rect.x + padding + indent;
  let x = start;
  context.textAlign = 'left';
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index]!;
    const measured = measuredRuns[index];
    if (!measured) continue;
    context.font = measured.font;
    context.fillStyle = run.style?.textColor ?? context.fillStyle;
    context.fillText(run.text, x, y);
    x += measured.widthPx;
  }
}

function layoutNeighbors(options: PaneDrawOptions, address: CellAddress, range: { startColumn: number; endColumn: number }): { left: CellLayoutNeighbor[]; right: CellLayoutNeighbor[] } {
  const cacheKey = `${address.row}:${range.startColumn}:${range.endColumn}`;
  const cached = options.neighborCache?.get(cacheKey);
  if (cached) return cached;
  const occupied = options.rowOccupancyCache?.get(address.row) ?? buildRowOccupancy(options, address.row);
  options.rowOccupancyCache?.set(address.row, occupied);
  const left: CellLayoutNeighbor[] = [];
  for (let column = range.startColumn - 1; column >= 0; column -= 1) {
    left.push({ column, widthPx: options.skeleton.getColumnWidth(column), occupied: occupied.has(column) });
    if (left.at(-1)?.occupied) break;
  }
  const right: CellLayoutNeighbor[] = [];
  for (let column = range.endColumn + 1; column < options.skeleton.columnCount; column += 1) {
    right.push({ column, widthPx: options.skeleton.getColumnWidth(column), occupied: occupied.has(column) });
    if (right.at(-1)?.occupied) break;
  }
  const result = { left, right };
  options.neighborCache?.set(cacheKey, result);
  return result;
}

function buildRowOccupancy(options: PaneDrawOptions, row: number): ReadonlySet<number> {
  const occupied = new Set<number>();
  for (let column = 0; column < options.skeleton.columnCount; column += 1) {
    if (hasRenderableCellContent(options.cellProvider({ row, column }))) occupied.add(column);
  }
  return occupied;
}

function hasRenderableCellContent(cell: CellRenderData | undefined): boolean {
  if (!cell) return false;
  if (cell.presentation || cell.editor || cell.formula) return true;
  const value = resolveDisplayText(cell);
  return value.length > 0;
}

function drawPhoneticGuide(context: CanvasRenderingContext2D, cell: CellRenderData, rect: Rect, baseFontSize: number, baseFontFamily: string | undefined, color: string): number {
  const metadata = cell.phonetic;
  if (!metadata?.visible || metadata.runs.length === 0) return 0;
  const text = metadata.runs.map((run) => run.text).join('');
  const size = metadata.fontSizePx ?? Math.max(7, baseFontSize * 0.55);
  context.save();
  context.font = `${size}px ${metadata.fontFamily ?? baseFontFamily ?? 'Microsoft YaHei'}, sans-serif`;
  context.fillStyle = color;
  context.textBaseline = 'top';
  context.textAlign = metadata.alignment === 'left' ? 'left' : 'center';
  const x = metadata.alignment === 'left' ? rect.x + 2 : rect.x + rect.width / 2;
  context.fillText(text, x, rect.y + 1, Math.max(1, rect.width - 4));
  context.restore();
  return size * 1.25;
}

function drawWrapped(
  context: CanvasRenderingContext2D,
  lines: readonly string[],
  lineHeight: number,
  rect: Rect,
  padding: number,
  hAlign: import('@react-sheets/core-model').HorizontalAlignment,
  vAlign: import('@react-sheets/core-model').VerticalAlignment,
  maxWidth: number,
): void {
  const textHeight = lines.length * lineHeight;
  const startY = vAlign === "top"
    ? rect.y + padding + lineHeight / 2
    : vAlign === "bottom"
      ? rect.y + rect.height - padding - ((lines.length - 1) * lineHeight) - lineHeight / 2
      : vAlign === 'justify' || vAlign === 'distributed'
        ? rect.y + padding + lineHeight / 2
      : rect.y + rect.height / 2 - ((lines.length - 1) * lineHeight) / 2;

  const canvasAlign = canvasTextAlign(hAlign);
  context.textAlign = canvasAlign;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!;
    const y = vAlign === 'justify' || vAlign === 'distributed'
      ? rect.y + padding + lineHeight / 2 + (lines.length === 1 ? (rect.height - textHeight) / 2 : i * (rect.height - textHeight) / Math.max(1, lines.length - 1))
      : startY + i * lineHeight;
    if ((hAlign === 'justify' || hAlign === 'distributed') && i < lines.length - 1) {
      drawDistributedText(context, line, rect.x + padding, y, maxWidth, hAlign === 'distributed');
    } else {
      const x = canvasAlign === 'center' ? rect.x + rect.width / 2 : canvasAlign === 'right' ? rect.x + rect.width - padding : rect.x + padding;
      context.fillText(line, x, y, maxWidth);
    }
  }
}

function canvasTextAlign(
  alignment: import('@react-sheets/core-model').HorizontalAlignment,
): CanvasTextAlign {
  if (alignment === 'center' || alignment === 'centerContinuous') return 'center';
  if (alignment === 'right') return 'right';
  return 'left';
}

function drawFilledText(
  context: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  width: number,
): void {
  if (!text || width <= 0) return;
  context.textAlign = 'left';
  const measured = context.measureText(text).width;
  if (measured <= 0) return;
  const count = Math.max(1, Math.ceil(width / measured));
  context.fillText(text.repeat(count).slice(0, Math.max(text.length, Math.floor(count * text.length))), x, y, width);
}

function drawDistributedText(
  context: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  width: number,
  includeEdges: boolean,
): void {
  if (!text || width <= 0) return;
  const characters = Array.from(text);
  if (characters.length < 2) {
    context.textAlign = 'left';
    context.fillText(text, x, y, width);
    return;
  }
  const textWidth = context.measureText(text).width;
  const gap = Math.max(0, (width - textWidth) / (includeEdges ? characters.length + 1 : characters.length - 1));
  let cursor = x + (includeEdges ? gap : 0);
  context.textAlign = 'left';
  for (const character of characters) {
    context.fillText(character, cursor, y);
    cursor += context.measureText(character).width + gap;
  }
}

function drawStackedText(
  context: CanvasRenderingContext2D,
  characters: readonly string[],
  lineHeight: number,
  rect: Rect,
  padding: number,
  vAlign: import('@react-sheets/core-model').VerticalAlignment,
): void {
  if (characters.length === 0) return;
  const contentHeight = characters.length * lineHeight;
  const startY = vAlign === 'top'
    ? rect.y + padding + lineHeight / 2
    : vAlign === 'bottom'
      ? rect.y + rect.height - padding - contentHeight + lineHeight / 2
      : rect.y + (rect.height - contentHeight) / 2 + lineHeight / 2;
  context.textAlign = 'center';
  for (let index = 0; index < characters.length; index += 1) {
    context.fillText(characters[index]!, rect.x + rect.width / 2, startY + index * lineHeight);
  }
}

function drawUnsupportedAlignment(context: CanvasRenderingContext2D, rect: Rect, theme: RenderTheme): void {
  context.save();
  context.strokeStyle = theme.invalidColor;
  context.lineWidth = 1;
  context.strokeRect(rect.x + 1, rect.y + 1, Math.max(0, rect.width - 2), Math.max(0, rect.height - 2));
  context.fillStyle = theme.invalidColor;
  context.font = '10px "Microsoft YaHei", "Segoe UI", sans-serif';
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  context.fillText('Unsupported alignment', rect.x + rect.width / 2, rect.y + rect.height / 2, Math.max(1, rect.width - 4));
  context.restore();
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
