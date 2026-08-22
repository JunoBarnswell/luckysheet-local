import { formatValue } from "@react-sheets/number-format";
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
  const background = { x: pane.offset.x, y: pane.offset.y, width: pane.rect.width, height: pane.rect.height };
  context.fillStyle = theme.canvasBackground;
  context.fillRect(background.x, background.y, background.width, background.height);
  if (!visibleRange) return;

  // 合并覆盖集合:被合并覆盖的非锚点单元格之间不画内部网格线
  const covered = collectMergeCovered(options, visibleRange);

  context.strokeStyle = theme.gridLine;
  context.lineWidth = 1;
  context.beginPath();
  for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
    if (!shouldDrawRect(drawRects, rowStrip(skeleton, visibleRange, row))) continue;
    const y = Math.round(skeleton.getRowTop(row)) + 0.5;
    for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column++) {
      if (covered.has(row + ":" + column)) continue;
      const x = Math.round(skeleton.getColumnLeft(column)) + 0.5;
      const width = skeleton.getColumnWidth(column);
      const height = skeleton.getRowHeight(row);
      context.moveTo(x, y);
      context.lineTo(x + width, y);
      context.moveTo(x, y + height);
      context.lineTo(x + width, y + height);
      break; // 行上下边线每行画一次即可(由最左列起)
    }
  }
  for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column++) {
    const x = Math.round(skeleton.getColumnLeft(column)) + 0.5;
    let blockedAbove = false;
    let blockedBelow = false;
    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
      if (covered.has(row + ":" + column)) { blockedAbove = true; blockedBelow = true; continue; }
      const height = skeleton.getRowHeight(row);
      const y = Math.round(skeleton.getRowTop(row)) + 0.5;
      if (!blockedAbove) { context.moveTo(x, y); context.lineTo(x + skeleton.getColumnWidth(column), y); }
      if (!blockedBelow) { context.moveTo(x, y + height); context.lineTo(x + skeleton.getColumnWidth(column), y + height); }
      break;
    }
  }
  // 右侧与底部收尾边界线
  const endX = Math.round(skeleton.getColumnLeft(visibleRange.endColumn) + skeleton.getColumnWidth(visibleRange.endColumn)) + 0.5;
  const endY = Math.round(skeleton.getRowTop(visibleRange.endRow) + skeleton.getRowHeight(visibleRange.endRow)) + 0.5;
  context.moveTo(endX, skeleton.getRowTop(visibleRange.startRow));
  context.lineTo(endX, endY - 0.5);
  context.moveTo(skeleton.getColumnLeft(visibleRange.startColumn), endY);
  context.lineTo(endX - 0.5, endY);
  context.stroke();

  void background;
}

function rowStrip(skeleton: SheetSkeleton, range: CellRange, row: number): Rect {
  return {
    x: skeleton.getColumnLeft(range.startColumn),
    y: skeleton.getRowTop(row),
    width: skeleton.getColumnLeft(range.endColumn) + skeleton.getColumnWidth(range.endColumn) - skeleton.getColumnLeft(range.startColumn),
    height: skeleton.getRowHeight(row),
  };
}

interface MergeSpanLike { startRow: number; endRow: number; startColumn: number; endColumn: number; isAnchor: boolean }

function collectMergeCovered(options: PaneDrawOptions, range: CellRange): Set<string> {
  const covered = new Set<string>();
  const { cellProvider } = options;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      const cell = cellProvider({ row, column });
      const merge = cell?.merge;
      if (!merge || !merge.isAnchor) continue;
      for (let r = merge.startRow; r <= merge.endRow; r++) {
        for (let c = merge.startColumn; c <= merge.endColumn; c++) {
          if (r === merge.startRow && c === merge.startColumn) continue;
          covered.add(r + ":" + c);
        }
      }
    }
  }
  return covered;
}

// ---------------- 内容层 ----------------

export function drawCellLayer(options: PaneDrawOptions): void {
  const { context, skeleton, visibleRange, cellProvider, theme, drawRects } = options;
  if (!visibleRange) return;
  context.textBaseline = "middle";

  for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
    for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column++) {
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
        drawCellValue(context, skeleton, options, address, cell, spanRect);
        if (cell.overlay?.icon) drawTrendIcon(context, spanRect, cell.overlay.icon);
      }
    }
  }
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

function resolveDisplayText(cell: CellRenderData): string {
  if (cell.displayValue !== undefined) return cell.displayValue;
  if (typeof cell.value === "number") {
    return formatValue(cell.value, cell.style?.numberFormat);
  }
  if (cell.value == null) return "";
  if (typeof cell.value === "boolean") return cell.value ? "TRUE" : "FALSE";
  return String(cell.value);
}

function fontOf(style: CellRenderData["style"], theme: RenderTheme): string {
  const size = style?.fontSize ?? 13;
  const family = style?.fontFamily ? '"' + style.fontFamily + '", sans-serif' : "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
  const weight = style?.bold ? "700" : "400";
  const slant = style?.italic ? " italic" : "";
  return slant + " " + weight + " " + size + "px " + family;
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
  const hAlign = style?.horizontalAlignment ?? "left";
  const vAlign = style?.verticalAlignment ?? "middle";

  context.save();
  context.font = fontOf(style, theme);
  context.fillStyle = style?.textColor ?? theme.cellText;
  context.textBaseline = "middle";

  const wrap = Boolean(style?.wrapText);
  const maxWidth = rect.width - padding * 2;
  const measured = context.measureText(text).width;

  if (wrap) {
    drawWrapped(context, text, rect, padding, hAlign, vAlign, maxWidth);
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
  else if (hAlign === "right") x = rect.x + rect.width - padding;
  else x = rect.x + padding;

  let y = rect.y + rect.height / 2;
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
  const words = text.split(/\s+/);
  const lineHeight = 14;
  const lines: string[] = [];
  let currentLine = "";
  for (const word of words) {
    const attempt = currentLine ? currentLine + " " + word : word;
    if (context.measureText(attempt).width > maxWidth && currentLine) {
      lines.push(currentLine);
      currentLine = word;
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
  const paneContent = { x: pane.offset.x, y: pane.offset.y, width: pane.rect.width, height: pane.rect.height };
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