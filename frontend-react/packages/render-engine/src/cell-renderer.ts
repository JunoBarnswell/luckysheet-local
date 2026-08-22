import { intersectRect } from './geometry';
import { SheetSkeleton } from './sheet-skeleton';
import {
  DEFAULT_RENDER_THEME,
  type BorderStyle,
  type CellBorders,
  type CellProvider,
  type CellRenderData,
  type CellRenderStyle,
  type CellRange,
  type Rect,
  type RenderTheme,
  type ViewportSnapshot,
} from './types';

function formatCellValue(cell: CellRenderData): string {
  if (cell.displayValue !== undefined) return cell.displayValue;
  if (cell.error) return cell.error;
  if (cell.value === null || cell.value === undefined) return '';

  if (typeof cell.value === 'number') {
    const format = cell.style?.numberFormat;
    if (format === '0%') {
      return `${Math.round(cell.value * 100)}%`;
    }
    if (format === '0.00%') {
      return `${(cell.value * 100).toFixed(2)}%`;
    }
    if (format === '$#,##0') {
      return `$${cell.value.toLocaleString('en-US', { maximumFractionDigits: 0 })}`;
    }
    if (format === '$#,##0.00') {
      return `$${cell.value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    if (format === '#,##0') {
      return cell.value.toLocaleString('en-US', { maximumFractionDigits: 0 });
    }
    if (format === '#,##0.00') {
      return cell.value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(cell.value);
  }

  if (typeof cell.value === 'boolean') {
    return cell.value ? 'TRUE' : 'FALSE';
  }

  return String(cell.value);
}

function withClip(
  context: CanvasRenderingContext2D,
  rects: readonly Rect[] | undefined,
  callback: () => void,
): void {
  if (!rects || rects.length === 0) {
    callback();
    return;
  }
  context.save();
  context.beginPath();
  for (const rect of rects) context.rect(rect.x, rect.y, rect.width, rect.height);
  context.clip();
  try {
    callback();
  } finally {
    context.restore();
  }
}

function viewportCellRect(
  skeleton: SheetSkeleton,
  viewport: ViewportSnapshot,
  row: number,
  column: number,
): Rect | null {
  const sheetRect = skeleton.getCellRect(row, column);
  if (!sheetRect) return null;
  return {
    x: sheetRect.x - viewport.scrollX,
    y: sheetRect.y - viewport.scrollY,
    width: sheetRect.width,
    height: sheetRect.height,
  };
}

function viewportRangeRect(
  skeleton: SheetSkeleton,
  viewport: ViewportSnapshot,
  range: CellRange,
): Rect | null {
  const sheetRect = skeleton.getRangeRect(range);
  if (!sheetRect) return null;
  return {
    x: sheetRect.x - viewport.scrollX,
    y: sheetRect.y - viewport.scrollY,
    width: sheetRect.width,
    height: sheetRect.height,
  };
}

function cellIsInDrawRegion(cellRect: Rect, drawRects: readonly Rect[] | undefined): boolean {
  if (!drawRects || drawRects.length === 0) return true;
  return drawRects.some((drawRect) => intersectRect(cellRect, drawRect) !== null);
}

export interface CellDrawOptions {
  context: CanvasRenderingContext2D;
  skeleton: SheetSkeleton;
  viewport: ViewportSnapshot;
  visibleRange: CellRange | null;
  cellProvider: CellProvider;
  theme?: RenderTheme;
  drawRects?: readonly Rect[];
}

export function drawGridLayer(options: CellDrawOptions): void {
  const theme = options.theme ?? DEFAULT_RENDER_THEME;
  const visibleRange = options.visibleRange;
  if (!visibleRange) return;
  const { context, skeleton, viewport } = options;

  withClip(context, options.drawRects, () => {
    context.fillStyle = theme.canvasBackground;
    context.fillRect(0, 0, viewport.width, viewport.height);

    // 1. Draw cell backgrounds
    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row += 1) {
      for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column += 1) {
        const cell = options.cellProvider({ row, column });
        const isMergedNonAnchor = cell?.merge && !cell.merge.isAnchor;
        if (isMergedNonAnchor) continue;

        let cellRect: Rect | null;
        if (cell?.merge?.isAnchor) {
          cellRect = viewportRangeRect(skeleton, viewport, {
            startRow: cell.merge.startRow,
            endRow: cell.merge.endRow,
            startColumn: cell.merge.startColumn,
            endColumn: cell.merge.endColumn,
          });
        } else {
          cellRect = viewportCellRect(skeleton, viewport, row, column);
        }

        if (!cellRect || !cellIsInDrawRegion(cellRect, options.drawRects)) continue;

        const background = cell?.style?.background;
        if (background) {
          context.fillStyle = background;
          context.fillRect(cellRect.x, cellRect.y, cellRect.width, cellRect.height);
        }
      }
    }

    // 2. Draw standard grid lines
    const left = skeleton.getColumnLeft(visibleRange.startColumn) - viewport.scrollX;
    const right =
      skeleton.getColumnLeft(visibleRange.endColumn) +
      skeleton.getColumnWidth(visibleRange.endColumn) -
      viewport.scrollX;
    const top = skeleton.getRowTop(visibleRange.startRow) - viewport.scrollY;
    const bottom =
      skeleton.getRowTop(visibleRange.endRow) +
      skeleton.getRowHeight(visibleRange.endRow) -
      viewport.scrollY;

    context.beginPath();
    for (let row = visibleRange.startRow; row <= visibleRange.endRow + 1; row += 1) {
      const y =
        row <= visibleRange.endRow ? skeleton.getRowTop(row) - viewport.scrollY : bottom;
      context.moveTo(left, y);
      context.lineTo(right, y);
    }
    for (let column = visibleRange.startColumn; column <= visibleRange.endColumn + 1; column += 1) {
      const x =
        column <= visibleRange.endColumn
          ? skeleton.getColumnLeft(column) - viewport.scrollX
          : right;
      context.moveTo(x, top);
      context.lineTo(x, bottom);
    }
    context.strokeStyle = theme.gridLine;
    context.lineWidth = 1;
    context.stroke();

    // 3. Draw custom borders
    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row += 1) {
      for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column += 1) {
        const cell = options.cellProvider({ row, column });
        if (!cell?.style?.borders) continue;
        const cellRect = viewportCellRect(skeleton, viewport, row, column);
        if (!cellRect || !cellIsInDrawRegion(cellRect, options.drawRects)) continue;
        drawCellBorders(context, cellRect, cell.style.borders);
      }
    }
  });
}

function drawCellBorders(context: CanvasRenderingContext2D, rect: Rect, borders: CellBorders): void {
  const drawSide = (side: BorderStyle | undefined, x1: number, y1: number, x2: number, y2: number) => {
    if (!side) return;
    context.save();
    context.strokeStyle = side.color;
    context.lineWidth = side.style === 'thick' ? 3 : side.style === 'medium' ? 2 : 1;
    if (side.style === 'dashed') context.setLineDash([4, 2]);
    context.beginPath();
    context.moveTo(x1, y1);
    context.lineTo(x2, y2);
    context.stroke();
    context.restore();
  };

  drawSide(borders.top, rect.x, rect.y, rect.x + rect.width, rect.y);
  drawSide(borders.bottom, rect.x, rect.y + rect.height, rect.x + rect.width, rect.y + rect.height);
  drawSide(borders.left, rect.x, rect.y, rect.x, rect.y + rect.height);
  drawSide(borders.right, rect.x + rect.width, rect.y, rect.x + rect.width, rect.y + rect.height);
}

export function drawCellLayer(options: CellDrawOptions): void {
  const visibleRange = options.visibleRange;
  if (!visibleRange) return;
  const theme = options.theme ?? DEFAULT_RENDER_THEME;
  const { context, skeleton, viewport } = options;

  withClip(context, options.drawRects, () => {
    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row += 1) {
      for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column += 1) {
        const cell = options.cellProvider({ row, column });
        if (!cell) continue;

        // Skip non-anchor cells in merged areas
        if (cell.merge && !cell.merge.isAnchor) continue;

        const value = formatCellValue(cell);
        if (value.length === 0) continue;

        let cellRect: Rect | null;
        if (cell.merge?.isAnchor) {
          cellRect = viewportRangeRect(skeleton, viewport, {
            startRow: cell.merge.startRow,
            endRow: cell.merge.endRow,
            startColumn: cell.merge.startColumn,
            endColumn: cell.merge.endColumn,
          });
        } else {
          cellRect = viewportCellRect(skeleton, viewport, row, column);
        }

        if (!cellRect || !cellIsInDrawRegion(cellRect, options.drawRects)) continue;

        drawCellContent(context, value, cellRect, cell.style, theme);
      }
    }
  });
}

function resolveFont(style: CellRenderStyle | undefined, defaultFont: string): string {
  if (style?.font) return style.font;
  const weight = style?.bold ? 'bold' : 'normal';
  const slant = style?.italic ? 'italic' : 'normal';
  const size = style?.fontSize ?? 13;
  const family = style?.fontFamily ?? 'Inter, -apple-system, sans-serif';
  return `${slant} ${weight} ${size}px ${family}`;
}

function drawCellContent(
  context: CanvasRenderingContext2D,
  value: string,
  rect: Rect,
  style: CellRenderStyle | undefined,
  theme: RenderTheme,
): void {
  const padding = Math.max(0, style?.padding ?? theme.cellPadding);
  const hAlign = style?.horizontalAlignment ?? 'left';
  const vAlign = style?.verticalAlignment ?? 'middle';
  const font = resolveFont(style, theme.cellFont);
  const textColor = style?.textColor ?? theme.cellText;

  context.save();
  context.beginPath();
  context.rect(rect.x, rect.y, rect.width, rect.height);
  context.clip();

  context.font = font;
  context.fillStyle = textColor;
  context.textAlign = hAlign;
  context.textBaseline = vAlign === 'middle' ? 'middle' : vAlign === 'top' ? 'top' : 'bottom';

  const x =
    hAlign === 'left'
      ? rect.x + padding
      : hAlign === 'right'
        ? rect.x + rect.width - padding
        : rect.x + rect.width / 2;

  const y =
    vAlign === 'top'
      ? rect.y + padding
      : vAlign === 'bottom'
        ? rect.y + rect.height - padding
        : rect.y + rect.height / 2;

  const maxWidth = Math.max(0, rect.width - padding * 2);

  if (style?.wrapText) {
    const words = value.split(' ');
    let currentLine = '';
    const lines: string[] = [];
    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const metrics = context.measureText(testLine);
      if (metrics.width > maxWidth && currentLine) {
        lines.push(currentLine);
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    }
    if (currentLine) lines.push(currentLine);

    const lineHeight = 16;
    const totalHeight = lines.length * lineHeight;
    let startY = rect.y + (rect.height - totalHeight) / 2 + lineHeight / 2;

    for (const line of lines) {
      context.fillText(line, x, startY, maxWidth);
      startY += lineHeight;
    }
  } else {
    context.fillText(value, x, y, maxWidth);
  }

  // Draw Underline or Strikethrough if specified
  if (style?.underline || style?.strikethrough) {
    const textWidth = Math.min(context.measureText(value).width, maxWidth);
    let lineX1 = x;
    if (hAlign === 'center') lineX1 = x - textWidth / 2;
    else if (hAlign === 'right') lineX1 = x - textWidth;
    const lineX2 = lineX1 + textWidth;

    context.strokeStyle = textColor;
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
}
