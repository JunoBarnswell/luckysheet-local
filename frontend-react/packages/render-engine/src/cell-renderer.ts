import { intersectRect } from './geometry';
import { SheetSkeleton } from './sheet-skeleton';
import {
  DEFAULT_RENDER_THEME,
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
  return intersectRect(
    {
      x: sheetRect.x - viewport.scrollX,
      y: sheetRect.y - viewport.scrollY,
      width: sheetRect.width,
      height: sheetRect.height,
    },
    { x: 0, y: 0, width: viewport.width, height: viewport.height },
  );
}

function cellIsInDrawRegion(cellRect: Rect, drawRects: readonly Rect[] | undefined): boolean {
  if (!drawRects || drawRects.length === 0) return true;
  return drawRects.some((drawRect) => intersectRect(cellRect, drawRect) !== null);
}

function cellStyle(style: CellRenderStyle | undefined, theme: RenderTheme): Required<Pick<
  CellRenderStyle,
  'background' | 'textColor' | 'font' | 'horizontalAlignment' | 'verticalAlignment' | 'padding'
>> {
  return {
    background: style?.background ?? theme.defaultCellBackground,
    textColor: style?.textColor ?? theme.cellText,
    font: style?.font ?? theme.cellFont,
    horizontalAlignment: style?.horizontalAlignment ?? 'left',
    verticalAlignment: style?.verticalAlignment ?? 'middle',
    padding: style?.padding ?? theme.cellPadding,
  };
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

    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row += 1) {
      for (let column = visibleRange.startColumn; column <= visibleRange.endColumn; column += 1) {
        const cellRect = viewportCellRect(skeleton, viewport, row, column);
        if (!cellRect || !cellIsInDrawRegion(cellRect, options.drawRects)) continue;
        const cell = options.cellProvider({ row, column });
        const background = cell?.style?.background;
        if (background) {
          context.fillStyle = background;
          context.fillRect(cellRect.x, cellRect.y, cellRect.width, cellRect.height);
        }
      }
    }

    const left = skeleton.getColumnLeft(visibleRange.startColumn) - viewport.scrollX;
    const right = skeleton.getColumnLeft(visibleRange.endColumn)
      + skeleton.getColumnWidth(visibleRange.endColumn)
      - viewport.scrollX;
    const top = skeleton.getRowTop(visibleRange.startRow) - viewport.scrollY;
    const bottom = skeleton.getRowTop(visibleRange.endRow)
      + skeleton.getRowHeight(visibleRange.endRow)
      - viewport.scrollY;

    context.beginPath();
    for (let row = visibleRange.startRow; row <= visibleRange.endRow + 1; row += 1) {
      const y = row <= visibleRange.endRow
        ? skeleton.getRowTop(row) - viewport.scrollY
        : bottom;
      context.moveTo(left, y);
      context.lineTo(right, y);
    }
    for (let column = visibleRange.startColumn; column <= visibleRange.endColumn + 1; column += 1) {
      const x = column <= visibleRange.endColumn
        ? skeleton.getColumnLeft(column) - viewport.scrollX
        : right;
      context.moveTo(x, top);
      context.lineTo(x, bottom);
    }
    context.strokeStyle = theme.gridLine;
    context.lineWidth = 1;
    context.stroke();
  });
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
        const value = formatCellValue(cell);
        if (value.length === 0) continue;
        const cellRect = viewportCellRect(skeleton, viewport, row, column);
        if (!cellRect || !cellIsInDrawRegion(cellRect, options.drawRects)) continue;
        drawCellText(context, value, cellRect, cell.style, theme);
      }
    }
  });
}

function drawCellText(
  context: CanvasRenderingContext2D,
  value: string,
  rect: Rect,
  style: CellRenderStyle | undefined,
  theme: RenderTheme,
): void {
  const resolved = cellStyle(style, theme);
  const padding = Math.max(0, resolved.padding);
  const baseline = resolved.verticalAlignment === 'top'
    ? rect.y + padding
    : resolved.verticalAlignment === 'bottom'
      ? rect.y + rect.height - padding
      : rect.y + rect.height / 2;
  const x = resolved.horizontalAlignment === 'left'
    ? rect.x + padding
    : resolved.horizontalAlignment === 'right'
      ? rect.x + rect.width - padding
      : rect.x + rect.width / 2;

  context.save();
  context.beginPath();
  context.rect(rect.x, rect.y, rect.width, rect.height);
  context.clip();
  context.font = resolved.font;
  context.fillStyle = resolved.textColor;
  context.textAlign = resolved.horizontalAlignment;
  context.textBaseline = resolved.verticalAlignment === 'middle' ? 'middle' : 'alphabetic';
  context.fillText(value, x, baseline, Math.max(0, rect.width - padding * 2));
  context.restore();
}
