import {
  COL_HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  type ChromeState,
  type Rect,
  type RenderPane,
  type RenderTheme,
} from "./types";
import { SheetSkeleton, columnLabelOf } from "./sheet-skeleton";
import { defaultHeaderOffset } from "./render-plan";
import type { RenderPlan } from "./render-plan";

export interface ChromeDrawOptions {
  context: CanvasRenderingContext2D;
  skeleton: SheetSkeleton;
  plan: RenderPlan;
  chrome: ChromeState;
  theme: RenderTheme;
}

const HEADER_FONT = "11px Inter, -apple-system, 'Segoe UI', Roboto, sans-serif";

/** 非滚动覆盖层:表头/选区/填充柄/冻结线/筛选漏斗/远端光标 */
export function drawChromeLayer(options: ChromeDrawOptions): void {
  const { context, plan, chrome, theme } = options;
  drawHeaderStrips(options);
  drawSelection(options);
  if (chrome.editing) drawEditingOutline(options, chrome.editing);
  drawFreezeTrapLines(context, plan);
  drawResizePreview(options);
  drawTableOutlines(options);
  drawOutlineControls(options);
  drawFilterFunnels(options);
  drawRemoteCursors(options);
  void theme;
}

function paneTransform(pane: RenderPane): { dx: number; dy: number } {
  return { dx: pane.screenRect.x - pane.contentOrigin.x, dy: pane.screenRect.y - pane.contentOrigin.y };
}

function contentRangeRect(skeleton: SheetSkeleton, range: { startRow: number; endRow: number; startColumn: number; endColumn: number }): Rect | null {
  return skeleton.getRangeRect(range);
}

function intersect(a: Rect, b: Rect): Rect | null {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  return right > x && bottom > y ? { x, y, width: right - x, height: bottom - y } : null;
}

function drawSelectionOutlines(
  context: CanvasRenderingContext2D,
  skeleton: SheetSkeleton,
  panes: readonly RenderPane[],
  ranges: readonly { startRow: number; endRow: number; startColumn: number; endColumn: number }[],
  strokeStyle: string,
  lineWidth: number,
): Array<{ rect: Rect; screen: Rect }> {
  const drawn: Array<{ rect: Rect; screen: Rect }> = [];
  for (const range of ranges) {
    const contentRect = contentRangeRect(skeleton, range);
    if (!contentRect) continue;
    for (const pane of panes) {
      if (pane.screenRect.width <= 0 || pane.screenRect.height <= 0) continue;
      const visible = intersect(
        { x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height },
        contentRect,
      );
      if (!visible) continue;
      const transform = paneTransform(pane);
      const screen = { x: visible.x + transform.dx, y: visible.y + transform.dy, width: visible.width, height: visible.height };
      context.strokeStyle = strokeStyle;
      context.lineWidth = lineWidth;
      context.strokeRect(screen.x + lineWidth / 2, screen.y + lineWidth / 2, screen.width - lineWidth, screen.height - lineWidth);
      drawn.push({ rect: contentRect, screen });
    }
  }
  return drawn;
}

function drawSelection(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome, theme } = options;
  const panes = plan.panes;
  const selectionRanges = chrome.selection.ranges.length > 0
    ? chrome.selection.ranges
    : [{ startRow: chrome.selection.primary.row, endRow: chrome.selection.primary.row, startColumn: chrome.selection.primary.column, endColumn: chrome.selection.primary.column }];

  // 选区底色
  for (const range of selectionRanges) {
    const contentRect = contentRangeRect(skeleton, range);
    if (!contentRect) continue;
    for (const pane of panes) {
      const visible = intersect({ x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height }, contentRect);
      if (!visible) continue;
      const t = paneTransform(pane);
      context.fillStyle = theme.selectionBackground;
      context.fillRect(visible.x + t.dx, visible.y + t.dy, visible.width, visible.height);
    }
  }

  // 主选区边框 + 次选区细框
  for (let i = 0; i < selectionRanges.length; i++) {
    const isPrimary = i === chrome.selection.primaryIndex;
    const drawn = drawSelectionOutlines(
      context,
      skeleton,
      panes,
      [selectionRanges[i]!],
      theme.selectionBorder,
      isPrimary ? 2 : 1,
    );
    if (isPrimary) {
      const primaryRange = selectionRanges[chrome.selection.primaryIndex] ?? selectionRanges[0]!;
      const primaryCell = { row: primaryRange.endRow, column: primaryRange.endColumn };
      const primaryRect = contentRangeRect(skeleton, {
        startRow: primaryCell.row,
        endRow: primaryCell.row,
        startColumn: primaryCell.column,
        endColumn: primaryCell.column,
      });
      if (primaryRect) {
        const targetPane = panes.find((pane) => pane.visibleRange
          && primaryCell.row >= pane.visibleRange.startRow && primaryCell.row <= pane.visibleRange.endRow
          && primaryCell.column >= pane.visibleRange.startColumn && primaryCell.column <= pane.visibleRange.endColumn)
          ?? panes.find((pane) => pane.id === "main")
          ?? panes.at(-1)!;
        const t = paneTransform(targetPane);
        const handleX = primaryRect.x + primaryRect.width + t.dx - theme.fillHandleSize / 2;
        const handleY = primaryRect.y + primaryRect.height + t.dy - theme.fillHandleSize / 2;
        context.fillStyle = theme.selectionBorder;
        context.strokeStyle = "#ffffff";
        context.lineWidth = 1;
        context.beginPath();
        context.rect(handleX, handleY, theme.fillHandleSize, theme.fillHandleSize);
        context.fill();
        context.stroke();
      }
      void drawn;
    }
  }
}

function drawEditingOutline(options: ChromeDrawOptions, editing: { row: number; column: number }): void {
  const { context, skeleton, plan, theme } = options;
  const rect = skeleton.getCellRect(editing.row, editing.column);
  if (!rect) return;
  for (const pane of plan.panes) {
    const visible = intersect({ x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height }, rect);
    if (!visible) continue;
    const t = paneTransform(pane);
    context.save();
    context.strokeStyle = theme.editingBorder;
    context.setLineDash([4, 3]);
    context.lineWidth = 1.5;
    context.strokeRect(visible.x + t.dx, visible.y + t.dy, visible.width, visible.height);
    context.restore();
  }
}

function drawFreezeTrapLines(context: CanvasRenderingContext2D, plan: RenderPlan): void {
  const leftBoundary = plan.panes.find((pane) => pane.id === 'topRight' || pane.id === 'main')?.screenRect.x;
  const topBoundary = plan.panes.find((pane) => pane.id === 'bottomLeft' || pane.id === 'main')?.screenRect.y;
  if (leftBoundary === undefined && topBoundary === undefined) return;
  context.save();
  context.strokeStyle = '#94a3b8';
  context.lineWidth = 1;
  context.beginPath();
  if (leftBoundary !== undefined) {
    context.moveTo(Math.round(leftBoundary) + 0.5, 0);
    context.lineTo(Math.round(leftBoundary) + 0.5, plan.viewport.height);
  }
  if (topBoundary !== undefined) {
    context.moveTo(0, Math.round(topBoundary) + 0.5);
    context.lineTo(plan.viewport.width, Math.round(topBoundary) + 0.5);
  }
  context.stroke();
  context.restore();
}

function drawResizePreview(options: ChromeDrawOptions): void {
  const { context, skeleton, chrome, plan } = options;
  const preview = chrome.resizePreview;
  if (!preview) return;
  const origin = defaultHeaderOffset();
  context.save();
  context.strokeStyle = "#2563eb";
  context.lineWidth = 1.5;
  context.beginPath();
  if (preview.axis === "column") {
    const left = skeleton.getColumnLeft(preview.index) + preview.sizePx;
    let frozenLeft = 0;
    const main = plan.panes.find((pane) => pane.id === "main") ?? plan.panes.at(-1);
    if (main) {
      const x = Math.round(left - main.contentOrigin.x + main.screenRect.x) + 0.5;
      context.moveTo(x, origin.y);
      context.lineTo(x, plan.viewport.height);
    }
    void frozenLeft;
  } else {
    const top = skeleton.getRowTop(preview.index) + preview.sizePx;
    const main = plan.panes.find((pane) => pane.id === "main") ?? plan.panes.at(-1);
    if (main) {
      const y = Math.round(top - main.contentOrigin.y + main.screenRect.y) + 0.5;
      context.moveTo(origin.x, y);
      context.lineTo(plan.viewport.width, y);
    }
  }
  context.stroke();
  if (preview.label) {
    context.font = '12px Inter, sans-serif';
    const width = context.measureText(preview.label).width + 12;
    const x = Math.max(origin.x + 4, Math.min(plan.viewport.width - width - 4, preview.axis === 'column' ? skeleton.getColumnLeft(preview.index) + preview.sizePx + 6 : origin.x + 6));
    const y = preview.axis === 'column' ? origin.y + 6 : Math.max(origin.y + 4, Math.min(plan.viewport.height - 26, skeleton.getRowTop(preview.index) + preview.sizePx + 6));
    context.fillStyle = '#0f172a';
    context.fillRect(x, y, width, 22);
    context.fillStyle = '#ffffff';
    context.textBaseline = 'middle';
    context.fillText(preview.label, x + 6, y + 11);
  }
  context.restore();
}

function drawTableOutlines(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome } = options;
  if (chrome.tableOutlines.length === 0) return;
  for (const outline of chrome.tableOutlines) {
    const contentRect = contentRangeRect(skeleton, outline);
    if (!contentRect) continue;
    for (const pane of plan.panes) {
      const visible = intersect(
        { x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height },
        contentRect,
      );
      if (!visible) continue;
      const transform = paneTransform(pane);
      const screen = {
        x: visible.x + transform.dx,
        y: visible.y + transform.dy,
        width: visible.width,
        height: visible.height,
      };
      context.save();
      context.strokeStyle = '#2F5597';
      context.lineWidth = 2;
      context.strokeRect(screen.x + 1, screen.y + 1, screen.width - 2, screen.height - 2);
      context.restore();
    }
  }
}

function drawOutlineButton(context: CanvasRenderingContext2D, buttonLeft: number, buttonTop: number, collapsed: boolean): void {
  context.save();
  context.fillStyle = '#ffffff';
  context.strokeStyle = '#64748b';
  context.lineWidth = 1;
  context.fillRect(buttonLeft, buttonTop, 10, 10);
  context.strokeRect(buttonLeft + 0.5, buttonTop + 0.5, 9, 9);
  context.strokeStyle = '#334155';
  context.beginPath();
  if (collapsed) {
    context.moveTo(buttonLeft + 2.5, buttonTop + 5);
    context.lineTo(buttonLeft + 7.5, buttonTop + 5);
    context.moveTo(buttonLeft + 5, buttonTop + 2.5);
    context.lineTo(buttonLeft + 5, buttonTop + 7.5);
  } else {
    context.moveTo(buttonLeft + 2.5, buttonTop + 5);
    context.lineTo(buttonLeft + 7.5, buttonTop + 5);
  }
  context.stroke();
  context.restore();
}

function drawOutlineControls(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome } = options;
  if (chrome.outlineControls.length === 0) return;
  const origin = defaultHeaderOffset();
  for (const control of chrome.outlineControls) {
    if (control.axis === 'row') {
      const rowTop = skeleton.getRowTop(control.index);
      const rowHeight = skeleton.getRowHeight(control.index);
      for (const pane of plan.panes) {
        const visibleTop = rowTop - pane.contentOrigin.y + pane.screenRect.y;
        if (visibleTop + rowHeight < origin.y || visibleTop > plan.viewport.height) continue;
        const buttonLeft = 4 + (control.level - 1) * 10;
        const buttonTop = visibleTop + Math.max(2, (rowHeight - 10) / 2);
        drawOutlineButton(context, buttonLeft, buttonTop, control.collapsed);
      }
      continue;
    }
    for (const pane of plan.panes) {
      if (pane.contentOrigin.y !== 0 || pane.id === 'bottomLeft') continue;
      const t = paneTransform(pane);
      const left = skeleton.getColumnLeft(control.index);
      const width = skeleton.getColumnWidth(control.index);
      if (left + width <= pane.contentOrigin.x || left >= pane.contentOrigin.x + pane.screenRect.width) continue;
      const buttonLeft = left + t.dx + 2;
      const buttonTop = 2 + (control.level - 1) * 10;
      drawOutlineButton(context, buttonLeft, buttonTop, control.collapsed);
    }
  }
}

function drawFilterFunnelIcon(context: CanvasRenderingContext2D, cx: number, cy: number): void {
  context.fillStyle = '#2563eb';
  context.beginPath();
  context.moveTo(cx, cy - 4);
  context.lineTo(cx + 8, cy - 4);
  context.lineTo(cx + 4.5, cy);
  context.lineTo(cx + 4.5, cy + 4);
  context.lineTo(cx + 3.5, cy + 4);
  context.lineTo(cx + 3.5, cy);
  context.closePath();
  context.fill();
}

function drawFilterFunnels(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome } = options;
  if (chrome.filterButtons.length > 0) {
    for (const pane of plan.panes) {
      const transform = paneTransform(pane);
      for (const button of chrome.filterButtons) {
        const rect = skeleton.getCellRect(button.row, button.column);
        if (!rect) continue;
        const visible = intersect(
          { x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height },
          rect,
        );
        if (!visible) continue;
        const cx = visible.x + transform.dx + visible.width - 12;
        const cy = visible.y + transform.dy + 8;
        drawFilterFunnelIcon(context, cx, cy);
      }
    }
    return;
  }

  if (chrome.filterColumns.length === 0) return;
  for (const pane of plan.panes) {
    if (pane.contentOrigin.y !== 0 || pane.id === "bottomLeft") continue;
    const t = paneTransform(pane);
    for (const column of chrome.filterColumns) {
      const left = skeleton.getColumnLeft(column);
      const width = skeleton.getColumnWidth(column);
      if (left + width <= pane.contentOrigin.x || left >= pane.contentOrigin.x + pane.screenRect.width) continue;
      const cx = left + width + t.dx - 9;
      const cy = COL_HEADER_HEIGHT / 2;
      drawFilterFunnelIcon(context, cx, cy);
    }
  }
}

function drawRemoteCursors(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome } = options;
  for (const cursor of chrome.remoteCursors) {
    const rect = skeleton.getCellRect(cursor.row, cursor.column);
    if (!rect) continue;
    for (const pane of plan.panes) {
      const visible = intersect({ x: pane.contentOrigin.x, y: pane.contentOrigin.y, width: pane.screenRect.width, height: pane.screenRect.height }, rect);
      if (!visible) continue;
      const t = paneTransform(pane);
      context.strokeStyle = cursor.color;
      context.lineWidth = 2;
      context.strokeRect(visible.x + t.dx, visible.y + t.dy, visible.width, visible.height);
      context.fillStyle = cursor.color;
      context.font = HEADER_FONT;
      context.textBaseline = "top";
      context.fillText(cursor.name, visible.x + t.dx + 2, visible.y + t.dy + 1);
    }
  }
}

function drawHeaderStrips(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome, theme } = options;
  const viewport = plan.viewport;
  const origin = defaultHeaderOffset();

  // 背景
  context.fillStyle = theme.headerBackground;
  context.fillRect(origin.x, origin.y, viewport.width - origin.x, COL_HEADER_HEIGHT);
  context.fillRect(0, origin.y, origin.x, viewport.height - origin.y);
  context.fillRect(0, 0, origin.x, origin.y);

  context.strokeStyle = theme.headerBorder;
  context.lineWidth = 1;
  context.strokeRect(0.5, 0.5, viewport.width - 1, viewport.height - 1);

  context.font = HEADER_FONT;
  context.fillStyle = theme.headerText;
  context.textBaseline = "middle";

  // 列头
  for (const pane of plan.panes) {
    if (pane.visibleRange == null) continue;
    const t = paneTransform(pane);
    for (let column = pane.visibleRange.startColumn; column <= pane.visibleRange.endColumn; column++) {
      const left = skeleton.getColumnLeft(column) + t.dx;
      const width = skeleton.getColumnWidth(column);
       const isSelected = isColumnSelected(chrome, column);
      if (isSelected) {
        context.fillStyle = "#dbeafe";
        context.fillRect(left, origin.y, width, COL_HEADER_HEIGHT);
        context.fillStyle = theme.headerText;
      }
      context.strokeStyle = theme.headerBorder;
      context.strokeRect(left + 0.5, origin.y + 0.5, width, COL_HEADER_HEIGHT - 1);
      context.textAlign = "center";
      context.fillText(columnLabelOf(column), left + width / 2, origin.y + COL_HEADER_HEIGHT / 2 + 1);
    }
    // 行头
    for (let row = pane.visibleRange.startRow; row <= pane.visibleRange.endRow; row++) {
      const top = skeleton.getRowTop(row) + t.dy;
      const height = skeleton.getRowHeight(row);
       const isSelected = isRowSelected(chrome, row);
      if (isSelected) {
        context.fillStyle = "#dbeafe";
        context.fillRect(0, top, origin.x, height);
        context.fillStyle = theme.headerText;
      }
      context.strokeStyle = theme.headerBorder;
      context.strokeRect(0.5, top + 0.5, origin.x - 1, height);
      context.textAlign = "right";
      context.fillText(String(row + 1), origin.x - 6, top + height / 2 + 1);
    }
  }

  // 全选角块
  context.fillStyle = chrome.selection.ranges.some((range) => isSelectAllRange(range, skeleton)) ? "#dbeafe" : theme.headerBackground;
  context.fillRect(0, 0, origin.x, origin.y);
  context.strokeStyle = theme.headerBorder;
  context.strokeRect(0.5, 0.5, origin.x - 1, origin.y - 1);
  context.textAlign = "left";
}

export function isColumnSelected(chrome: ChromeState, column: number): boolean {
  return chrome.selection.ranges.some((range) =>
    range.startColumn <= column && column <= range.endColumn);
}

export function isRowSelected(chrome: ChromeState, row: number): boolean {
  return chrome.selection.ranges.some((range) =>
    range.startRow <= row && row <= range.endRow);
}

export function isSelectAllRange(
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  skeleton: SheetSkeleton,
): boolean {
  return range.startRow === 0 && range.startColumn === 0
    && range.endRow >= skeleton.rowCount - 1 && range.endColumn >= skeleton.columnCount - 1;
}

void ROW_HEADER_WIDTH;
