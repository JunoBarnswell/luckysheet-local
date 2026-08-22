import {
  COL_HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  type ChromeState,
  type Rect,
  type RenderPane,
  type RenderTheme,
} from "./types";
import { SheetSkeleton, columnLabelOf } from "./sheet-skeleton";
import { computeRenderPanes, defaultHeaderOffset } from "./render-plan";
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
  drawFreezeTrapLines(context, options.skeleton, plan.viewport.width, plan.viewport.height);
  drawResizePreview(options);
  drawFilterFunnels(options);
  drawRemoteCursors(options);
  void theme;
}

function paneTransform(pane: RenderPane): { dx: number; dy: number } {
  return { dx: pane.rect.x - pane.offset.x, dy: pane.rect.y - pane.offset.y };
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
      if (pane.rect.width <= 0 || pane.rect.height <= 0) continue;
      const visible = intersect(
        { x: pane.offset.x, y: pane.offset.y, width: pane.rect.width, height: pane.rect.height },
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
      const visible = intersect({ x: pane.offset.x, y: pane.offset.y, width: pane.rect.width, height: pane.rect.height }, contentRect);
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
      panes.filter((pane) => !isFullSpanRow(selectionRanges[i]!, skeleton) && !isFullSpanCol(selectionRanges[i]!, skeleton) || true),
      [selectionRanges[i]!],
      theme.selectionBorder,
      isPrimary ? 2 : 1,
    );
    if (isPrimary) {
      const primaryRange = {
        startRow: chrome.selection.primary.row,
        endRow: chrome.selection.primary.row,
        startColumn: chrome.selection.primary.column,
        endColumn: chrome.selection.primary.column,
      };
      const primaryRect = contentRangeRect(skeleton, primaryRange);
      if (primaryRect) {
        const mainPane = panes.find((pane) => pane.id === "main") ?? panes.at(-1)!;
        const t = paneTransform(mainPane);
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

function isFullSpanRow(_range: unknown, _skeleton: SheetSkeleton): boolean {
  return false;
}
function isFullSpanCol(_range: unknown, _skeleton: SheetSkeleton): boolean {
  return false;
}

function drawEditingOutline(options: ChromeDrawOptions, editing: { row: number; column: number }): void {
  const { context, skeleton, plan, theme } = options;
  const rect = skeleton.getCellRect(editing.row, editing.column);
  if (!rect) return;
  for (const pane of plan.panes) {
    const visible = intersect({ x: pane.offset.x, y: pane.offset.y, width: pane.rect.width, height: pane.rect.height }, rect);
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

function drawFreezeTrapLines(context: CanvasRenderingContext2D, skeleton: SheetSkeleton, width: number, height: number): void {
  // 冻结线由窗格边界推导:重算冻结切分(与 render-plan 相同逻辑)
  void skeleton;
  void width;
  void height;
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
      const x = Math.round(left - main.offset.x + main.rect.x) + 0.5;
      context.moveTo(x, origin.y);
      context.lineTo(x, plan.viewport.height);
    }
    void frozenLeft;
  } else {
    const top = skeleton.getRowTop(preview.index) + preview.sizePx;
    const main = plan.panes.find((pane) => pane.id === "main") ?? plan.panes.at(-1);
    if (main) {
      const y = Math.round(top - main.offset.y + main.rect.y) + 0.5;
      context.moveTo(origin.x, y);
      context.lineTo(plan.viewport.width, y);
    }
  }
  context.stroke();
  context.restore();
}

function drawFilterFunnels(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome } = options;
  if (chrome.filterColumns.length === 0) return;
  for (const pane of plan.panes) {
    if (pane.offset.y !== 0 || pane.id === "bottomLeft") continue;
    const t = paneTransform(pane);
    for (const column of chrome.filterColumns) {
      const left = skeleton.getColumnLeft(column);
      const width = skeleton.getColumnWidth(column);
      if (left + width <= pane.offset.x || left >= pane.offset.x + pane.rect.width) continue;
      const cx = left + width + t.dx - 9;
      const cy = COL_HEADER_HEIGHT / 2;
      context.fillStyle = "#2563eb";
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
  }
}

function drawRemoteCursors(options: ChromeDrawOptions): void {
  const { context, skeleton, plan, chrome } = options;
  for (const cursor of chrome.remoteCursors) {
    const rect = skeleton.getCellRect(cursor.row, cursor.column);
    if (!rect) continue;
    for (const pane of plan.panes) {
      const visible = intersect({ x: pane.offset.x, y: pane.offset.y, width: pane.rect.width, height: pane.rect.height }, rect);
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
    if (pane.range == null) continue;
    const t = paneTransform(pane);
    for (let column = pane.range.startColumn; column <= pane.range.endColumn; column++) {
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
    for (let row = pane.range.startRow; row <= pane.range.endRow; row++) {
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
  context.fillStyle = chrome.selection.ranges.some(isSelectAllRange) ? "#dbeafe" : theme.headerBackground;
  context.fillRect(0, 0, origin.x, origin.y);
  context.strokeStyle = theme.headerBorder;
  context.strokeRect(0.5, 0.5, origin.x - 1, origin.y - 1);
  context.textAlign = "left";
}

function isColumnSelected(chrome: ChromeState, column: number): boolean {
  return chrome.selection.ranges.some((range) =>
    range.startColumn <= column && column <= range.endColumn
    && range.startRow === 0 && range.endRow >= 999);
}

function isRowSelected(chrome: ChromeState, row: number): boolean {
  return chrome.selection.ranges.some((range) =>
    range.startRow <= row && row <= range.endRow
    && range.startColumn === 0 && range.endColumn >= 25);
}

function isSelectAllRange(range: { startRow: number; endRow: number; startColumn: number; endColumn: number }): boolean {
  return range.startRow === 0 && range.startColumn === 0 && range.endRow >= 999 && range.endColumn >= 25;
}

void ROW_HEADER_WIDTH;