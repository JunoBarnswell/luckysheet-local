import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Box,
  ContextMenu,
  type ContextMenuItem,
  Panel,
  Stack,
  StatePanel,
  Text,
  Button,
  Inline,
} from "@react-sheets/ui-system";
import {
  CanvasRenderSurface,
  CanvasRenderEngine,
  SheetSkeleton,
  type CellRenderData,
  type ChromeState,
  type FloatingDrawable,
  type FloatingHit,
  type HeaderHit,
  type Rect,
  createEmptyChromeState,
} from "@react-sheets/render-engine";
import { drawChartOnCanvas, drawShapeOnCanvas, drawSparklineOnCanvas } from "@react-sheets/pro-features";
import type { ChartModel, PivotResultTree, RangeRef, ShapeModel, SparklineModel } from "@react-sheets/core-model";
import { CellEditor } from "./CellEditor";
import { FilterPopover } from "./FilterPopover";
import type { PeerCursor, SelectionState, SheetCell, SheetView, WorkspacePhase } from "../state/workspace";

const CHART_PALETTE = ["#2563eb", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899", "#06b6d4"];

export interface SheetCanvasProps {
  sheet: SheetView;
  sheetId: string;
  selection: SelectionState;
  activeCell: string;
  formulaDraft: string;
  editingCell: { row: number; column: number } | null;
  phase: WorkspacePhase;
  zoom: number;
  peers: PeerCursor[];
  charts?: ChartModel[];
  pivotResults?: Record<string, PivotResultTree>;
  shapes?: ShapeModel[];
  sparklines?: SparklineModel[];
  selectedFloatingId: string | null;
  onSelectionChange: (selection: SelectionState) => void;
  onCommitCell: (value: string) => void;
  onBeginEdit: (initialText?: string) => void;
  onCancelEdit: () => void;
  onCommitEdit: (moveAfter?: "down" | "up" | "left" | "right" | "none") => void;
  onFormulaDraftChange: (value: string) => void;
  onInsertRef: (refText: string) => void;
  onToggleAbsolute: () => void;
  onJumpEdge: (direction: "up" | "down" | "left" | "right", extend?: boolean) => void;
  onSelectAll: () => void;
  onResizeRow: (row: number, heightPx: number) => void;
  onResizeColumn: (column: number, widthPx: number) => void;
  onFillRange: (target: { startRow: number; endRow: number; startColumn: number; endColumn: number }) => void;
  onFloatingSelect: (hit: FloatingHit | null) => void;
  onFloatingMove: (kind: "chart" | "shape" | "image", id: string, bounds: Rect) => void;
  onFloatingRemove: (kind: "chart" | "shape" | "image", id: string) => void;
  onAction: (action: string, payload?: unknown) => void;
  onApplyFilter: (column: number, patch: { selectedValues?: string[] | null }) => void;
  getValidationList: (row: number, column: number) => string[] | undefined;
  onRetry: () => void;
  onCreateSheet: () => void;
}

interface DragState {
  kind: "select" | "fill" | "col-resize" | "row-resize" | "floating-move" | "floating-resize";
  startRow: number;
  startColumn: number;
  anchorRow: number;
  anchorColumn: number;
  currentRow: number;
  currentColumn: number;
  additive: boolean;
  resizeStartSize: number;
  resizeIndex: number;
  floating?: { id: string; handle?: string; startBounds: Rect; startLocal: { x: number; y: number } };
}

function toChromeSelection(selection: SelectionState): ChromeState['selection'] {
  return {
    ranges: selection.ranges.map((range) => ({
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    })),
    primary: { row: selection.primaryRowIndex, column: selection.primaryColumnIndex },
    primaryIndex: selection.primaryRangeIndex,
  };
}

export function SheetCanvas({
  sheet,
  sheetId,
  selection,
  activeCell,
  formulaDraft,
  editingCell,
  phase,
  zoom,
  peers,
  charts = [],
  pivotResults = {},
  shapes = [],
  sparklines = [],
  selectedFloatingId,
  onSelectionChange,
  onCommitCell,
  onBeginEdit,
  onCancelEdit,
  onCommitEdit,
  onFormulaDraftChange,
  onInsertRef,
  onToggleAbsolute,
  onJumpEdge,
  onSelectAll,
  onResizeRow,
  onResizeColumn,
  onFillRange,
  onFloatingSelect,
  onFloatingMove,
  onFloatingRemove,
  onAction,
  onApplyFilter,
  getValidationList,
  onRetry,
  onCreateSheet,
}: SheetCanvasProps) {
  const engineRef = useRef<CanvasRenderEngine | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const dragRef = useRef<DragState | null>(null);
  const [contextMenu, setContextMenu] = useState({ x: 0, y: 0, open: false });
  const [filterPopover, setFilterPopover] = useState<{ column: number; x: number; y: number } | null>(null);
  const [validationDropdown, setValidationDropdown] = useState<{ row: number; column: number; options: string[] } | null>(null);
  const [fillPreview, setFillPreview] = useState<{ startRow: number; endRow: number; startColumn: number; endColumn: number } | null>(null);
  const [scrollTick, setScrollTick] = useState(0);
  const transientSelectionRef = useRef<SelectionState | null>(null);
  const transientSelectionFrameRef = useRef<number | null>(null);

  const zoomFactor = zoom / 100;

  const skeleton = useMemo(
    () =>
      new SheetSkeleton({
        rowCount: Math.max(sheet.rowCount, 200),
        columnCount: Math.max(sheet.columns.length, 26),
        defaultRowHeight: 28,
        defaultColumnWidth: 110,
        rowHeights: new Map(Object.entries(sheet.rowHeights).map(([key, value]) => [Number(key), value])),
        columnWidths: new Map(Object.entries(sheet.columnWidths).map(([key, value]) => [Number(key), value])),
        hiddenRows: new Set(sheet.hiddenRows),
        zoom: zoomFactor,
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet.rowCount, sheet.columns.length, sheet.rowHeights, sheet.columnWidths, sheet.hiddenRows, zoomFactor],
  );

  const cellProvider = useCallback(({ row, column }: { row: number; column: number }): CellRenderData | undefined => {
    const cell = sheet.getCell(row, column);
    if (!cell) return undefined;
    const merge = sheet.merges.find((span) =>
      row >= span.range.startRow && row <= span.range.endRow
      && column >= span.range.startColumn && column <= span.range.endColumn);
    const isAnchor = merge ? merge.anchor.row === row && merge.anchor.column === column : true;
    return {
      value: parseCellValue(cell),
      formula: cell.formula,
      displayValue: cell.value,
      style: cell.style,
      overlay: cell.overlay
        ? {
            dataBar: cell.overlay.dataBar,
            colorScale: cell.overlay.colorScale,
            icon: cell.overlay.icon,
          }
        : undefined,
      hasComment: cell.hasComment,
      invalid: cell.invalid,
      merge: merge
        ? {
            startRow: merge.range.startRow,
            endRow: merge.range.endRow,
            startColumn: merge.range.startColumn,
            endColumn: merge.range.endColumn,
            isAnchor,
          }
        : undefined,
    };
  }, [sheet]);

  // ---------- 浮动对象绘制器 ----------

  const floatables = useMemo<FloatingDrawable[]>(() => {
    const drawables: FloatingDrawable[] = [];
    for (const chart of charts) {
      const data = getChartSeries(chart);
      const series = data.series.map((entry, index) => ({
        ...entry,
        color: CHART_PALETTE[index % CHART_PALETTE.length]!,
      }));
      drawables.push({
        kind: "chart",
        id: chart.id,
        bounds: chart.bounds,
        draw: (context, rect) =>
          drawChartOnCanvas({ context, chart: { ...chart, bounds: rect }, categories: data.categories, series }),
      });
    }
    for (const shape of shapes) {
      drawables.push({
        kind: "shape",
        id: shape.id,
        bounds: shape.bounds,
        draw: (context, rect) => drawShapeOnCanvas({ context, shape: { ...shape, bounds: rect } }),
      });
    }
    for (const sparkline of sparklines) {
      const rect = skeleton.getCellRect(sparkline.anchor.row, sparkline.anchor.column);
      if (!rect) continue;
      drawables.push({
        kind: "shape",
        id: sparkline.id,
        bounds: rect,
        draw: (context, target) =>
          drawSparklineOnCanvas({ context, sparkline, values: getSparklineValues(sparkline), rect: target }),
      });
    }
    return drawables;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [charts, pivotResults, shapes, sparklines, skeleton, sheetId]);

  function getChartSeries(chart: ChartModel): { categories: string[]; series: Array<{ name: string; values: number[] }> } {
    const categories: string[] = [];
    const series: Array<{ name: string; values: number[] }> = [];
    const pivot = chart.pivotId ? pivotResults[chart.pivotId] : undefined;
    if (pivot) {
      const leaves: typeof pivot.rows = [];
      const collect = (nodes: typeof pivot.rows) => nodes.forEach((node) => node.children.length ? collect(node.children) : leaves.push(node));
      collect(pivot.rows);
      const valueCount = pivot.rows[0]?.values[0]?.values.length ?? 0;
      for (let index = 0; index < valueCount; index++) series.push({ name: chart.title ?? `Value ${index + 1}`, values: [] });
      for (const node of leaves) {
        categories.push(node.label);
        const cell = node.values[0];
        for (let index = 0; index < valueCount; index++) {
          const numeric = Number(cell?.values[index]);
          series[index]?.values.push(Number.isFinite(numeric) ? numeric : 0);
        }
      }
      return { categories, series };
    }
    const source = chart.sourceRanges[0];
    if (!source) return { categories, series };
    const values: number[] = [];
    for (let row = source.startRow; row <= source.endRow; row++) {
      for (let column = source.startColumn; column <= source.endColumn; column++) {
        const cell = sheet.getCell(row, column);
        if (!cell) continue;
        const numeric = Number(cell.value.replace(/[$,%]/g, ""));
        if (Number.isFinite(numeric) && cell.value !== "") {
          values.push(numeric);
        } else if (cell.value !== "") {
          categories.push(cell.value);
        }
      }
    }
    if (values.length > 0) series.push({ name: chart.title ?? "Series 1", values });
    return { categories, series };
  }

  function getSparklineValues(sparkline: SparklineModel): number[] {
    const values: number[] = [];
    const source = sparkline.sourceRange;
    for (let row = source.startRow; row <= source.endRow; row++) {
      for (let column = source.startColumn; column <= source.endColumn; column++) {
        const cell = sheet.getCell(row, column);
        if (!cell) continue;
        const numeric = Number(cell.value.replace(/[$,%]/g, ""));
        if (Number.isFinite(numeric) && cell.value !== "") values.push(numeric);
      }
    }
    return values;
  }

  // ---------- 引擎生命周期与 chrome 同步 ----------

  const chromeState = useMemo<ChromeState>(() => {
    const state = createEmptyChromeState();
    state.selection = toChromeSelection(selection);
    state.editing = editingCell ? { row: editingCell.row, column: editingCell.column } : null;
    state.filterColumns = sheet.filterColumns;
    state.remoteCursors = peers.map((peer) => ({
      actorId: peer.actorId,
      color: peer.color,
      name: peer.name,
      row: peer.row,
      column: peer.column,
    }));
    state.selectedFloatingId = selectedFloatingId;
    return state;
  }, [editingCell, peers, selectedFloatingId, selection, sheet.filterColumns]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setCellProvider(cellProvider);
  }, [cellProvider]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setSkeleton(skeleton);
  }, [skeleton]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setFreeze(
      sheet.freeze.xSplit > 0 || sheet.freeze.ySplit > 0
        ? { xSplit: sheet.freeze.xSplit, ySplit: sheet.freeze.ySplit }
        : null,
    );
  }, [sheet.freeze.xSplit, sheet.freeze.ySplit]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setFloating(floatables, selectedFloatingId);
  }, [floatables, selectedFloatingId]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setChrome(chromeState);
  }, [chromeState]);

  const clearTransientSelection = useCallback(() => {
    transientSelectionRef.current = null;
    if (transientSelectionFrameRef.current !== null) {
      if (typeof window !== 'undefined') window.cancelAnimationFrame(transientSelectionFrameRef.current);
      transientSelectionFrameRef.current = null;
    }
    const engine = engineRef.current;
    if (engine) engine.setChrome(chromeState);
  }, [chromeState]);

  const queueTransientSelection = useCallback((nextSelection: SelectionState) => {
    transientSelectionRef.current = nextSelection;
    if (transientSelectionFrameRef.current !== null) return;
    const draw = () => {
      transientSelectionFrameRef.current = null;
      const preview = transientSelectionRef.current;
      const engine = engineRef.current;
      if (!preview || !engine) return;
      engine.setChrome({ ...chromeState, selection: toChromeSelection(preview) });
    };
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
      transientSelectionFrameRef.current = window.requestAnimationFrame(draw);
    } else {
      transientSelectionFrameRef.current = setTimeout(draw, 0) as unknown as number;
    }
  }, [chromeState]);

  useEffect(() => () => {
    if (transientSelectionFrameRef.current === null) return;
    if (typeof window !== 'undefined') window.cancelAnimationFrame(transientSelectionFrameRef.current);
    transientSelectionFrameRef.current = null;
  }, []);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    const detach = engine.onViewportChanged(() => setScrollTick((tick) => tick + 1));
    return detach;
  }, []);

  // 选区变化 → 滚动至可见
  useEffect(() => {
    const engine = engineRef.current;
    if (!engine || dragRef.current) return;
    engine.ensureVisible({ row: selection.primaryRowIndex, column: selection.primaryColumnIndex });
  }, [selection.primaryRowIndex, selection.primaryColumnIndex]);

  // ---------- 指针交互 ----------

  const localPointOf = useCallback((event: { clientX: number; clientY: number }) => {
    const host = containerRef.current;
    if (!host) return { x: 0, y: 0 };
    const bounds = host.getBoundingClientRect();
    return { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
  }, []);

  const handlePointerDown = useCallback(
    (event: React.PointerEvent) => {
      if (phase !== "ready") return;
      if (event.button === 2) return; // 右键交给 contextmenu
      const engine = engineRef.current;
      const host = containerRef.current;
      if (!engine || !host) return;
      host.focus();
      setFilterPopover(null);
      setValidationDropdown(null);
      const local = localPointOf(event);

      // 1) 浮动对象优先
      const floatingHit = engine.hitTestFloating(local);
      if (floatingHit) {
        const drawableBounds = floatables.find((item) => item.id === floatingHit.id)?.bounds;
        if (drawableBounds) {
          dragRef.current = {
            kind: floatingHit.handle ? "floating-resize" : "floating-move",
            startRow: 0,
            startColumn: 0,
            anchorRow: 0,
            anchorColumn: 0,
            currentRow: 0,
            currentColumn: 0,
            additive: false,
            resizeStartSize: 0,
            resizeIndex: 0,
            floating: {
              id: floatingHit.id,
              handle: floatingHit.handle,
              startBounds: { ...drawableBounds },
              startLocal: local,
            },
          };
          onFloatingSelect(floatingHit);
          (event.target as Element).setPointerCapture?.(event.pointerId);
          return;
        }
      }
      onFloatingSelect(null);

      // 2) 表头区
      const headerHit = engine.headerHitAtLocal(local);
      if (headerHit) {
        if (headerHit.kind === "corner") {
          onSelectAll();
          return;
        }
        if (headerHit.resizeBoundaryPx !== undefined) {
          dragRef.current = {
            kind: headerHit.kind === "col" ? "col-resize" : "row-resize",
            startRow: 0,
            startColumn: 0,
            anchorRow: 0,
            anchorColumn: 0,
            currentRow: 0,
            currentColumn: 0,
            additive: false,
            resizeStartSize: headerHit.kind === "col" ? skeleton.getColumnWidth(headerHit.index) : skeleton.getRowHeight(headerHit.index),
            resizeIndex: headerHit.index,
          };
          (event.target as Element).setPointerCapture?.(event.pointerId);
          return;
        }
        const additive = event.ctrlKey || event.metaKey;
        if (headerHit.kind === "col") {
          if (sheet.filterColumns.includes(headerHit.index)) {
            setFilterPopover({ column: headerHit.index, x: event.clientX, y: event.clientY });
          }
        }
        dragRef.current = {
          kind: "select",
          startRow: headerHit.kind === "row" ? headerHit.index : 0,
          startColumn: headerHit.kind === "col" ? headerHit.index : 0,
          anchorRow: headerHit.kind === "row" ? headerHit.index : 0,
          anchorColumn: headerHit.kind === "col" ? headerHit.index : 0,
          currentRow: headerHit.kind === "row" ? headerHit.index : 0,
          currentColumn: headerHit.kind === "col" ? headerHit.index : 0,
          additive,
          resizeStartSize: 0,
          resizeIndex: 0,
          floating: { id: headerHit.kind, handle: undefined, startBounds: { x: 0, y: 0, width: 0, height: 0 }, startLocal: { x: 0, y: 0 } },
        };
        (event.target as Element).setPointerCapture?.(event.pointerId);
        return;
      }

      // 3) 填充柄
      const primaryRange = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
      if (primaryRange) {
        const rect = skeleton.getRangeRect({
          startRow: primaryRange.endRow,
          endRow: primaryRange.endRow,
          startColumn: primaryRange.endColumn,
          endColumn: primaryRange.endColumn,
        });
        if (rect) {
          const screen = engine.contentToMainScreen({ x: rect.x + rect.width, y: rect.y + rect.height });
          const half = 5;
          if (Math.abs(local.x - screen.x) <= half && Math.abs(local.y - screen.y) <= half) {
            dragRef.current = {
              kind: "fill",
              startRow: primaryRange.startRow,
              startColumn: primaryRange.startColumn,
              anchorRow: primaryRange.endRow,
              anchorColumn: primaryRange.endColumn,
              currentRow: primaryRange.endRow,
              currentColumn: primaryRange.endColumn,
              additive: false,
              resizeStartSize: 0,
              resizeIndex: 0,
            };
            (event.target as Element).setPointerCapture?.(event.pointerId);
            return;
          }
        }
      }

      // 4) 普通单元格选择/拖选
      const cell = engine.cellAtLocalPoint(local);
      if (!cell) return;
      const additive = event.ctrlKey || event.metaKey;
      dragRef.current = {
        kind: "select",
        startRow: cell.row,
        startColumn: cell.column,
        anchorRow: cell.row,
        anchorColumn: cell.column,
        currentRow: cell.row,
        currentColumn: cell.column,
        additive,
        resizeStartSize: 0,
        resizeIndex: 0,
        floating: undefined,
      };
      if (!additive) {
        onSelectionChange({
          ranges: [{ sheetId, startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column }],
          primaryRowIndex: cell.row,
          primaryColumnIndex: cell.column,
          primaryRangeIndex: 0,
        });
      }
      (event.target as Element).setPointerCapture?.(event.pointerId);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [floatables, localPointOf, onFloatingSelect, onSelectAll, onSelectionChange, phase, selection, sheet.filterColumns, sheetId, skeleton],
  );

  const handlePointerMove = useCallback(
    (event: React.PointerEvent) => {
      const engine = engineRef.current;
      if (!engine) return;
      const local = localPointOf(event);
      const drag = dragRef.current;

      if (!drag) {
        // 悬停光标提示
        const headerHit = engine.headerHitAtLocal(local);
        const host = containerRef.current;
        if (host) {
          host.style.cursor = headerHit?.resizeBoundaryPx !== undefined
            ? (headerHit.kind === "col" ? "col-resize" : "row-resize")
            : "default";
        }
        return;
      }

      if (drag.kind === "col-resize" || drag.kind === "row-resize") {
        const content = engine.localToContent(local);
        const boundary = drag.kind === "col-resize"
          ? skeleton.getColumnLeft(drag.resizeIndex)
          : skeleton.getRowTop(drag.resizeIndex);
        const size = Math.max(24, (drag.kind === "col-resize" ? content.x : content.y) - boundary);
        const nextChrome = createEmptyChromeState();
        nextChrome.resizePreview = { axis: drag.kind === "col-resize" ? "column" : "row", index: drag.resizeIndex, sizePx: size };
        engine.setChrome({ ...chromeState, resizePreview: nextChrome.resizePreview });
        return;
      }

      if (drag.kind === "floating-move" && drag.floating) {
        const deltaX = local.x - drag.floating.startLocal.x;
        const deltaY = local.y - drag.floating.startLocal.y;
        const content = engine.localToContent(local);
        void content;
        onFloatingMove(
          drag.kind === "floating-move" ? "shape" : "shape",
          drag.floating.id,
          {
            x: drag.floating.startBounds.x + deltaX,
            y: drag.floating.startBounds.y + deltaY,
            width: drag.floating.startBounds.width,
            height: drag.floating.startBounds.height,
          },
        );
        return;
      }

      if (drag.kind === "floating-resize" && drag.floating?.handle) {
        const handle = drag.floating.handle;
        const start = drag.floating.startBounds;
        const deltaX = local.x - drag.floating.startLocal.x;
        const deltaY = local.y - drag.floating.startLocal.y;
        let x = start.x;
        let y = start.y;
        let width = start.width;
        let height = start.height;
        if (handle.includes("e")) width = Math.max(40, start.width + deltaX);
        if (handle.includes("s")) height = Math.max(30, start.height + deltaY);
        if (handle.includes("w")) { width = Math.max(40, start.width - deltaX); x = start.x + (start.width - width); }
        if (handle.includes("n")) { height = Math.max(30, start.height - deltaY); y = start.y + (start.height - height); }
        onFloatingMove("shape", drag.floating.id, { x, y, width, height });
        return;
      }

      // select / fill:更新当前行列
      const cell = engine.cellAtLocalPoint(local);
      if (!cell) return;
      if (drag.kind === "fill") {
        const vertical = Math.abs(cell.row - drag.anchorRow) >= Math.abs(cell.column - drag.anchorColumn);
        drag.currentRow = vertical ? cell.row : drag.anchorRow;
        drag.currentColumn = vertical ? drag.anchorColumn : cell.column;
        setFillPreview({
          startRow: Math.min(drag.startRow, drag.currentRow),
          endRow: Math.max(drag.anchorRow, drag.currentRow),
          startColumn: Math.min(drag.startColumn, drag.currentColumn),
          endColumn: Math.max(drag.anchorColumn, drag.currentColumn),
        });
        return;
      }
      drag.currentRow = cell.row;
      drag.currentColumn = cell.column;
      const startRow = Math.min(drag.anchorRow, cell.row);
      const endRow = Math.max(drag.anchorRow, cell.row);
      const startColumn = Math.min(drag.anchorColumn, cell.column);
      const endColumn = Math.max(drag.anchorColumn, cell.column);
      const isRowDrag = drag.floating?.id === "row";
      const isColDrag = drag.floating?.id === "col";
      const nextSelection: SelectionState = {
        ranges: [{
          sheetId,
          startRow: isRowDrag || isColDrag ? drag.startRow : startRow,
          endRow: isRowDrag ? drag.currentRow : isColDrag ? Math.max(0, skeleton.rowCount - 1) : endRow,
          startColumn: isColDrag ? drag.startColumn : startColumn,
          endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : endColumn,
        }],
        primaryRowIndex: isRowDrag || isColDrag ? drag.anchorRow : drag.anchorRow,
        primaryColumnIndex: drag.anchorColumn,
        primaryRangeIndex: 0,
      };
      const previewSelection = drag.additive
        ? { ...selection, ranges: [...selection.ranges, ...nextSelection.ranges], primaryRangeIndex: selection.ranges.length, primaryRowIndex: nextSelection.primaryRowIndex, primaryColumnIndex: nextSelection.primaryColumnIndex }
        : nextSelection;
      queueTransientSelection(previewSelection);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [localPointOf, onFloatingMove, queueTransientSelection, selection, sheetId, skeleton],
  );

  const handlePointerUp = useCallback(
    (event: React.PointerEvent) => {
      const drag = dragRef.current;
      dragRef.current = null;
      const engine = engineRef.current;
      if (!drag || !engine) return;
      (event.target as Element).releasePointerCapture?.(event.pointerId);

      if (drag.kind === "col-resize") {
        const content = engine.localToContent(localPointOf(event));
        const width = Math.max(24, content.x - skeleton.getColumnLeft(drag.resizeIndex));
        onResizeColumn(drag.resizeIndex, Math.round(width / (zoom / 100)));
        return;
      }
      if (drag.kind === "row-resize") {
        const content = engine.localToContent(localPointOf(event));
        const height = Math.max(18, content.y - skeleton.getRowTop(drag.resizeIndex));
        onResizeRow(drag.resizeIndex, Math.round(height / (zoom / 100)));
        return;
      }
      if (drag.kind === "fill") {
        setFillPreview(null);
        const target = {
          startRow: Math.min(drag.startRow, drag.currentRow),
          endRow: Math.max(drag.anchorRow, drag.currentRow),
          startColumn: Math.min(drag.startColumn, drag.currentColumn),
          endColumn: Math.max(drag.anchorColumn, drag.currentColumn),
        };
        if (target.endRow !== drag.anchorRow || target.endColumn !== drag.anchorColumn) {
          onFillRange(target);
        }
        return;
      }
      if (drag.kind === "select") {
        const startRow = Math.min(drag.anchorRow, drag.currentRow);
        const endRow = Math.max(drag.anchorRow, drag.currentRow);
        const startColumn = Math.min(drag.anchorColumn, drag.currentColumn);
        const endColumn = Math.max(drag.anchorColumn, drag.currentColumn);
        const isRowDrag = drag.floating?.id === "row";
        const isColDrag = drag.floating?.id === "col";
        const range: RangeRef = {
          sheetId,
          startRow: isRowDrag || isColDrag ? drag.startRow : startRow,
          endRow: isRowDrag ? drag.currentRow : isColDrag ? Math.max(0, skeleton.rowCount - 1) : endRow,
          startColumn: isColDrag ? drag.startColumn : startColumn,
          endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : endColumn,
        };
        const nextSelection: SelectionState = drag.additive
          ? { ...selection, ranges: [...selection.ranges, range], primaryRangeIndex: selection.ranges.length, primaryRowIndex: drag.anchorRow, primaryColumnIndex: drag.anchorColumn }
          : { ranges: [range], primaryRowIndex: drag.anchorRow, primaryColumnIndex: drag.anchorColumn, primaryRangeIndex: 0 };
        clearTransientSelection();
        onSelectionChange(nextSelection);
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [clearTransientSelection, localPointOf, onFillRange, onResizeColumn, onResizeRow, onSelectionChange, selection, sheetId, skeleton, zoom],
  );

  const handleDoubleClick = useCallback(
    (event: React.PointerEvent | React.MouseEvent) => {
      const engine = engineRef.current;
      if (!engine) return;
      const local = localPointOf(event);
      const headerHit = engine.headerHitAtLocal(local);
      if (headerHit?.resizeBoundaryPx !== undefined) {
        // 双击边界 = 自适应内容宽(近似:取列内最长显示文本)
        const column = headerHit.index;
        let maxWidth = 60;
        const context = engine.getCanvas("content")?.getContext("2d");
        if (context) {
          context.font = "13px Inter, sans-serif";
          for (let row = 0; row < Math.min(sheet.rowCount, 200); row += 1) {
            const cell = sheet.getCell(row, column);
            if (cell?.value) maxWidth = Math.max(maxWidth, context.measureText(cell.value).width + 16);
          }
        }
        onResizeColumn(column, Math.round(maxWidth / (zoom / 100)));
        return;
      }
      const cell = engine.cellAtLocalPoint(local);
      if (!cell) return;
      const validationList = getValidationList(cell.row, cell.column);
      if (validationList && validationList.length > 0) {
        setValidationDropdown({ row: cell.row, column: cell.column, options: validationList });
        return;
      }
      onBeginEdit();
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [getValidationList, localPointOf, onBeginEdit, onResizeColumn, sheet, zoom],
  );

  const handleWheel = useCallback(
    (event: React.WheelEvent) => {
      const engine = engineRef.current;
      if (!engine) return;
      event.preventDefault();
      engine.scrollBy(event.deltaX, event.deltaY);
    },
    [],
  );

  // ---------- 键盘 ----------

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent) => {
      if (phase !== "ready") return;
      const key = event.key;
      const ctrl = event.ctrlKey || event.metaKey;

      if (editingCell) {
        // 编辑态按键由 CellEditor 处理;此处仅拦截 Escape 兜底
        if (key === "Escape") {
          event.preventDefault();
          onCancelEdit();
        }
        return;
      }

      if (ctrl && (key === "z" || key === "Z")) { event.preventDefault(); onAction("undo"); return; }
      if (ctrl && (key === "y" || key === "Y")) { event.preventDefault(); onAction("redo"); return; }
      if (ctrl && (key === "c" || key === "C")) { event.preventDefault(); onAction("copy"); return; }
      if (ctrl && (key === "x" || key === "X")) { event.preventDefault(); onAction("cut"); return; }
      if (ctrl && (key === "v" || key === "V")) { event.preventDefault(); onAction("paste"); return; }
      if (ctrl && (key === "b" || key === "B")) { event.preventDefault(); onAction("bold"); return; }
      if (ctrl && (key === "i" || key === "I")) { event.preventDefault(); onAction("italic"); return; }
      if (ctrl && (key === "u" || key === "U")) { event.preventDefault(); onAction("underline"); return; }
      if (key === "F2") { event.preventDefault(); onBeginEdit(); return; }
      if (key === "F4") { event.preventDefault(); onToggleAbsolute(); return; }
      if (key === "Delete" || key === "Backspace") { event.preventDefault(); onAction("clear-range"); return; }
      if (key === "Enter") { event.preventDefault(); onBeginEdit(); return; }

      const moves: Record<string, [number, number]> = {
        ArrowUp: [-1, 0],
        ArrowDown: [1, 0],
        ArrowLeft: [0, -1],
        ArrowRight: [0, 1],
        Tab: [0, event.shiftKey ? -1 : 1],
      };
      if (key in moves && !ctrl) {
        event.preventDefault();
        const [dr, dc] = moves[key]!;
        onSelectionChange({
          ...selection,
          primaryRowIndex: selection.primaryRowIndex + dr,
          primaryColumnIndex: selection.primaryColumnIndex + dc,
        });
        return;
      }
      if (key in moves && ctrl) {
        event.preventDefault();
        const direction = key === "ArrowUp" ? "up" : key === "ArrowDown" ? "down" : key === "ArrowLeft" ? "left" : "right";
        onJumpEdge(direction, event.shiftKey);
        return;
      }
      if (key === "Home") {
        event.preventDefault();
        onSelectionChange({ ...selection, primaryColumnIndex: 0, primaryRangeIndex: 0 });
        return;
      }
      if (key === "PageDown" || key === "PageUp") {
        event.preventDefault();
        const rows = Math.max(1, Math.floor((containerRef.current?.clientHeight ?? 600) / (28 * zoomFactor)) - 2);
        const delta = key === "PageDown" ? rows : -rows;
        onSelectionChange({ ...selection, primaryRowIndex: selection.primaryRowIndex + delta, primaryRangeIndex: 0 });
        return;
      }
      // 直接输入进入编辑
      if (key.length === 1 && !ctrl && !event.altKey) {
        event.preventDefault();
        onBeginEdit(key);
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editingCell, onCancelEdit, onBeginEdit, onInsertRef, onJumpEdge, onAction, onSelectionChange, phase, selection, zoomFactor],
  );

  // ---------- 右键菜单 ----------

  const contextMenuItems = useMemo<ContextMenuItem[]>(() => {
    const items: ContextMenuItem[] = [
      { id: "cut", label: "Cut", shortcut: "Ctrl+X", onSelect: () => onAction("cut") },
      { id: "copy", label: "Copy", shortcut: "Ctrl+C", onSelect: () => onAction("copy") },
      { id: "paste", label: "Paste", shortcut: "Ctrl+V", onSelect: () => onAction("paste") },
      { id: "sep-1", label: "", separator: true },
      { id: "insert-row", label: "Insert row above", onSelect: () => onAction("insert-row") },
      { id: "insert-column", label: "Insert column left", onSelect: () => onAction("insert-column") },
      { id: "delete-row", label: "Delete row", danger: true, onSelect: () => onAction("delete-row") },
      { id: "delete-column", label: "Delete column", danger: true, onSelect: () => onAction("delete-column") },
      { id: "sep-2", label: "", separator: true },
      { id: "hide-row", label: "Hide rows", onSelect: () => onAction("hide-row") },
      { id: "hide-col", label: "Hide columns", onSelect: () => onAction("hide-column") },
      { id: "unhide-all", label: "Unhide all", onSelect: () => onAction("unhide-all") },
      { id: "sep-3", label: "", separator: true },
      { id: "clear", label: "Clear contents", onSelect: () => onAction("clear-range") },
      { id: "clear-formats", label: "Clear formats", onSelect: () => onAction("clear-formats") },
      { id: "comment-add", label: "Add comment", onSelect: () => onAction("open-comments") },
    ];
    return items;
  }, [onAction]);

  // ---------- 编辑器定位(随滚动更新) ----------

  // 编辑器随滚动重定位:依赖 scrollTick 触发重算
  const editorRect = useMemo(() => {
    void scrollTick;
    const engine = engineRef.current;
    if (!engine || !editingCell) return null;
    const rect = skeleton.getCellRect(editingCell.row, editingCell.column);
    if (!rect) return null;
    const topLeft = engine.contentToMainScreen({ x: rect.x, y: rect.y });
    return {
      x: topLeft.x,
      y: topLeft.y,
      width: rect.width,
      height: rect.height,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editingCell, skeleton, scrollTick]);

  if (phase === "empty") {
    return (
      <Panel className="m-4 flex-1">
        <StatePanel
          kind="empty"
          description="Create a workbook to start editing cells."
          actionLabel="Create workbook"
          onAction={onCreateSheet}
          title="No workbook loaded"
        />
      </Panel>
    );
  }

  if (phase === "error") {
    return (
      <Panel className="m-4 flex-1">
        <StatePanel
          kind="error"
          description="The workbook engine failed to initialize. Retry to recover."
          actionLabel="Retry"
          onAction={onRetry}
          title="Engine error"
        />
      </Panel>
    );
  }

  return (
    <Panel className="h-full min-h-0 flex-1 overflow-hidden">
      <Stack gap="none" className="h-full">
        <Inline gap="xs" className="items-center justify-between border-b border-slate-100 px-3 py-1.5">
          <Inline gap="xs" className="items-center">
            <Text size="xs" tone="muted">Sheet</Text>
            <Text size="xs" weight="semibold">{sheet.name}</Text>
            {sheet.freeze.xSplit > 0 || sheet.freeze.ySplit > 0 ? (
              <Text size="xs" tone="subtle">frozen {sheet.freeze.xSplit}x{sheet.freeze.ySplit}</Text>
            ) : null}
          </Inline>
          <Text size="xs" tone="subtle">{activeCell}</Text>
        </Inline>
        <Box className="relative min-h-0 flex-1">
          <Box
            ref={containerRef}
            role="grid"
            aria-label="Spreadsheet canvas"
            tabIndex={0}
            className="absolute inset-0 outline-none"
            onPointerDown={handlePointerDown}
            onPointerMove={handlePointerMove}
            onPointerUp={handlePointerUp}
            onDoubleClick={handleDoubleClick}
            onWheel={handleWheel}
            onKeyDown={handleKeyDown}
            onContextMenu={(event) => {
              event.preventDefault();
              setContextMenu({ x: event.clientX, y: event.clientY, open: true });
            }}
          >
            <CanvasRenderSurface
              onReady={(engine) => {
                engineRef.current = engine;
                engine.setCellProvider(cellProvider);
                engine.setSkeleton(skeleton);
                engine.setFloating(floatables, selectedFloatingId);
                engine.setChrome(chromeState);
              }}
              className="absolute inset-0"
            />
          </Box>

          {editorRect && editingCell ? (
            <Box
              className="absolute z-20 border-2 border-blue-600 bg-white shadow-lg"
              style={{ left: editorRect.x - 1, top: editorRect.y - 1, minWidth: Math.max(editorRect.width + 2, 120) }}
            >
              <CellEditor
                initialText={formulaDraft}
                onCancel={onCancelEdit}
                onChange={onFormulaDraftChange}
                onCommit={onCommitEdit}
                onInsertRef={onInsertRef}
              />
            </Box>
          ) : null}

          {fillPreview ? (
            <FillPreviewOverlay skeleton={skeleton} engine={engineRef.current} preview={fillPreview} />
          ) : null}

          {validationDropdown ? (
            <ValidationDropdown
              options={validationDropdown.options}
              onPick={(value) => {
                onCommitCell(value);
                setValidationDropdown(null);
              }}
              onClose={() => setValidationDropdown(null)}
            />
          ) : null}

          {filterPopover ? (
            <FilterPopover
              column={filterPopover.column}
              sheet={sheet}
              onApply={(patch) => {
                onApplyFilter(filterPopover.column, patch);
                setFilterPopover(null);
              }}
              onClose={() => setFilterPopover(null)}
            />
          ) : null}
        </Box>
      </Stack>

      <ContextMenu
        x={contextMenu.x}
        y={contextMenu.y}
        open={contextMenu.open}
        items={contextMenuItems}
        onClose={() => setContextMenu((previous) => ({ ...previous, open: false }))}
      />
    </Panel>
  );
}

function parseCellValue(cell: SheetCell): string | number | boolean | null {
  const numeric = Number(cell.value.replace(/[$,]/g, ""));
  if (cell.value !== "" && Number.isFinite(numeric) && /\d/.test(cell.value)) return numeric;
  if (cell.value === "TRUE") return true;
  if (cell.value === "FALSE") return false;
  return cell.value;
}

function FillPreviewOverlay({
  skeleton,
  engine,
  preview,
}: {
  skeleton: SheetSkeleton;
  engine: CanvasRenderEngine | null;
  preview: { startRow: number; endRow: number; startColumn: number; endColumn: number };
}): React.ReactElement | null {
  if (!engine) return null;
  const rect = skeleton.getRangeRect(preview);
  if (!rect) return null;
  const screen = engine.contentToMainScreen(rect);
  return (
    <Box
      className="pointer-events-none absolute z-10 border-2 border-dashed border-blue-500 bg-blue-500/5"
      style={{ left: screen.x, top: screen.y, width: rect.width, height: rect.height }}
    />
  );
}

function ValidationDropdown({
  options,
  onPick,
  onClose,
}: {
  options: string[];
  onPick: (value: string) => void;
  onClose: () => void;
}): React.ReactElement {
  return (
    <Box className="absolute left-1/2 top-1/2 z-30 -translate-x-1/2 -translate-y-1/2 rounded-lg border border-slate-200 bg-white p-1 shadow-xl">
      <Stack gap="none">
        {options.map((option) => (
          <Button
            key={option}
            size="sm"
            variant="ghost"
            className="justify-start"
            onClick={() => onPick(option)}
          >
            {option}
          </Button>
        ))}
        <Button size="sm" variant="ghost" className="justify-start text-slate-400" onClick={onClose}>
          Cancel
        </Button>
      </Stack>
    </Box>
  );
}
