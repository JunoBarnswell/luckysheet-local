import React, { useCallback, useEffect, useRef } from "react";
import {
  CanvasRenderEngine,
  type ChromeState,
  type FloatingDrawable,
  type FloatingHandle,
  type FloatingHit,
  type Rect,
  SheetSkeleton,
} from "@react-sheets/render-engine";
import type {
  DrawingObject,
  DrawingPayload,
  PivotProjectionCell,
  PivotSourceRowPath,
  RangeRef,
} from "@react-sheets/core-model";
import {
  createSpreadsheetShortcutRegistry,
  expandSelectionRangeForMerges,
  resolveContextHit,
  resolveSelectionTarget,
  selectionFromGesture,
  intersectsRange,
  containsRange,
  type AppPhase,
  type CanvasSheetSnapshot,
  type ResolvedContextHit,
  type SelectionGesture,
  type SelectionState,
} from "@react-sheets/spreadsheet-app";
import {
  claimPointerGesture,
  ownsPointerGesture,
  releasePointerGesture,
  releasePointerGesturesForSurface,
  resolvePointerGestureOwner,
} from "@react-sheets/ui-system";
import type { PivotControlAction } from "./drawing-renderers";
import {
  beginDimensionResizeGesture,
  updateDimensionResizeGesture,
  type DimensionResizeGesture,
} from "./dimension-resize-gesture";
import { resolveAutoScrollExtentGrowth } from './sheet-extent-growth';

export interface CanvasFillPreview {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export interface CanvasFilterPopoverState {
  column: number;
  x: number;
  y: number;
}

export interface CanvasValidationDropdownState {
  row: number;
  column: number;
  options: string[];
}

export interface CanvasContextMenuState {
  x: number;
  y: number;
  open: boolean;
}

interface DragState {
  kind: "select" | "fill" | "col-resize" | "row-resize" | "floating-move" | "floating-resize" | "textbox-placement";
  pointerId: number;
  startRow: number;
  startColumn: number;
  anchorRow: number;
  anchorColumn: number;
  currentRow: number;
  currentColumn: number;
  additive: boolean;
  extend: boolean;
  dimensionResize?: DimensionResizeGesture;
  floating?: {
    id: string;
    kind: "chart" | "shape" | "image" | "pivot-control";
    handle?: FloatingHandle;
    rotation?: number;
    startBounds: Rect;
    startLocal: { x: number; y: number };
  };
  textBox?: { startContent: { x: number; y: number } };
}

export interface CanvasInteractionOptions {
  engineRef: React.MutableRefObject<CanvasRenderEngine | null>;
  containerRef: React.MutableRefObject<HTMLDivElement | null>;
  contextRangeRef: React.MutableRefObject<RangeRef | null>;
  sheet: CanvasSheetSnapshot;
  sheetId: string;
  selection: SelectionState;
  editingCell: { row: number; column: number } | null;
  formulaDraft: string;
  phase: AppPhase;
  zoom: number;
  skeleton: SheetSkeleton;
  chromeState: ChromeState;
  floatables: readonly FloatingDrawable[];
  drawings: readonly DrawingObject[];
  selectedFloatingId: string | null;
  drawingSelectionMode: boolean;
  textBoxPlacementActive: boolean;
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  formatPainterActive: boolean;
  canRepeat: boolean;
  onPivotContextHit?: (hit: ResolvedContextHit | null) => void;
  /** Control-child actions are consumed before generic floating move/resize. */
  onPivotControlAction?: (drawingId: string, action: PivotControlAction) => void;
  onPivotShowDetails: (request: {
    pivotId: string;
    sourceRowPaths: readonly PivotSourceRowPath[];
    hit: ResolvedContextHit;
  }) => void;
  onPivotExpansionToggle: (pivotId: string, nodeId: string) => void;
  onSelectionChange: (selection: SelectionState) => void;
  onMovePrimary: (rowDelta: number, columnDelta: number, opts?: { extend?: boolean }) => void;
  onRequestExtentGrowth: (axes: { rows?: boolean; columns?: boolean }) => void;
  onBeginEdit: (initialText?: string) => void;
  onCancelEdit: () => void;
  onCommitEdit: (moveAfter?: "down" | "up" | "left" | "right" | "none") => void;
  onFormulaDraftChange: (value: string) => void;
  editComposing?: boolean;
  onAppendFormulaDraft?: (fragment: string) => void;
  onToggleAbsolute: () => void;
  onJumpEdge: (direction: "up" | "down" | "left" | "right", extend?: boolean) => void;
  onSelectAll: () => void;
  onExtendSelection?: (row: number, column: number) => void;
  onResizeRow: (row: number, heightPx: number) => void;
  onResizeColumn: (column: number, widthPx: number) => void;
  onAutoFitColumn: (column: number) => void | Promise<void>;
  onAutoFitRow: (row: number) => void | Promise<void>;
  formatColumnWidthPreview: (widthPx: number) => { widthPx: number; excelWidth: number };
  onFillRange: (target: CanvasFillPreview) => void;
  onExitDrawingSelectionMode?: () => void;
  onFloatingSelect: (hit: FloatingHit | null, mode?: "replace" | "add" | "toggle") => void;
  onToggleCheckbox: (ranges: RangeRef[]) => void;
  onFloatingMove: (drawingId: string, bounds: Rect, rotation?: number) => void;
  onTextBoxPlacementCommit: (bounds: Rect) => void;
  onCancelTextBoxPlacement: () => void;
  onBeginTextBoxEdit: (drawingId: string, initialText?: string) => void;
  onToggleOutline?: (groupId: string) => void;
  onShortcut?: (id: string) => boolean;
  onCancelFormatPainter?: () => void;
  onPivotResolve: (sheet: CanvasSheetSnapshot, row: number, column: number) => ResolvedContextHit | null;
  findPivotProjectionCell: (sheet: CanvasSheetSnapshot, row: number, column: number) => { projection: { pivotId: string }; cell: PivotProjectionCell } | null;
  isPivotValueCell: (cell: PivotProjectionCell) => boolean;
  getValidationList: (row: number, column: number) => string[] | undefined;
  setFillPreview: React.Dispatch<React.SetStateAction<CanvasFillPreview | null>>;
  setFilterPopover: React.Dispatch<React.SetStateAction<CanvasFilterPopoverState | null>>;
  setValidationDropdown: React.Dispatch<React.SetStateAction<CanvasValidationDropdownState | null>>;
  setContextMenu: React.Dispatch<React.SetStateAction<CanvasContextMenuState>>;
  setContextHit: React.Dispatch<React.SetStateAction<ResolvedContextHit | null>>;
}

function resolveDragCell(
  engine: CanvasRenderEngine,
  sheet: CanvasSheetSnapshot,
  local: { x: number; y: number },
  drag: DragState,
): { row: number; column: number } | null {
  const headerHit = engine.headerHitAtLocal(local);
  const hitCell = engine.cellAtLocalPoint(local);
  const isRowDrag = drag.floating?.id === "row";
  const isColDrag = drag.floating?.id === "col";
  if (isRowDrag) {
    if (headerHit?.kind === "row") return { row: headerHit.index, column: 0 };
    return hitCell ? { row: hitCell.row, column: 0 } : null;
  }
  if (isColDrag) {
    if (headerHit?.kind === "col") return { row: 0, column: headerHit.index };
    return hitCell ? { row: 0, column: hitCell.column } : null;
  }
  return hitCell ? resolveSelectionTarget(sheet, hitCell, 'cells', sheet.id).cell : null;
}

function toChromeSelection(selection: SelectionState): ChromeState["selection"] {
  return {
    ranges: selection.ranges.map((range) => ({
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    })),
    primary: { row: selection.activeCell.row, column: selection.activeCell.column },
    primaryIndex: selection.primaryRangeIndex,
  };
}

export function useCanvasInteraction(options: CanvasInteractionOptions) {
  const {
    canRepeat,
    chromeState,
    containerRef,
    contextRangeRef,
    drawings,
    drawingPayloads,
    drawingSelectionMode,
    textBoxPlacementActive,
    editingCell,
    engineRef,
    findPivotProjectionCell,
    floatables,
    formatPainterActive,
    formulaDraft,
    getValidationList,
    isPivotValueCell,
    onAppendFormulaDraft,
    onBeginEdit,
    onCancelEdit,
    onCancelFormatPainter,
    onCommitEdit,
    onExitDrawingSelectionMode,
    onExtendSelection,
    onFillRange,
    onFloatingMove,
    onFloatingSelect,
    onTextBoxPlacementCommit,
    onCancelTextBoxPlacement,
    onBeginTextBoxEdit,
    onToggleCheckbox,
    onFormulaDraftChange,
    editComposing,
    onJumpEdge,
    onMovePrimary,
    onRequestExtentGrowth,
    onPivotContextHit,
    onPivotControlAction,
    onPivotResolve,
    onPivotShowDetails,
    onPivotExpansionToggle,
    onResizeColumn,
    onAutoFitColumn,
    onAutoFitRow,
    formatColumnWidthPreview,
    onResizeRow,
    onSelectAll,
    onSelectionChange,
    onShortcut,
    onToggleAbsolute,
    onToggleOutline,
    phase,
    selectedFloatingId,
    selection,
    setContextHit,
    setContextMenu,
    setFillPreview,
    setFilterPopover,
    setValidationDropdown,
    sheet,
    sheetId,
    skeleton,
    zoom,
  } = options;
  const dragRef = useRef<DragState | null>(null);
  const shortcutRegistryRef = useRef<ReturnType<typeof createSpreadsheetShortcutRegistry> | null>(null);
  if (!shortcutRegistryRef.current) shortcutRegistryRef.current = createSpreadsheetShortcutRegistry();
  const editingActiveRef = useRef(false);
  const autoScrollFrameRef = useRef<number | null>(null);
  const autoScrollPointRef = useRef<{ x: number; y: number } | null>(null);
  const transientSelectionRef = useRef<{ gesture: SelectionGesture; sheetId: string } | null>(null);
  const transientSelectionFrameRef = useRef<number | null>(null);

  useEffect(() => {
    editingActiveRef.current = Boolean(editingCell);
  }, [editingCell]);

  const localPointOf = useCallback((event: { clientX: number; clientY: number }) => {
    const host = containerRef.current;
    if (!host) return { x: 0, y: 0 };
    const bounds = host.getBoundingClientRect();
    return { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
  }, [containerRef]);

  const filterPopoverAnchor = useCallback((column: number, fallback: { x: number; y: number }): { x: number; y: number } => {
    const engine = engineRef.current;
    const host = containerRef.current;
    const button = sheet.filterButtons.find((entry) => entry.column === column);
    if (!engine || !host || !button) return fallback;
    const rect = skeleton.getCellRect(button.row, button.column);
    if (!rect) return fallback;
    const bounds = host.getBoundingClientRect();
    const topLeft = engine.contentToScreen({ x: rect.x, y: rect.y }, { row: button.row, column: button.column });
    return { x: bounds.left + topLeft.x, y: bounds.top + topLeft.y + rect.height };
  }, [containerRef, engineRef, sheet.filterButtons, skeleton]);

  const stopAutoScroll = useCallback(() => {
    autoScrollPointRef.current = null;
    if (autoScrollFrameRef.current === null) return;
    if (typeof window !== "undefined") window.cancelAnimationFrame(autoScrollFrameRef.current);
    autoScrollFrameRef.current = null;
  }, []);

  const queueTransientSelection = useCallback((gesture: SelectionGesture, gestureSheetId: string) => {
    transientSelectionRef.current = { gesture, sheetId: gestureSheetId };
    if (transientSelectionFrameRef.current !== null) return;
    const draw = () => {
      transientSelectionFrameRef.current = null;
      const transient = transientSelectionRef.current;
      const engine = engineRef.current;
      if (!transient || !engine) return;
      const preview = selectionFromGesture(selection, transient.gesture, transient.sheetId);
      engine.setChrome({ ...chromeState, selection: toChromeSelection(preview) });
    };
    if (typeof window !== "undefined" && typeof window.requestAnimationFrame === "function") {
      transientSelectionFrameRef.current = window.requestAnimationFrame(draw);
    } else {
      transientSelectionFrameRef.current = setTimeout(draw, 0) as unknown as number;
    }
  }, [chromeState, engineRef, selection]);

  const clearTransientSelection = useCallback(() => {
    transientSelectionRef.current = null;
    if (transientSelectionFrameRef.current !== null) {
      if (typeof window !== "undefined") window.cancelAnimationFrame(transientSelectionFrameRef.current);
      transientSelectionFrameRef.current = null;
    }
    const engine = engineRef.current;
    if (engine) engine.setChrome(chromeState);
  }, [chromeState, engineRef]);

  const updateAutoScroll = useCallback((local: { x: number; y: number }) => {
    const host = containerRef.current;
    const engine = engineRef.current;
    const drag = dragRef.current;
    if (!host || !engine || !drag || (drag.kind !== "select" && drag.kind !== "fill")) {
      stopAutoScroll();
      return;
    }
    autoScrollPointRef.current = local;
    const origin = engine.headerOffset;
    const threshold = 24;
    const edge = local.x <= origin.x + threshold || local.x >= host.clientWidth - threshold
      || local.y <= origin.y + threshold || local.y >= host.clientHeight - threshold;
    if (!edge || typeof window === "undefined") {
      stopAutoScroll();
      return;
    }
    if (autoScrollFrameRef.current !== null) return;
    const tick = () => {
      autoScrollFrameRef.current = null;
      const currentDrag = dragRef.current;
      const point = autoScrollPointRef.current;
      const currentEngine = engineRef.current;
      const currentHost = containerRef.current;
      if (!currentDrag || !point || !currentEngine || !currentHost || (currentDrag.kind !== "select" && currentDrag.kind !== "fill")) return;
      const currentOrigin = currentEngine.headerOffset;
      const left = point.x < currentOrigin.x + threshold;
      const right = point.x > currentHost.clientWidth - threshold;
      const top = point.y < currentOrigin.y + threshold;
      const bottom = point.y > currentHost.clientHeight - threshold;
      const speed = (distance: number) => Math.max(2, Math.min(24, Math.round(distance / threshold * 24)));
      const dx = left ? -speed(currentOrigin.x + threshold - point.x) : right ? speed(point.x - (currentHost.clientWidth - threshold)) : 0;
      const dy = top ? -speed(currentOrigin.y + threshold - point.y) : bottom ? speed(point.y - (currentHost.clientHeight - threshold)) : 0;
      const viewport = currentEngine.viewport.getSnapshot();
      const content = currentEngine.skeleton.contentSize;
      const { rows, columns } = resolveAutoScrollExtentGrowth({
        right,
        bottom,
        viewport,
        content,
        defaultRowHeight: currentEngine.skeleton.defaultRowHeight,
        defaultColumnWidth: currentEngine.skeleton.defaultColumnWidth,
      });
      if (rows || columns) onRequestExtentGrowth({ rows, columns });
      if (dx !== 0 || dy !== 0) currentEngine.scrollBy(dx, dy);
      const queryPoint = {
        x: Math.max(currentOrigin.x + 1, Math.min(currentHost.clientWidth - 1, point.x)),
        y: Math.max(currentOrigin.y + 1, Math.min(currentHost.clientHeight - 1, point.y)),
      };
      const cell = resolveDragCell(currentEngine, sheet, queryPoint, currentDrag);
      if (cell) {
        currentDrag.currentRow = cell.row;
        currentDrag.currentColumn = cell.column;
        if (currentDrag.kind === "fill") {
          setFillPreview({
            startRow: Math.min(currentDrag.startRow, currentDrag.currentRow),
            endRow: Math.max(currentDrag.anchorRow, currentDrag.currentRow),
            startColumn: Math.min(currentDrag.startColumn, currentDrag.currentColumn),
            endColumn: Math.max(currentDrag.anchorColumn, currentDrag.currentColumn),
          });
        }
      } else {
        const isRowDrag = currentDrag.floating?.id === "row";
        const isColDrag = currentDrag.floating?.id === "col";
        const baseRange: RangeRef = {
          sheetId,
          startRow: isRowDrag || isColDrag ? Math.min(currentDrag.startRow, currentDrag.currentRow) : Math.min(currentDrag.anchorRow, currentDrag.currentRow),
          endRow: isRowDrag ? Math.max(currentDrag.startRow, currentDrag.currentRow) : isColDrag ? Math.max(0, skeleton.rowCount - 1) : Math.max(currentDrag.anchorRow, currentDrag.currentRow),
          startColumn: isColDrag ? Math.min(currentDrag.startColumn, currentDrag.currentColumn) : Math.min(currentDrag.anchorColumn, currentDrag.currentColumn),
          endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : isColDrag ? Math.max(currentDrag.startColumn, currentDrag.currentColumn) : Math.max(currentDrag.anchorColumn, currentDrag.currentColumn),
        };
        const range = expandSelectionRangeForMerges(sheet, baseRange);
        queueTransientSelection({
          origin: { row: currentDrag.anchorRow, column: currentDrag.anchorColumn },
          target: { row: currentDrag.currentRow, column: currentDrag.currentColumn },
          pointerId: currentDrag.pointerId,
          kind: isRowDrag ? 'rows' : isColDrag ? 'columns' : 'cells',
          additive: currentDrag.additive,
          expandedRange: range,
        }, sheetId);
      }
      if (autoScrollPointRef.current && typeof window !== "undefined") autoScrollFrameRef.current = window.requestAnimationFrame(tick);
    };
    autoScrollFrameRef.current = window.requestAnimationFrame(tick);
  }, [containerRef, engineRef, onRequestExtentGrowth, queueTransientSelection, selection, setFillPreview, sheet, sheetId, skeleton, stopAutoScroll]);

  useEffect(() => () => {
    stopAutoScroll();
    const host = containerRef.current;
    if (host) releasePointerGesturesForSurface(host.ownerDocument, host);
    if (transientSelectionFrameRef.current === null) return;
    if (typeof window !== "undefined") window.cancelAnimationFrame(transientSelectionFrameRef.current);
    transientSelectionFrameRef.current = null;
  }, [containerRef, stopAutoScroll]);

  const handlePointerDown = useCallback((event: React.PointerEvent) => {
    if (phase !== "ready") return;
    if (event.button === 2) return;
    if (resolvePointerGestureOwner(event.target) !== "worksheet") return;
    stopAutoScroll();
    setFillPreview(null);
    const engine = engineRef.current;
    const host = containerRef.current;
    if (!engine || !host) return;
    if (!claimPointerGesture(host.ownerDocument, event.pointerId, "worksheet", host)) return;
    host.setPointerCapture(event.pointerId);
    host.focus();
    if (editingCell || editingActiveRef.current) {
      editingActiveRef.current = false;
      onCommitEdit("none");
    }
    setFilterPopover(null);
    setValidationDropdown(null);
    const local = localPointOf(event);
    onPivotContextHit?.(null);

    if (textBoxPlacementActive) {
      if (!engine.cellAtLocalPoint(local)) return;
      const content = engine.localToContent(local);
      dragRef.current = {
        kind: "textbox-placement",
        pointerId: event.pointerId,
        startRow: 0,
        startColumn: 0,
        anchorRow: 0,
        anchorColumn: 0,
        currentRow: 0,
        currentColumn: 0,
        additive: false,
        extend: false,
        textBox: { startContent: content },
      };
      (event.target as Element).setPointerCapture?.(event.pointerId);
      return;
    }

    const floatingHit = engine.hitTestFloating(local);
    if (floatingHit) {
      if (floatingHit.control) {
        const action = floatingHit.control.data;
        if (action && typeof action === 'object' && 'kind' in action) {
          onPivotControlAction?.(floatingHit.id, action as PivotControlAction);
        }
        // A malformed/unsupported child hit is fail-closed: it must never
        // fall through into a move gesture and mutate drawing geometry.
        onFloatingSelect({ kind: floatingHit.kind, id: floatingHit.id }, event.shiftKey ? "add" : event.ctrlKey || event.metaKey ? "toggle" : "replace");
        (event.target as Element).setPointerCapture?.(event.pointerId);
        return;
      }
      const drawableBounds = floatables.find((item) => item.id === floatingHit.id)?.bounds;
      if (drawableBounds) {
        dragRef.current = {
          kind: floatingHit.handle ? "floating-resize" : "floating-move",
          pointerId: event.pointerId,
          startRow: 0,
          startColumn: 0,
          anchorRow: 0,
          anchorColumn: 0,
          currentRow: 0,
          currentColumn: 0,
          additive: false,
          extend: false,
          floating: {
            id: floatingHit.id,
            kind: floatingHit.kind,
            handle: floatingHit.handle,
            rotation: drawings.find((drawing) => drawing.id === floatingHit.id)?.transform.rotation,
            startBounds: { ...drawableBounds },
            startLocal: local,
          },
        };
        onFloatingSelect(floatingHit, event.shiftKey ? "add" : event.ctrlKey || event.metaKey ? "toggle" : "replace");
        (event.target as Element).setPointerCapture?.(event.pointerId);
        return;
      }
    }
    onFloatingSelect(null);
    if (drawingSelectionMode) return;

    const headerHit = engine.headerHitAtLocal(local);
    if (headerHit) {
      if (headerHit.kind === "corner") {
        onSelectAll();
        return;
      }
      if (headerHit.resizeBoundaryPx !== undefined) {
        dragRef.current = {
          kind: headerHit.kind === "col" ? "col-resize" : "row-resize",
          pointerId: event.pointerId,
          startRow: 0,
          startColumn: 0,
          anchorRow: 0,
          anchorColumn: 0,
          currentRow: 0,
          currentColumn: 0,
          additive: false,
          extend: false,
          dimensionResize: beginDimensionResizeGesture({
            axis: headerHit.kind === "col" ? "column" : "row",
            boundaryIndex: headerHit.index,
            startModelSizePx: (headerHit.kind === "col" ? skeleton.getColumnWidth(headerHit.index) : skeleton.getRowHeight(headerHit.index)) / (zoom / 100),
            startPointerScreenPx: headerHit.kind === "col" ? local.x : local.y,
            zoomScale: zoom / 100,
            minimumModelSizePx: headerHit.kind === "col" ? 1 : 18,
          }),
        };
        (event.target as Element).setPointerCapture?.(event.pointerId);
        return;
      }
      const additive = event.ctrlKey || event.metaKey;
      if (headerHit.kind === "row") {
        for (const control of sheet.outlineControls) {
          if (control.axis !== "row" || control.index !== headerHit.index) continue;
          const buttonLeft = 4 + (control.level - 1) * 10;
          if (local.x >= buttonLeft && local.x <= buttonLeft + 10) {
            onToggleOutline?.(control.groupId);
            return;
          }
        }
      }
      if (headerHit.kind === "col") {
        for (const control of sheet.outlineControls) {
          if (control.axis !== "column" || control.index !== headerHit.index) continue;
          const buttonTop = 2 + (control.level - 1) * 10;
          if (local.y >= buttonTop && local.y <= buttonTop + 10 && local.x >= 2 && local.x <= 12) {
            onToggleOutline?.(control.groupId);
            return;
          }
        }
        if (sheet.filterRangeColumns.includes(headerHit.index)) {
          const point = filterPopoverAnchor(headerHit.index, { x: event.clientX, y: event.clientY });
          setFilterPopover({ column: headerHit.index, ...point });
        }
      }
      dragRef.current = {
        kind: "select",
        pointerId: event.pointerId,
        startRow: headerHit.kind === "row" ? headerHit.index : 0,
        startColumn: headerHit.kind === "col" ? headerHit.index : 0,
        anchorRow: headerHit.kind === "row" ? headerHit.index : 0,
        anchorColumn: headerHit.kind === "col" ? headerHit.index : 0,
        currentRow: headerHit.kind === "row" ? headerHit.index : 0,
        currentColumn: headerHit.kind === "col" ? headerHit.index : 0,
        additive,
        extend: false,
        floating: { id: headerHit.kind, kind: "shape", handle: undefined, startBounds: { x: 0, y: 0, width: 0, height: 0 }, startLocal: { x: 0, y: 0 } },
      };
      (event.target as Element).setPointerCapture?.(event.pointerId);
      return;
    }

    const primaryRange = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
    const activePivotContextHit = onPivotResolve(sheet, selection.activeCell.row, selection.activeCell.column);
    if (primaryRange && !activePivotContextHit) {
      const rect = skeleton.getRangeRect({ startRow: primaryRange.endRow, endRow: primaryRange.endRow, startColumn: primaryRange.endColumn, endColumn: primaryRange.endColumn });
      if (rect) {
        const screen = engine.contentToScreen({ x: rect.x + rect.width, y: rect.y + rect.height }, { row: primaryRange.endRow, column: primaryRange.endColumn });
        if (Math.abs(local.x - screen.x) <= 5 && Math.abs(local.y - screen.y) <= 5) {
          dragRef.current = {
            kind: "fill",
            pointerId: event.pointerId,
            startRow: primaryRange.startRow,
            startColumn: primaryRange.startColumn,
            anchorRow: primaryRange.endRow,
            anchorColumn: primaryRange.endColumn,
            currentRow: primaryRange.endRow,
            currentColumn: primaryRange.endColumn,
            additive: false,
            extend: false,
          };
          (event.target as Element).setPointerCapture?.(event.pointerId);
          return;
        }
      }
    }

    const hitCell = engine.cellAtLocalPoint(local);
    if (!hitCell) return;
    const pivotContextHit = hitCell ? onPivotResolve(sheet, hitCell.row, hitCell.column) : null;
    if (pivotContextHit) {
      onPivotContextHit?.(pivotContextHit);
      const pivotTarget = findPivotProjectionCell(sheet, hitCell.row, hitCell.column);
      if (pivotTarget?.cell.kind === 'expand-toggle' && pivotTarget.cell.nodeId) {
        onPivotExpansionToggle(pivotTarget.projection.pivotId, pivotTarget.cell.nodeId);
        return;
      }
    }
    const cell = resolveSelectionTarget(sheet, hitCell, 'cells', sheet.id).cell;
    const checkboxRect = skeleton.getCellRect(cell.row, cell.column);
    const contentPoint = engine.localToContent(local);
    const checkboxSize = checkboxRect ? Math.min(14, Math.max(10, checkboxRect.height - 8)) : 0;
    const checkboxHit = Boolean(checkboxRect && sheet.getCell(cell.row, cell.column)?.editor?.kind === 'checkbox'
      && contentPoint.x >= checkboxRect.x + 2
      && contentPoint.x <= checkboxRect.x + 6 + checkboxSize
      && contentPoint.y >= checkboxRect.y + (checkboxRect.height - checkboxSize) / 2 - 2
      && contentPoint.y <= checkboxRect.y + (checkboxRect.height + checkboxSize) / 2 + 2);
    if (checkboxHit) {
      const checkboxRange = { sheetId, startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column };
      const additive = event.ctrlKey || event.metaKey;
      const extend = event.shiftKey && !additive;
      if (additive) {
        onSelectionChange(selectionFromGesture(selection, { origin: { row: cell.row, column: cell.column }, target: { row: cell.row, column: cell.column }, additive: true, expandedRange: checkboxRange }, sheetId));
      } else if (extend) {
        const range = {
          ...checkboxRange,
          startRow: Math.min(selection.anchorCell.row, cell.row),
          endRow: Math.max(selection.anchorCell.row, cell.row),
          startColumn: Math.min(selection.anchorCell.column, cell.column),
          endColumn: Math.max(selection.anchorCell.column, cell.column),
        };
        onSelectionChange(selectionFromGesture(selection, { origin: selection.anchorCell, target: { row: cell.row, column: cell.column }, expandedRange: range }, sheetId));
      } else {
        onSelectionChange(selectionFromGesture(selection, { origin: { row: cell.row, column: cell.column }, target: { row: cell.row, column: cell.column }, expandedRange: checkboxRange }, sheetId));
      }
      onToggleCheckbox([checkboxRange]);
      return;
    }
    const filterButton = sheet.filterButtons.find((button) => button.row === hitCell.row && button.column === hitCell.column);
    if (filterButton) {
      const cellRect = skeleton.getCellRect(hitCell.row, hitCell.column);
      if (cellRect) {
        const content = engine.localToContent(local);
        if (content.x >= cellRect.x + cellRect.width - 18) {
          const point = filterPopoverAnchor(hitCell.column, { x: event.clientX, y: event.clientY });
          setFilterPopover({ column: hitCell.column, ...point });
          return;
        }
      }
    }
    const additive = event.ctrlKey || event.metaKey;
    const extend = event.shiftKey && !additive;
    dragRef.current = {
      kind: "select",
      pointerId: event.pointerId,
      startRow: cell.row,
      startColumn: cell.column,
      anchorRow: extend ? selection.anchorCell.row : cell.row,
      anchorColumn: extend ? selection.anchorCell.column : cell.column,
      currentRow: cell.row,
      currentColumn: cell.column,
      additive,
      extend,
    };
    if (!additive && !extend) {
      onSelectionChange(selectionFromGesture(selection, { origin: { row: cell.row, column: cell.column }, target: { row: cell.row, column: cell.column }, expandedRange: { sheetId, startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column } }, sheetId));
    }
    (event.target as Element).setPointerCapture?.(event.pointerId);
  }, [containerRef, drawingSelectionMode, drawings, drawingPayloads, editingCell, engineRef, filterPopoverAnchor, findPivotProjectionCell, floatables, localPointOf, onBeginEdit, onCommitEdit, onFloatingSelect, onPivotContextHit, onPivotControlAction, onPivotExpansionToggle, onPivotResolve, onSelectAll, onSelectionChange, onToggleCheckbox, onToggleOutline, onTextBoxPlacementCommit, phase, selection, setFillPreview, setFilterPopover, setValidationDropdown, sheet, sheetId, skeleton, stopAutoScroll, textBoxPlacementActive]);

  const handlePointerMove = useCallback((event: React.PointerEvent) => {
    const engine = engineRef.current;
    if (!engine) return;
    const local = localPointOf(event);
    const drag = dragRef.current;
    if (!drag) {
      if (resolvePointerGestureOwner(event.target) !== "worksheet") return;
      stopAutoScroll();
      const headerHit = engine.headerHitAtLocal(local);
      const host = containerRef.current;
      if (host) host.style.cursor = headerHit?.resizeBoundaryPx !== undefined ? (headerHit.kind === "col" ? "col-resize" : "row-resize") : "default";
      return;
    }
    const host = containerRef.current;
    if (!host || drag.pointerId !== event.pointerId
      || !ownsPointerGesture(host.ownerDocument, event.pointerId, "worksheet", host)) return;
    if (drag.kind === "col-resize" || drag.kind === "row-resize") {
      setFillPreview(null);
      stopAutoScroll();
      if (!drag.dimensionResize) throw new Error('Dimension resize drag is missing its canonical gesture');
      drag.dimensionResize = updateDimensionResizeGesture(
        drag.dimensionResize,
        drag.kind === "col-resize" ? local.x : local.y,
      );
      const modelSizePx = drag.dimensionResize.currentModelSizePx;
      const size = modelSizePx * drag.dimensionResize.zoomScale;
      const columnPreview = drag.kind === 'col-resize' ? formatColumnWidthPreview(modelSizePx) : undefined;
      const label = drag.kind === 'col-resize'
        ? `${columnPreview!.excelWidth.toFixed(2)} chars (${columnPreview!.widthPx}px)`
        : `${modelSizePx}px`;
      engine.setChrome({ ...chromeState, resizePreview: { axis: drag.dimensionResize.axis, index: drag.dimensionResize.boundaryIndex, sizePx: size, label } });
      return;
    }
    if (drag.kind === "floating-move" && drag.floating) {
      setFillPreview(null);
      stopAutoScroll();
      const deltaX = local.x - drag.floating.startLocal.x;
      const deltaY = local.y - drag.floating.startLocal.y;
      onFloatingMove(drag.floating.id, { x: drag.floating.startBounds.x + deltaX, y: drag.floating.startBounds.y + deltaY, width: drag.floating.startBounds.width, height: drag.floating.startBounds.height }, drag.floating.rotation);
      return;
    }
    if (drag.kind === "floating-resize" && drag.floating?.handle) {
      setFillPreview(null);
      stopAutoScroll();
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
      onFloatingMove(drag.floating.id, { x, y, width, height }, drag.floating.rotation);
      return;
    }
    if (drag.kind === "textbox-placement" && drag.textBox) {
      stopAutoScroll();
      return;
    }
    updateAutoScroll(local);
    const headerHit = engine.headerHitAtLocal(local);
    const hitCell = engine.cellAtLocalPoint(local);
    const isRowDrag = drag.floating?.id === "row";
    const isColDrag = drag.floating?.id === "col";
    const cell = isRowDrag
      ? headerHit?.kind === "row" ? { row: headerHit.index, column: 0 } : hitCell ? { row: hitCell.row, column: 0 } : null
      : isColDrag
        ? headerHit?.kind === "col" ? { row: 0, column: headerHit.index } : hitCell ? { row: 0, column: hitCell.column } : null
        : hitCell ? resolveSelectionTarget(sheet, hitCell, 'cells', sheet.id).cell : null;
    if (!cell) return;
    if (drag.kind === "fill") {
      drag.currentRow = cell.row;
      drag.currentColumn = cell.column;
      setFillPreview({ startRow: Math.min(drag.startRow, drag.currentRow), endRow: Math.max(drag.anchorRow, drag.currentRow), startColumn: Math.min(drag.startColumn, drag.currentColumn), endColumn: Math.max(drag.anchorColumn, drag.currentColumn) });
      return;
    }
    drag.currentRow = cell.row;
    drag.currentColumn = cell.column;
    const baseRange: RangeRef = {
      sheetId,
      startRow: isRowDrag || isColDrag ? Math.min(drag.startRow, drag.currentRow) : Math.min(drag.anchorRow, drag.currentRow),
      endRow: isRowDrag ? Math.max(drag.startRow, drag.currentRow) : isColDrag ? Math.max(0, skeleton.rowCount - 1) : Math.max(drag.anchorRow, drag.currentRow),
      startColumn: isColDrag ? Math.min(drag.startColumn, drag.currentColumn) : Math.min(drag.anchorColumn, drag.currentColumn),
      endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : isColDrag ? Math.max(drag.startColumn, drag.currentColumn) : Math.max(drag.anchorColumn, drag.currentColumn),
    };
    const range = expandSelectionRangeForMerges(sheet, baseRange);
    queueTransientSelection({
      origin: { row: drag.anchorRow, column: drag.anchorColumn },
      target: { row: cell.row, column: cell.column },
      pointerId: drag.pointerId,
      kind: isRowDrag ? 'rows' : isColDrag ? 'columns' : 'cells',
      additive: drag.additive,
      expandedRange: range,
    }, sheetId);
  }, [chromeState, containerRef, engineRef, formatColumnWidthPreview, localPointOf, onFloatingMove, queueTransientSelection, selection, setFillPreview, sheet, sheetId, skeleton, stopAutoScroll, updateAutoScroll, zoom]);

  const handlePointerUp = useCallback((event: React.PointerEvent) => {
    const host = containerRef.current;
    if (!host || !ownsPointerGesture(host.ownerDocument, event.pointerId, "worksheet", host)) return;
    releasePointerGesture(host.ownerDocument, event.pointerId, "worksheet", host);
    const drag = dragRef.current;
    dragRef.current = null;
    const engine = engineRef.current;
    if (host.hasPointerCapture(event.pointerId)) host.releasePointerCapture(event.pointerId);
    if (!drag || !engine) return;
    if (drag.kind === "select" || drag.kind === "fill") {
      const finalCell = resolveDragCell(engine, sheet, localPointOf(event), drag);
      if (finalCell) onPivotContextHit?.(onPivotResolve(sheet, finalCell.row, finalCell.column));
    }
    stopAutoScroll();
    if (drag.kind === "col-resize") {
      if (!drag.dimensionResize) throw new Error('Column resize drag is missing its canonical gesture');
      onResizeColumn(drag.dimensionResize.boundaryIndex, drag.dimensionResize.currentModelSizePx);
      return;
    }
    if (drag.kind === "row-resize") {
      if (!drag.dimensionResize) throw new Error('Row resize drag is missing its canonical gesture');
      onResizeRow(drag.dimensionResize.boundaryIndex, drag.dimensionResize.currentModelSizePx);
      return;
    }
    if (drag.kind === "textbox-placement" && drag.textBox) {
      const end = engine.localToContent(localPointOf(event));
      const start = drag.textBox.startContent;
      const x = Math.min(start.x, end.x);
      const y = Math.min(start.y, end.y);
      const width = Math.max(40, Math.abs(end.x - start.x) || 220);
      const height = Math.max(30, Math.abs(end.y - start.y) || 72);
      onTextBoxPlacementCommit({ x, y, width, height });
      return;
    }
    if (drag.kind === "fill") {
      setFillPreview(null);
      const target: RangeRef = { sheetId, startRow: Math.min(drag.startRow, drag.currentRow), endRow: Math.max(drag.anchorRow, drag.currentRow), startColumn: Math.min(drag.startColumn, drag.currentColumn), endColumn: Math.max(drag.anchorColumn, drag.currentColumn) };
      const partialMerge = sheet.merges.some((merge) => intersectsRange(target, merge.range) && !containsRange(target, merge.range));
      if (!partialMerge && (target.endRow !== drag.anchorRow || target.endColumn !== drag.anchorColumn)) {
        onFillRange(target);
        onSelectionChange(selectionFromGesture(selection, {
          origin: { row: drag.startRow, column: drag.startColumn },
          target: { row: drag.currentRow, column: drag.currentColumn },
          expandedRange: target,
        }, sheetId));
      }
      return;
    }
    if (drag.kind === "select") {
      setFillPreview(null);
      if (drag.extend) {
        clearTransientSelection();
        onExtendSelection?.(drag.currentRow, drag.currentColumn);
        return;
      }
      const isRowDrag = drag.floating?.id === "row";
      const isColDrag = drag.floating?.id === "col";
      const baseRange: RangeRef = {
        sheetId,
        startRow: isRowDrag || isColDrag ? Math.min(drag.startRow, drag.currentRow) : Math.min(drag.anchorRow, drag.currentRow),
        endRow: isRowDrag ? Math.max(drag.startRow, drag.currentRow) : isColDrag ? Math.max(0, skeleton.rowCount - 1) : Math.max(drag.anchorRow, drag.currentRow),
        startColumn: isColDrag ? Math.min(drag.startColumn, drag.currentColumn) : Math.min(drag.anchorColumn, drag.currentColumn),
        endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : isColDrag ? Math.max(drag.startColumn, drag.currentColumn) : Math.max(drag.anchorColumn, drag.currentColumn),
      };
      const range = expandSelectionRangeForMerges(sheet, baseRange);
      const nextSelection = selectionFromGesture(selection, {
        origin: { row: drag.anchorRow, column: drag.anchorColumn },
        target: { row: drag.currentRow, column: drag.currentColumn },
        kind: isRowDrag ? 'rows' : isColDrag ? 'columns' : 'cells',
        additive: drag.additive,
        expandedRange: range,
      }, sheetId);
      clearTransientSelection();
      onSelectionChange(nextSelection);
    }
  }, [clearTransientSelection, containerRef, engineRef, localPointOf, onExtendSelection, onFillRange, onPivotContextHit, onPivotResolve, onResizeColumn, onResizeRow, onSelectionChange, selection, setFillPreview, sheet, sheetId, skeleton, stopAutoScroll, zoom]);

  const handlePointerCancel = useCallback((event: React.PointerEvent) => {
    const host = containerRef.current;
    if (!host || !ownsPointerGesture(host.ownerDocument, event.pointerId, "worksheet", host)) return;
    releasePointerGesture(host.ownerDocument, event.pointerId, "worksheet", host);
    dragRef.current = null;
    stopAutoScroll();
    setFillPreview(null);
    clearTransientSelection();
    if (host.hasPointerCapture(event.pointerId)) host.releasePointerCapture(event.pointerId);
  }, [clearTransientSelection, containerRef, setFillPreview, stopAutoScroll]);

  const handleDoubleClick = useCallback((event: React.PointerEvent | React.MouseEvent) => {
    const engine = engineRef.current;
    if (!engine) return;
    const local = localPointOf(event);
    onPivotContextHit?.(null);
    const floatingHit = engine.hitTestFloating(local);
    if (floatingHit) {
      if (floatingHit.control) return;
      const drawing = drawings.find((entry) => entry.id === floatingHit.id);
      const payload = drawing ? drawingPayloads.get(drawing.payloadId) : undefined;
      if (payload?.kind === 'textbox') {
        onBeginTextBoxEdit(floatingHit.id);
        return;
      }
    }
    const headerHit = engine.headerHitAtLocal(local);
    if (headerHit?.resizeBoundaryPx !== undefined) {
      if (headerHit.kind === 'col') void onAutoFitColumn(headerHit.index);
      else void onAutoFitRow(headerHit.index);
      return;
    }
    const hitCell = engine.cellAtLocalPoint(local);
    if (!hitCell) return;
    const pivotContextHit = onPivotResolve(sheet, hitCell.row, hitCell.column);
    if (pivotContextHit) {
      onPivotContextHit?.(pivotContextHit);
      const pivotTarget = findPivotProjectionCell(sheet, hitCell.row, hitCell.column);
      if (pivotTarget?.cell.kind === 'expand-toggle' && pivotTarget.cell.nodeId) return;
      if (pivotTarget && isPivotValueCell(pivotTarget.cell) && pivotTarget.cell.sourceRowPaths && pivotTarget.cell.sourceRowPaths.length > 0) {
        onPivotShowDetails({ pivotId: pivotTarget.projection.pivotId, sourceRowPaths: pivotTarget.cell.sourceRowPaths, hit: pivotContextHit });
      }
      return;
    }
    const cell = resolveSelectionTarget(sheet, hitCell, 'cells', sheet.id).cell;
    const editor = sheet.getCell(cell.row, cell.column)?.editor;
    if (editor?.kind === 'checkbox') {
      // Checkbox activation is owned by the glyph pointer and Spacebar. A
      // double-click never re-enters the generic text editor.
      return;
    }
    const validationList = getValidationList(cell.row, cell.column);
    if (validationList && validationList.length > 0) {
      setValidationDropdown({ row: cell.row, column: cell.column, options: validationList });
      return;
    }
    onBeginEdit();
  }, [drawingPayloads, drawings, engineRef, findPivotProjectionCell, getValidationList, isPivotValueCell, localPointOf, onAutoFitColumn, onAutoFitRow, onBeginEdit, onBeginTextBoxEdit, onPivotContextHit, onPivotExpansionToggle, onPivotResolve, onPivotShowDetails, setValidationDropdown, sheet]);

  const handleWheel = useCallback((event: React.WheelEvent) => {
    const engine = engineRef.current;
    if (!engine) return;
    event.preventDefault();
    engine.scrollBy(event.deltaX, event.deltaY);
  }, [engineRef]);

  const handleKeyDown = useCallback((event: React.KeyboardEvent) => {
    if (phase !== "ready") return;
    const key = event.key;
    const ctrl = event.ctrlKey || event.metaKey;
    const isEditing = Boolean(editingCell) || editingActiveRef.current;
    if (event.nativeEvent.isComposing || editComposing) {
      event.stopPropagation();
      return;
    }
    const activePivotContextHit = onPivotResolve(sheet, selection.activeCell.row, selection.activeCell.column);
    const editingPivotContextHit = editingCell ? onPivotResolve(sheet, editingCell.row, editingCell.column) : null;
    if (editingPivotContextHit) {
      event.preventDefault();
      editingActiveRef.current = false;
      onPivotContextHit?.(editingPivotContextHit);
      onCancelEdit();
      return;
    }
    if (activePivotContextHit && isEditing) {
      event.preventDefault();
      editingActiveRef.current = false;
      onPivotContextHit?.(activePivotContextHit);
      onCancelEdit();
      return;
    }
    if (isEditing) {
      if (key === "Escape") { event.preventDefault(); editingActiveRef.current = false; onCancelEdit(); return; }
      if (key === "Enter" && !event.shiftKey) { event.preventDefault(); editingActiveRef.current = false; onCommitEdit("down"); return; }
      if (key === "Tab") { event.preventDefault(); editingActiveRef.current = false; onCommitEdit(event.shiftKey ? "left" : "right"); return; }
      if (key === "F4") { event.preventDefault(); onToggleAbsolute(); return; }
      if (key.length === 1 && !ctrl && !event.altKey) {
        event.preventDefault();
        if (onAppendFormulaDraft) onAppendFormulaDraft(key);
        else onFormulaDraftChange(formulaDraft + key);
      }
      return;
    }
    if (key === "Escape" && formatPainterActive) { event.preventDefault(); onCancelFormatPainter?.(); return; }
    if (key === "Escape" && drawingSelectionMode) { event.preventDefault(); onExitDrawingSelectionMode?.(); return; }
    if (textBoxPlacementActive && key === "Escape") { event.preventDefault(); onCancelTextBoxPlacement(); return; }
    const activeScope = activePivotContextHit ? "pivot" as const : selectedFloatingId ? "drawing" as const : "grid" as const;
    const shortcut = shortcutRegistryRef.current?.resolve(event, { scope: activeScope, canRepeat }) ?? (activeScope === "grid" ? undefined : shortcutRegistryRef.current?.resolve(event, { scope: "grid", canRepeat }));
    if (shortcut?.id === "context.open") {
      event.preventDefault();
      const activeRange = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? { sheetId, startRow: selection.activeCell.row, endRow: selection.activeCell.row, startColumn: selection.activeCell.column, endColumn: selection.activeCell.column };
      contextRangeRef.current = activeRange;
      setContextHit(activePivotContextHit ?? resolveContextHit({
        sheetId,
        cell: {
          row: selection.activeCell.row,
          column: selection.activeCell.column,
          range: activeRange,
        },
      }));
      const bounds = containerRef.current?.getBoundingClientRect();
      setContextMenu({ x: (bounds?.left ?? 0) + 24, y: (bounds?.top ?? 0) + 24, open: true });
      if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit);
      return;
    }
    if (shortcut && onShortcut?.(shortcut.id)) { event.preventDefault(); return; }
    const checkboxRanges = selection.ranges.filter((range) => {
      for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        if (sheet.getCell(row, column)?.editor?.kind !== 'checkbox') return false;
      }
      return true;
    });
    const checkboxSelection = checkboxRanges.length === selection.ranges.length && checkboxRanges.length > 0;
    if (key === ' ' && checkboxSelection) { event.preventDefault(); onToggleCheckbox(selection.ranges); return; }
    if (key === "Enter") {
      event.preventDefault();
      if (selectedFloatingId) {
        const drawing = drawings.find((entry) => entry.id === selectedFloatingId);
        const payload = drawing ? drawingPayloads.get(drawing.payloadId) : undefined;
        if (payload?.kind === 'textbox') { onBeginTextBoxEdit(selectedFloatingId); return; }
      }
      const activePivotCell = activePivotContextHit ? findPivotProjectionCell(sheet, selection.activeCell.row, selection.activeCell.column)?.cell : undefined;
      const activePivotProjection = activePivotContextHit ? findPivotProjectionCell(sheet, selection.activeCell.row, selection.activeCell.column)?.projection : undefined;
      if (activePivotCell?.kind === 'expand-toggle' && activePivotCell.nodeId && activePivotProjection) onPivotExpansionToggle(activePivotProjection.pivotId, activePivotCell.nodeId);
      else if (checkboxSelection) onMovePrimary(1, 0);
      else if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit);
      else onBeginEdit();
      return;
    }
    const moves: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, event.shiftKey ? -1 : 1] };
    if (key in moves && !ctrl) { event.preventDefault(); const [dr, dc] = moves[key]!; onMovePrimary(dr, dc, { extend: event.shiftKey }); return; }
    if (key in moves && ctrl) { event.preventDefault(); const direction = key === "ArrowUp" ? "up" : key === "ArrowDown" ? "down" : key === "ArrowLeft" ? "left" : "right"; onJumpEdge(direction, event.shiftKey); return; }
    if (key === "Home") { event.preventDefault(); onMovePrimary(0, -selection.activeCell.column, { extend: event.shiftKey }); return; }
    if (key === "PageDown" || key === "PageUp") {
      event.preventDefault();
      const rowHeight = Math.max(1, skeleton.getRowHeight(selection.activeCell.row));
      const rows = Math.max(1, Math.floor((containerRef.current?.clientHeight ?? 600) / rowHeight) - 2);
      onMovePrimary(key === "PageDown" ? rows : -rows, 0, { extend: event.shiftKey });
      return;
    }
    if (key.length === 1 && !ctrl && !event.altKey) {
      event.preventDefault();
      if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit);
      else if (selectedFloatingId) {
        const drawing = drawings.find((entry) => entry.id === selectedFloatingId);
        const payload = drawing ? drawingPayloads.get(drawing.payloadId) : undefined;
        if (payload?.kind === 'textbox') { onBeginTextBoxEdit(selectedFloatingId, key); return; }
      }
      else if (editingCell || editingActiveRef.current) onAppendFormulaDraft?.(key);
      else { editingActiveRef.current = true; onBeginEdit(key); }
    }
  }, [canRepeat, containerRef, contextRangeRef, drawingPayloads, drawings, drawingSelectionMode, editComposing, editingCell, findPivotProjectionCell, formatPainterActive, formulaDraft, onAppendFormulaDraft, onBeginEdit, onBeginTextBoxEdit, onCancelEdit, onCancelFormatPainter, onCancelTextBoxPlacement, onCommitEdit, onExitDrawingSelectionMode, onFormulaDraftChange, onJumpEdge, onMovePrimary, onPivotContextHit, onPivotControlAction, onPivotExpansionToggle, onPivotResolve, onShortcut, onToggleAbsolute, onToggleCheckbox, onTextBoxPlacementCommit, phase, selectedFloatingId, selection, setContextHit, setContextMenu, sheet, sheetId, skeleton, textBoxPlacementActive]);

  return {
    clearTransientSelection,
    dragRef,
    handleDoubleClick,
    handleKeyDown,
    handlePointerDown,
    handlePointerMove,
    handlePointerCancel,
    handlePointerUp,
    handleWheel,
    localPointOf,
  };
}
