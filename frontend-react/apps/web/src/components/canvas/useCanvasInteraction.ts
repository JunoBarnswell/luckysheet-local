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
  PivotProjectionCell,
  PivotSourceRowPath,
  RangeRef,
} from "@react-sheets/core-model";
import {
  createSpreadsheetShortcutRegistry,
  resolveContextHit,
  type AppPhase,
  type CanvasSheetSnapshot,
  type ResolvedContextHit,
  type SelectionState,
} from "@react-sheets/spreadsheet-app";

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
  kind: "select" | "fill" | "col-resize" | "row-resize" | "floating-move" | "floating-resize";
  startRow: number;
  startColumn: number;
  anchorRow: number;
  anchorColumn: number;
  currentRow: number;
  currentColumn: number;
  additive: boolean;
  extend: boolean;
  resizeStartSize: number;
  resizeIndex: number;
  floating?: {
    id: string;
    kind: "chart" | "shape" | "image";
    handle?: FloatingHandle;
    rotation?: number;
    startBounds: Rect;
    startLocal: { x: number; y: number };
  };
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
  formatPainterActive: boolean;
  canRepeat: boolean;
  onPivotContextHit?: (hit: ResolvedContextHit | null) => void;
  onPivotShowDetails?: (request: {
    pivotId: string;
    sourceRowPaths: readonly PivotSourceRowPath[];
    hit: ResolvedContextHit;
  }) => void;
  onSelectionChange: (selection: SelectionState) => void;
  onMovePrimary: (rowDelta: number, columnDelta: number, opts?: { extend?: boolean }) => void;
  onCommitCell: (value: string) => void;
  onBeginEdit: (initialText?: string) => void;
  onCancelEdit: () => void;
  onCommitEdit: (moveAfter?: "down" | "up" | "left" | "right" | "none") => void;
  onFormulaDraftChange: (value: string) => void;
  onAppendFormulaDraft?: (fragment: string) => void;
  onToggleAbsolute: () => void;
  onJumpEdge: (direction: "up" | "down" | "left" | "right", extend?: boolean) => void;
  onSelectAll: () => void;
  onExtendSelection?: (row: number, column: number) => void;
  onResizeRow: (row: number, heightPx: number) => void;
  onResizeColumn: (column: number, widthPx: number) => void;
  onFillRange: (target: CanvasFillPreview) => void;
  onExitDrawingSelectionMode?: () => void;
  onFloatingSelect: (hit: FloatingHit | null, mode?: "replace" | "add" | "toggle") => void;
  onFloatingMove: (drawingId: string, bounds: Rect, rotation?: number) => void;
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

function mergeAtCell(sheet: CanvasSheetSnapshot, cell: { row: number; column: number }) {
  return sheet.merges.find((merge) =>
    cell.row >= merge.range.startRow && cell.row <= merge.range.endRow
    && cell.column >= merge.range.startColumn && cell.column <= merge.range.endColumn);
}

function resolveMergedCell(sheet: CanvasSheetSnapshot, cell: { row: number; column: number }): { row: number; column: number } {
  const merge = mergeAtCell(sheet, cell);
  return merge ? { ...merge.anchor } : { ...cell };
}

function intersectsRange(
  first: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  second: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  return first.startRow <= second.endRow && second.startRow <= first.endRow
    && first.startColumn <= second.endColumn && second.startColumn <= first.endColumn;
}

function containsRange(
  outer: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  inner: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  return outer.startRow <= inner.startRow && outer.endRow >= inner.endRow
    && outer.startColumn <= inner.startColumn && outer.endColumn >= inner.endColumn;
}

function expandRangeForMerges(sheet: CanvasSheetSnapshot, range: RangeRef): RangeRef {
  let expanded = { ...range };
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of sheet.merges) {
      if (!intersectsRange(expanded, merge.range)) continue;
      const next = {
        startRow: Math.min(expanded.startRow, merge.range.startRow),
        endRow: Math.max(expanded.endRow, merge.range.endRow),
        startColumn: Math.min(expanded.startColumn, merge.range.startColumn),
        endColumn: Math.max(expanded.endColumn, merge.range.endColumn),
      };
      changed = next.startRow !== expanded.startRow || next.endRow !== expanded.endRow
        || next.startColumn !== expanded.startColumn || next.endColumn !== expanded.endColumn;
      expanded = { ...expanded, ...next };
    }
  }
  return expanded;
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
  return hitCell ? resolveMergedCell(sheet, hitCell) : null;
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
    drawingSelectionMode,
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
    onCommitCell,
    onExitDrawingSelectionMode,
    onExtendSelection,
    onFillRange,
    onFloatingMove,
    onFloatingSelect,
    onFormulaDraftChange,
    onJumpEdge,
    onMovePrimary,
    onPivotContextHit,
    onPivotResolve,
    onPivotShowDetails,
    onResizeColumn,
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
  const transientSelectionRef = useRef<SelectionState | null>(null);
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

  const stopAutoScroll = useCallback(() => {
    autoScrollPointRef.current = null;
    if (autoScrollFrameRef.current === null) return;
    if (typeof window !== "undefined") window.cancelAnimationFrame(autoScrollFrameRef.current);
    autoScrollFrameRef.current = null;
  }, []);

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
    if (typeof window !== "undefined" && typeof window.requestAnimationFrame === "function") {
      transientSelectionFrameRef.current = window.requestAnimationFrame(draw);
    } else {
      transientSelectionFrameRef.current = setTimeout(draw, 0) as unknown as number;
    }
  }, [chromeState, engineRef]);

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
        const range = expandRangeForMerges(sheet, baseRange);
        const nextSelection: SelectionState = {
          ranges: [range],
          primaryRangeIndex: 0,
          activeCell: { row: currentDrag.currentRow, column: currentDrag.currentColumn },
          anchorCell: { row: currentDrag.anchorRow, column: currentDrag.anchorColumn },
        };
        queueTransientSelection(currentDrag.additive
          ? { ...selection, ranges: [...selection.ranges, range], primaryRangeIndex: selection.ranges.length, activeCell: nextSelection.activeCell, anchorCell: nextSelection.anchorCell }
          : nextSelection);
      }
      if (autoScrollPointRef.current && typeof window !== "undefined") autoScrollFrameRef.current = window.requestAnimationFrame(tick);
    };
    autoScrollFrameRef.current = window.requestAnimationFrame(tick);
  }, [containerRef, engineRef, queueTransientSelection, selection, setFillPreview, sheet, sheetId, skeleton, stopAutoScroll]);

  useEffect(() => () => {
    stopAutoScroll();
    if (transientSelectionFrameRef.current === null) return;
    if (typeof window !== "undefined") window.cancelAnimationFrame(transientSelectionFrameRef.current);
    transientSelectionFrameRef.current = null;
  }, [stopAutoScroll]);

  const handlePointerDown = useCallback((event: React.PointerEvent) => {
    if (phase !== "ready") return;
    if (event.button === 2) return;
    if ((event.target as Element).closest('[aria-label="Cell editor"]')) return;
    stopAutoScroll();
    setFillPreview(null);
    const engine = engineRef.current;
    const host = containerRef.current;
    if (!engine || !host) return;
    host.focus();
    if (editingCell || editingActiveRef.current) {
      editingActiveRef.current = false;
      onCommitEdit("none");
    }
    setFilterPopover(null);
    setValidationDropdown(null);
    const local = localPointOf(event);
    onPivotContextHit?.(null);

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
          extend: false,
          resizeStartSize: 0,
          resizeIndex: 0,
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
          startRow: 0,
          startColumn: 0,
          anchorRow: 0,
          anchorColumn: 0,
          currentRow: 0,
          currentColumn: 0,
          additive: false,
          extend: false,
          resizeStartSize: headerHit.kind === "col" ? skeleton.getColumnWidth(headerHit.index) : skeleton.getRowHeight(headerHit.index),
          resizeIndex: headerHit.index,
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
        if (sheet.filterColumns.includes(headerHit.index)) setFilterPopover({ column: headerHit.index, x: event.clientX, y: event.clientY });
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
        extend: false,
        resizeStartSize: 0,
        resizeIndex: 0,
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
        const screen = engine.contentToMainScreen({ x: rect.x + rect.width, y: rect.y + rect.height }, { row: primaryRange.endRow, column: primaryRange.endColumn });
        if (Math.abs(local.x - screen.x) <= 5 && Math.abs(local.y - screen.y) <= 5) {
          dragRef.current = {
            kind: "fill",
            startRow: primaryRange.startRow,
            startColumn: primaryRange.startColumn,
            anchorRow: primaryRange.endRow,
            anchorColumn: primaryRange.endColumn,
            currentRow: primaryRange.endRow,
            currentColumn: primaryRange.endColumn,
            additive: false,
            extend: false,
            resizeStartSize: 0,
            resizeIndex: 0,
          };
          (event.target as Element).setPointerCapture?.(event.pointerId);
          return;
        }
      }
    }

    const hitCell = engine.cellAtLocalPoint(local);
    const pivotContextHit = hitCell ? onPivotResolve(sheet, hitCell.row, hitCell.column) : null;
    if (pivotContextHit) onPivotContextHit?.(pivotContextHit);
    if (!hitCell) return;
    const cell = resolveMergedCell(sheet, hitCell);
    const filterButton = sheet.filterButtons.find((button) => button.row === hitCell.row && button.column === hitCell.column);
    if (filterButton) {
      const cellRect = skeleton.getCellRect(hitCell.row, hitCell.column);
      if (cellRect) {
        const content = engine.localToContent(local);
        if (content.x >= cellRect.x + cellRect.width - 18) {
          setFilterPopover({ column: hitCell.column, x: event.clientX, y: event.clientY });
          return;
        }
      }
    }
    const additive = event.ctrlKey || event.metaKey;
    const extend = event.shiftKey && !additive;
    dragRef.current = {
      kind: "select",
      startRow: cell.row,
      startColumn: cell.column,
      anchorRow: extend ? selection.anchorCell.row : cell.row,
      anchorColumn: extend ? selection.anchorCell.column : cell.column,
      currentRow: cell.row,
      currentColumn: cell.column,
      additive,
      extend,
      resizeStartSize: 0,
      resizeIndex: 0,
    };
    if (!additive && !extend) {
      onSelectionChange({ ranges: [{ sheetId, startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column }], primaryRangeIndex: 0, activeCell: { row: cell.row, column: cell.column }, anchorCell: { row: cell.row, column: cell.column } });
    }
    (event.target as Element).setPointerCapture?.(event.pointerId);
  }, [containerRef, drawingSelectionMode, drawings, editingCell, engineRef, floatables, localPointOf, onBeginEdit, onCommitEdit, onFloatingSelect, onPivotContextHit, onPivotResolve, onSelectAll, onSelectionChange, onToggleOutline, phase, selection, setFillPreview, setFilterPopover, setValidationDropdown, sheet, sheetId, skeleton, stopAutoScroll]);

  const handlePointerMove = useCallback((event: React.PointerEvent) => {
    const engine = engineRef.current;
    if (!engine) return;
    const local = localPointOf(event);
    const drag = dragRef.current;
    if (!drag) {
      stopAutoScroll();
      const headerHit = engine.headerHitAtLocal(local);
      const host = containerRef.current;
      if (host) host.style.cursor = headerHit?.resizeBoundaryPx !== undefined ? (headerHit.kind === "col" ? "col-resize" : "row-resize") : "default";
      return;
    }
    if (drag.kind === "col-resize" || drag.kind === "row-resize") {
      setFillPreview(null);
      stopAutoScroll();
      const content = engine.localToContent(local);
      const boundary = drag.kind === "col-resize" ? skeleton.getColumnLeft(drag.resizeIndex) : skeleton.getRowTop(drag.resizeIndex);
      const size = Math.max(24, (drag.kind === "col-resize" ? content.x : content.y) - boundary);
      engine.setChrome({ ...chromeState, resizePreview: { axis: drag.kind === "col-resize" ? "column" : "row", index: drag.resizeIndex, sizePx: size } });
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
    updateAutoScroll(local);
    const headerHit = engine.headerHitAtLocal(local);
    const hitCell = engine.cellAtLocalPoint(local);
    const isRowDrag = drag.floating?.id === "row";
    const isColDrag = drag.floating?.id === "col";
    const cell = isRowDrag
      ? headerHit?.kind === "row" ? { row: headerHit.index, column: 0 } : hitCell ? { row: hitCell.row, column: 0 } : null
      : isColDrag
        ? headerHit?.kind === "col" ? { row: 0, column: headerHit.index } : hitCell ? { row: 0, column: hitCell.column } : null
        : hitCell ? resolveMergedCell(sheet, hitCell) : null;
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
    const range = expandRangeForMerges(sheet, baseRange);
    const nextSelection: SelectionState = { ranges: [range], activeCell: { row: cell.row, column: cell.column }, primaryRangeIndex: 0, anchorCell: { row: drag.anchorRow, column: drag.anchorColumn } };
    queueTransientSelection(drag.additive ? { ...selection, ranges: [...selection.ranges, range], primaryRangeIndex: selection.ranges.length, activeCell: nextSelection.activeCell, anchorCell: nextSelection.anchorCell } : nextSelection);
  }, [chromeState, containerRef, engineRef, localPointOf, onFloatingMove, queueTransientSelection, selection, setFillPreview, sheet, sheetId, skeleton, stopAutoScroll, updateAutoScroll]);

  const handlePointerUp = useCallback((event: React.PointerEvent) => {
    const drag = dragRef.current;
    dragRef.current = null;
    const engine = engineRef.current;
    if (!drag || !engine) return;
    (event.target as Element).releasePointerCapture?.(event.pointerId);
    if (drag.kind === "select" || drag.kind === "fill") {
      const finalCell = resolveDragCell(engine, sheet, localPointOf(event), drag);
      if (finalCell) onPivotContextHit?.(onPivotResolve(sheet, finalCell.row, finalCell.column));
    }
    stopAutoScroll();
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
      const target: RangeRef = { sheetId, startRow: Math.min(drag.startRow, drag.currentRow), endRow: Math.max(drag.anchorRow, drag.currentRow), startColumn: Math.min(drag.startColumn, drag.currentColumn), endColumn: Math.max(drag.anchorColumn, drag.currentColumn) };
      const partialMerge = sheet.merges.some((merge) => intersectsRange(target, merge.range) && !containsRange(target, merge.range));
      if (!partialMerge && (target.endRow !== drag.anchorRow || target.endColumn !== drag.anchorColumn)) {
        onFillRange(target);
        onSelectionChange({ ranges: [target], primaryRangeIndex: 0, activeCell: { row: drag.currentRow, column: drag.currentColumn }, anchorCell: { row: drag.startRow, column: drag.startColumn } });
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
      const range = expandRangeForMerges(sheet, baseRange);
      const nextSelection: SelectionState = drag.additive
        ? { ...selection, ranges: [...selection.ranges, range], primaryRangeIndex: selection.ranges.length, activeCell: { row: drag.currentRow, column: drag.currentColumn }, anchorCell: { row: drag.anchorRow, column: drag.anchorColumn } }
        : { ranges: [range], primaryRangeIndex: 0, activeCell: { row: drag.currentRow, column: drag.currentColumn }, anchorCell: { row: drag.anchorRow, column: drag.anchorColumn } };
      clearTransientSelection();
      onSelectionChange(nextSelection);
    }
  }, [clearTransientSelection, engineRef, localPointOf, onExtendSelection, onFillRange, onPivotContextHit, onPivotResolve, onResizeColumn, onResizeRow, onSelectionChange, selection, setFillPreview, sheet, sheetId, skeleton, stopAutoScroll, zoom]);

  const handleDoubleClick = useCallback((event: React.PointerEvent | React.MouseEvent) => {
    const engine = engineRef.current;
    if (!engine) return;
    const local = localPointOf(event);
    onPivotContextHit?.(null);
    const headerHit = engine.headerHitAtLocal(local);
    if (headerHit?.resizeBoundaryPx !== undefined) {
      const column = headerHit.index;
      let maxWidth = 60;
      const context = engine.getCanvas("content")?.getContext("2d");
      if (context) {
        context.font = "13px Segoe UI, sans-serif";
        for (let row = 0; row < Math.min(sheet.rowCount, 200); row += 1) {
          const cell = sheet.getCell(row, column);
          if (cell?.value) maxWidth = Math.max(maxWidth, context.measureText(cell.value).width + 16);
        }
      }
      onResizeColumn(column, Math.round(maxWidth / (zoom / 100)));
      return;
    }
    const hitCell = engine.cellAtLocalPoint(local);
    if (!hitCell) return;
    const pivotContextHit = onPivotResolve(sheet, hitCell.row, hitCell.column);
    if (pivotContextHit) {
      onPivotContextHit?.(pivotContextHit);
      const pivotTarget = findPivotProjectionCell(sheet, hitCell.row, hitCell.column);
      if (pivotTarget && isPivotValueCell(pivotTarget.cell) && pivotTarget.cell.sourceRowPaths && pivotTarget.cell.sourceRowPaths.length > 0) {
        onPivotShowDetails?.({ pivotId: pivotTarget.projection.pivotId, sourceRowPaths: pivotTarget.cell.sourceRowPaths, hit: pivotContextHit });
      }
      return;
    }
    const cell = resolveMergedCell(sheet, hitCell);
    const validationList = getValidationList(cell.row, cell.column);
    if (validationList && validationList.length > 0) {
      setValidationDropdown({ row: cell.row, column: cell.column, options: validationList });
      return;
    }
    onBeginEdit();
  }, [engineRef, findPivotProjectionCell, getValidationList, isPivotValueCell, localPointOf, onBeginEdit, onPivotContextHit, onPivotResolve, onPivotShowDetails, onResizeColumn, setValidationDropdown, sheet, zoom]);

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
    if (key === "Enter") { event.preventDefault(); if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit); else onBeginEdit(); return; }
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
      else if (editingCell || editingActiveRef.current) onAppendFormulaDraft?.(key);
      else { editingActiveRef.current = true; onBeginEdit(key); }
    }
  }, [canRepeat, containerRef, contextRangeRef, drawingSelectionMode, editingCell, formatPainterActive, formulaDraft, onAppendFormulaDraft, onBeginEdit, onCancelEdit, onCancelFormatPainter, onCommitEdit, onExitDrawingSelectionMode, onFormulaDraftChange, onJumpEdge, onMovePrimary, onPivotContextHit, onPivotResolve, onShortcut, onToggleAbsolute, phase, selectedFloatingId, selection, setContextHit, setContextMenu, sheet, sheetId, skeleton]);

  return {
    clearTransientSelection,
    dragRef,
    handleDoubleClick,
    handleKeyDown,
    handlePointerDown,
    handlePointerMove,
    handlePointerUp,
    handleWheel,
    localPointOf,
  };
}
