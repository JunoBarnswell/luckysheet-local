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
  type Rect,
  createEmptyChromeState,
} from "@react-sheets/render-engine";
import type {
  DrawingObject,
  DrawingPayload,
  PivotGridProjection,
  PivotHitTest,
  PivotProjectionCell,
  PivotSourceRowPath,
  PivotResultTree,
  RangeRef,
  SparklineModel,
} from "@react-sheets/core-model";
import { CellEditor } from "./CellEditor";
import { FilterPopover } from "./FilterPopover";
import { resolveContextHit, type PeerCursor, type ResolvedContextHit, type SelectionState, type CanvasSheetSnapshot, type AppPhase } from "@react-sheets/spreadsheet-app";
import type { CanvasCellSnapshot } from "@react-sheets/spreadsheet-app";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { createCanvasFloatingDrawables } from "./canvas/drawing-renderers";
import { useCanvasInteraction } from "./canvas/useCanvasInteraction";
import type { ColumnDimensionController } from '../editor/column-dimension-controller';

export interface SheetCanvasProps {
  sheet: CanvasSheetSnapshot;
  sheetId: string;
  selection: SelectionState;
  activeCell: string;
  formulaDraft: string;
  editingCell: { row: number; column: number } | null;
  phase: AppPhase;
  zoom: number;
  peers: PeerCursor[];
  drawings?: readonly DrawingObject[];
  drawingPayloads?: ReadonlyMap<string, DrawingPayload>;
  allSheets?: readonly CanvasSheetSnapshot[];
  pivotResults?: Record<string, PivotResultTree>;
  sparklines?: SparklineModel[];
  selectedFloatingId: string | null;
  showFormulas?: boolean;
  /** Notifies the host when a visible Pivot projection becomes/leaves the active context. */
  onPivotContextHit?: (hit: ResolvedContextHit | null) => void;
  /** Lets the host add/replace Pivot-specific right-click commands. */
  getPivotContextMenuItems?: (hit: ResolvedContextHit) => readonly ContextMenuItem[];
  /** Opens a real details-sheet flow for a Pivot value/double-click or menu action. */
  onPivotShowDetails?: (request: PivotShowDetailsRequest) => void;
  onSelectionChange: (selection: SelectionState) => void;
  onMovePrimary: (rowDelta: number, columnDelta: number, opts?: { extend?: boolean }) => void;
  onCommitCell: (value: string) => void;
  onBeginEdit: (initialText?: string) => void;
  onCancelEdit: () => void;
  onCommitEdit: (moveAfter?: "down" | "up" | "left" | "right" | "none") => void;
  onFormulaDraftChange: (value: string) => void;
  onAppendFormulaDraft?: (fragment: string) => void;
  onInsertRef: (refText: string) => void;
  onToggleAbsolute: () => void;
  onJumpEdge: (direction: "up" | "down" | "left" | "right", extend?: boolean) => void;
  onSelectAll: () => void;
  onExtendSelection?: (row: number, column: number) => void;
  onResizeRow: (row: number, heightPx: number) => void;
  columnDimensions: ColumnDimensionController;
  onOpenColumnWidthDialog: (columns: number[]) => void;
  onFillRange: (target: { startRow: number; endRow: number; startColumn: number; endColumn: number }) => void;
  drawingSelectionMode?: boolean;
  onExitDrawingSelectionMode?: () => void;
  onFloatingSelect: (hit: FloatingHit | null, mode?: 'replace' | 'add' | 'toggle') => void;
  onFloatingMove: (drawingId: string, bounds: Rect, rotation?: number) => void;
  onFloatingRemove: (drawingId: string) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClearSelection?: (mode: "contents" | "formats") => void;
  /** Format painter is a transient session interaction, never canvas-local state. */
  formatPainterActive?: boolean;
  onCancelFormatPainter?: () => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onUndo: () => void;
  onRedo: () => void;
  /** Host owns command/session execution after the shared registry resolves a shortcut. */
  onShortcut?: (id: string) => boolean;
  canRepeat?: boolean;
  onOpenInspector: () => void;
  onApplyFilter: (column: number, patch: { selectedValues?: string[] | null }) => void;
  onToggleOutline?: (groupId: string) => void;
  getValidationList: (row: number, column: number) => string[] | undefined;
  onRetry: () => void;
  onCreateSheet: () => void;
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

function toChromeSelection(selection: SelectionState): ChromeState['selection'] {
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

/**
 * A projection stores cell coordinates relative to its target anchor. The
 * canvas, on the other hand, always asks for worksheet coordinates. Keeping
 * this conversion at the UI boundary prevents the derived projection from
 * leaking into ordinary worksheet cells or selection state.
 */
export function findPivotProjectionCell(
  sheet: CanvasSheetSnapshot,
  row: number,
  column: number,
): { projection: PivotGridProjection; cell: PivotProjectionCell } | null {
  for (const projection of Object.values(sheet.pivotProjections)) {
    if (projection.sheetId !== sheet.id || projection.collision.status !== "clear") continue;
    const relativeRow = row - projection.target.anchor.row;
    const relativeColumn = column - projection.target.anchor.column;
    const cell = projection.cells.find((candidate) => candidate.row === relativeRow && candidate.column === relativeColumn);
    if (cell) return { projection, cell };
  }
  return null;
}

function pivotHitKind(cell: PivotProjectionCell): PivotHitTest['kind'] {
  if (cell.kind === "expand-toggle") return "expand-toggle";
  if (cell.kind === "filter") return "filter";
  if (cell.kind === "title" || cell.kind === "column-header" || cell.kind === "row-header") return "header";
  return "cell";
}

/** Resolve an absolute worksheet coordinate through the canonical context resolver. */
export function resolvePivotProjectionHit(
  sheet: CanvasSheetSnapshot,
  row: number,
  column: number,
): ResolvedContextHit | null {
  const target = findPivotProjectionCell(sheet, row, column);
  if (!target) return null;
  const hit: PivotHitTest = {
    kind: pivotHitKind(target.cell),
    pivotId: target.projection.pivotId,
    cellId: target.cell.id,
    row,
    column,
    nodeId: target.cell.nodeId,
    sourceRowPaths: target.cell.sourceRowPaths,
  };
  return resolveContextHit({ sheetId: sheet.id, pivot: hit });
}

export function isPivotValueCell(cell: PivotProjectionCell): boolean {
  return cell.kind === "value" || cell.kind === "subtotal" || cell.kind === "grand-total";
}

/** Convert a derived projection cell to the render-engine cell contract. */
export function pivotProjectionCellRenderData(cell: PivotProjectionCell): CellRenderData {
  const text = cell.kind === "expand-toggle"
    ? `${cell.expanded ? "▾" : "▸"} ${cell.text}`
    : cell.text;
  const style: NonNullable<CellRenderData["style"]> = {
    background: cell.kind === "title"
      ? "#dbeafe"
      : cell.kind === "column-header" || cell.kind === "filter"
        ? "#f1f5f9"
        : cell.kind === "grand-total"
          ? "#eff6ff"
          : cell.kind === "subtotal"
            ? "#f8fafc"
            : "#ffffff",
    textColor: cell.kind === "error" ? "#b91c1c" : cell.kind === "loading" ? "#92400e" : "#1e293b",
    bold: cell.kind === "title" || cell.kind === "column-header" || cell.kind === "subtotal" || cell.kind === "grand-total",
    italic: cell.kind === "filter",
    horizontalAlignment: isPivotValueCell(cell) ? "right" : "left",
    verticalAlignment: "middle",
    wrapText: cell.kind === "loading" || cell.kind === "error",
    borders: {
      bottom: { color: "#cbd5e1", style: cell.kind === "grand-total" ? "double" : "thin" },
    },
  };
  return {
    value: cell.value,
    displayValue: text,
    style,
    error: cell.kind === "error" ? text : undefined,
    invalid: cell.kind === "error",
  };
}

export interface PivotShowDetailsRequest {
  pivotId: string;
  sourceRowPaths: readonly PivotSourceRowPath[];
  hit: ResolvedContextHit;
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
  drawings = sheet.drawings,
  drawingPayloads = sheet.drawingPayloads,
  allSheets = [],
  pivotResults = {},
  sparklines = [],
  selectedFloatingId,
  showFormulas = false,
  onPivotContextHit,
  getPivotContextMenuItems,
  onPivotShowDetails,
  onSelectionChange,
  onMovePrimary,
  onCommitCell,
  onBeginEdit,
  onCancelEdit,
  onCommitEdit,
  onFormulaDraftChange,
  onAppendFormulaDraft,
  onInsertRef,
  onToggleAbsolute,
  onJumpEdge,
  onSelectAll,
  onExtendSelection,
  onResizeRow,
  columnDimensions,
  onOpenColumnWidthDialog,
  onFillRange,
  drawingSelectionMode = false,
  onExitDrawingSelectionMode,
  onFloatingSelect,
  onFloatingMove,
  onFloatingRemove,
  onCommand,
  onClearSelection,
  formatPainterActive = false,
  onCancelFormatPainter,
  onCopy,
  onCut,
  onPaste,
  onUndo,
  onRedo,
  onShortcut,
  canRepeat = false,
  onOpenInspector,
  onApplyFilter,
  onToggleOutline,
  getValidationList,
  onRetry,
  onCreateSheet,
}: SheetCanvasProps) {
  const engineRef = useRef<CanvasRenderEngine | null>(null);
  const imageCacheRef = useRef(new Map<string, HTMLImageElement>());
  const containerRef = useRef<HTMLDivElement | null>(null);
  const contextRangeRef = useRef<RangeRef | null>(null);
  const [contextMenu, setContextMenu] = useState({ x: 0, y: 0, open: false });
  const [contextHit, setContextHit] = useState<ResolvedContextHit | null>(null);
  const [contextHeader, setContextHeader] = useState<'row' | 'column' | null>(null);
  const [filterPopover, setFilterPopover] = useState<{ column: number; x: number; y: number } | null>(null);
  const [validationDropdown, setValidationDropdown] = useState<{ row: number; column: number; options: string[] } | null>(null);
  const [fillPreview, setFillPreview] = useState<{ startRow: number; endRow: number; startColumn: number; endColumn: number } | null>(null);
  const [scrollTick, setScrollTick] = useState(0);
  const [engineReady, setEngineReady] = useState(false);

  const zoomFactor = zoom / 100;

  const skeleton = useMemo(
    () =>
      new SheetSkeleton({
        rowCount: Math.max(sheet.rowCount, 200),
        columnCount: Math.max(sheet.columnCount, 26),
        defaultRowHeight: sheet.defaultRowHeightPx,
        defaultColumnWidth: sheet.defaultColumnWidthPx,
        rowHeights: new Map(Object.entries(sheet.rowHeightsPx).map(([key, value]) => [Number(key), value])),
        columnWidths: new Map(Object.entries(sheet.columnWidthsPx).map(([key, value]) => [Number(key), value])),
        hiddenRows: new Set(sheet.hiddenRows),
        hiddenColumns: new Set(sheet.hiddenColumns ?? []),
        zoom: zoomFactor,
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet.rowCount, sheet.columnCount, sheet.defaultRowHeightPx, sheet.defaultColumnWidthPx, sheet.rowHeightsPx, sheet.columnWidthsPx, sheet.hiddenRows, sheet.hiddenColumns, zoomFactor],
  );

  const cellProvider = useCallback(({ row, column }: { row: number; column: number }): CellRenderData | undefined => {
    const pivotCell = findPivotProjectionCell(sheet, row, column);
    if (pivotCell) return pivotProjectionCellRenderData(pivotCell.cell);

    const cell = sheet.getCell(row, column);
    const merge = sheet.merges.find((span) =>
      row >= span.range.startRow && row <= span.range.endRow
      && column >= span.range.startColumn && column <= span.range.endColumn);
    // Empty cells inside a merge still need a render record so the grid layer
    // can suppress the merge's internal boundaries. Returning undefined here
    // made blank merged areas look like ordinary cells.
    if (!cell) {
      if (!merge) return undefined;
      return {
        value: undefined,
        merge: {
          startRow: merge.range.startRow,
          endRow: merge.range.endRow,
          startColumn: merge.range.startColumn,
          endColumn: merge.range.endColumn,
          isAnchor: merge.anchor.row === row && merge.anchor.column === column,
        },
      };
    }
    const isAnchor = merge ? merge.anchor.row === row && merge.anchor.column === column : true;
    return {
      value: showFormulas && cell.formula ? cell.formula : parseCellValue(cell),
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
  }, [sheet, showFormulas]);

  const pivotStatusProjections = useMemo(
    () => Object.values(sheet.pivotProjections).filter((projection) =>
      projection.collision.status === "collision"
      || projection.refresh.status === "error"
      || projection.refresh.status === "refreshing"
      || projection.refresh.status === "stale"),
    [sheet.pivotProjections],
  );

  // ---------- 浮动对象绘制器 ----------

  const floatables = useMemo<FloatingDrawable[]>(() => createCanvasFloatingDrawables({
    allSheets,
    drawingPayloads,
    drawings,
    imageCache: imageCacheRef.current,
    pivotResults,
    requestRender: () => engineRef.current?.requestRender(),
    sheet,
    skeleton,
    sparklines,
  }), [allSheets, drawingPayloads, drawings, pivotResults, sparklines, skeleton, sheet]);

  // ---------- 引擎生命周期与 chrome 同步 ----------

  const chromeState = useMemo<ChromeState>(() => {
    const state = createEmptyChromeState();
    state.selection = toChromeSelection(selection);
    state.editing = editingCell ? { row: editingCell.row, column: editingCell.column } : null;
    state.filterColumns = sheet.filterColumns;
    state.filterButtons = sheet.filterButtons;
    state.tableOutlines = sheet.sheetTables.map((table) => ({
      startRow: table.range.startRow,
      endRow: table.range.endRow,
      startColumn: table.range.startColumn,
      endColumn: table.range.endColumn,
    }));
    state.outlineControls = sheet.outlineControls;
    state.remoteCursors = peers.map((peer) => ({
      actorId: peer.actorId,
      color: peer.color,
      name: peer.name,
      row: peer.row,
      column: peer.column,
    }));
    state.selectedFloatingId = selectedFloatingId;
    return state;
  }, [editingCell, peers, selectedFloatingId, selection, sheet.filterButtons, sheet.filterColumns, sheet.outlineControls, sheet.sheetTables]);

  const canvasInteraction = useCanvasInteraction({
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
    onCommitCell,
    onCommitEdit,
    onExitDrawingSelectionMode,
    onExtendSelection,
    onFillRange,
    onFloatingMove,
    onFloatingSelect,
    onFormulaDraftChange,
    onJumpEdge,
    onMovePrimary,
    onPivotContextHit,
    onPivotResolve: resolvePivotProjectionHit,
    onPivotShowDetails,
    onResizeColumn={(column, widthPx) => columnDimensions.resizeBoundary(column, widthPx)}
    onAutoFitColumn={(column) => columnDimensions.autoFit(columnDimensions.columnsForBoundary(column))}
    onAutoFitRow={(row) => columnDimensions.autoFitRows([row])}
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
  });

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
      sheet.pane.kind === 'frozen' && (sheet.pane.xSplit > 0 || sheet.pane.ySplit > 0)
        ? { xSplit: sheet.pane.xSplit, ySplit: sheet.pane.ySplit }
        : null,
    );
  }, [sheet.pane]);

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

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    const detach = engine.onViewportChanged(() => setScrollTick((tick) => tick + 1));
    return detach;
  }, []);

  // 选区变化 → 滚动至可见
  useEffect(() => {
    const engine = engineRef.current;
    if (!engine || canvasInteraction.dragRef.current) return;
    engine.ensureVisible(selection.activeCell);
  }, [selection.activeCell]);

  // Pointer, keyboard, drag-selection, and auto-scroll are implemented by useCanvasInteraction.
  // ---------- 右键菜单 ----------

  const contextMenuItems = useMemo<ContextMenuItem[]>(() => {
    if (contextHit?.kind === "pivot" && contextHit.pivot) {
      const supplied = getPivotContextMenuItems?.(contextHit);
      if (supplied) return [...supplied];
      const sourceRowPaths = contextHit.pivot.sourceRowPaths ?? [];
      return [
        {
          id: "pivot-show-details",
          label: "Show Details",
          disabled: sourceRowPaths.length === 0 || !onPivotShowDetails,
          onSelect: () => onPivotShowDetails?.({
            pivotId: contextHit.pivot!.pivotId,
            sourceRowPaths,
            hit: contextHit,
          }),
        },
        {
          id: "pivot-refresh",
          label: "Refresh PivotTable",
          onSelect: () => onCommand({ commandId: "pivot.refresh", params: { sheetId, pivotId: contextHit.pivot!.pivotId } }),
        },
      ];
    }
    const getContextRange = () => contextRangeRef.current ?? selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
    if (contextHeader === 'column') {
      const columns = columnDimensions.selectedColumns();
      return [
        { id: 'column-width', label: 'Column Width…', onSelect: () => onOpenColumnWidthDialog(columns) },
        { id: 'column-autofit', label: 'AutoFit Column Width', onSelect: () => { void columnDimensions.autoFit(columns); } },
        { id: 'column-hide', label: 'Hide Columns', onSelect: () => columnDimensions.setHidden(columns, true) },
        { id: 'column-unhide', label: 'Unhide Columns', onSelect: () => columnDimensions.setHidden(columns, false) },
      ];
    }
    const items: ContextMenuItem[] = [
      { id: "cut", label: "Cut", shortcut: "Ctrl+X", onSelect: onCut },
      { id: "copy", label: "Copy", shortcut: "Ctrl+C", onSelect: onCopy },
      { id: "paste", label: "Paste", shortcut: "Ctrl+V", onSelect: onPaste },
      { id: "sep-1", label: "", separator: true },
      { id: "insert-row", label: "Insert row above", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.rows.insert", params: { sheetId, at: range.startRow, count: 1 } }); } },
      { id: "insert-column", label: "Insert column left", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.columns.insert", params: { sheetId, at: range.startColumn, count: 1 } }); } },
      { id: "delete-row", label: "Delete row", danger: true, onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.rows.delete", params: { sheetId, at: range.startRow, count: range.endRow - range.startRow + 1 } }); } },
      { id: "delete-column", label: "Delete column", danger: true, onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.columns.delete", params: { sheetId, at: range.startColumn, count: range.endColumn - range.startColumn + 1 } }); } },
      { id: "sep-2", label: "", separator: true },
      { id: "hide-row", label: "Hide rows", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.row.hide", params: { sheetId, index: range.startRow } }); } },
      { id: "hide-col", label: "Hide columns", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.column.hide", params: { sheetId, index: range.startColumn } }); } },
      { id: "unhide-all", label: "Unhide all", onSelect: () => onCommand({ commandId: "sheet.rows.unhide.all", params: { sheetId } }) },
      { id: "sep-3", label: "", separator: true },
      { id: "clear", label: "Clear contents", onSelect: () => onClearSelection?.("contents") },
      { id: "clear-formats", label: "Clear formats", onSelect: () => onClearSelection?.("formats") },
      { id: "comment-add", label: "Add comment", onSelect: onOpenInspector },
    ];
    return items;
  }, [columnDimensions, contextHeader, contextHit, getPivotContextMenuItems, onClearSelection, onCommand, onCopy, onCut, onOpenColumnWidthDialog, onOpenInspector, onPaste, onPivotShowDetails, selection, sheetId]);

  const handleContextMenu = useCallback((event: React.MouseEvent) => {
    event.preventDefault();
    setFillPreview(null);
    onPivotContextHit?.(null);
    const engine = engineRef.current;
    if (!engine || phase !== "ready") {
      setContextHit(null);
      setContextMenu({ x: event.clientX, y: event.clientY, open: true });
      return;
    }
    const local = canvasInteraction.localPointOf(event);
    const headerHit = engine.headerHitAtLocal(local);
    setContextHit(null);
    setContextHeader(headerHit?.kind === 'col' ? 'column' : headerHit?.kind === 'row' ? 'row' : null);
    contextRangeRef.current = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? null;
    if (headerHit?.kind === "corner") {
      contextRangeRef.current = { sheetId, startRow: 0, endRow: Math.max(0, skeleton.rowCount - 1), startColumn: 0, endColumn: Math.max(0, skeleton.columnCount - 1) };
      onSelectAll();
    } else if (headerHit?.kind === "row") {
      const row = headerHit.index;
      contextRangeRef.current = { sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: Math.max(0, skeleton.columnCount - 1) };
      onSelectionChange({
        ranges: [{ sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: Math.max(0, skeleton.columnCount - 1) }],
        primaryRangeIndex: 0,
        activeCell: { row, column: 0 },
        anchorCell: { row, column: 0 },
      });
    } else if (headerHit?.kind === "col") {
      const column = headerHit.index;
      const alreadySelected = selection.ranges.some((range) => range.startRow === 0 && range.endRow >= sheet.rowCount - 1 && column >= range.startColumn && column <= range.endColumn);
      if (!alreadySelected) {
        contextRangeRef.current = { sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: column, endColumn: column };
        onSelectionChange({
          ranges: [contextRangeRef.current],
          primaryRangeIndex: 0,
          activeCell: { row: 0, column },
          anchorCell: { row: 0, column },
        });
      }
    } else {
      const hitCell = engine.cellAtLocalPoint(local);
      if (hitCell) {
        const pivotContextHit = resolvePivotProjectionHit(sheet, hitCell.row, hitCell.column);
        if (pivotContextHit) {
          setContextHit(pivotContextHit);
          onPivotContextHit?.(pivotContextHit);
          const targetRange: RangeRef = {
            sheetId,
            startRow: hitCell.row,
            endRow: hitCell.row,
            startColumn: hitCell.column,
            endColumn: hitCell.column,
          };
          const alreadySelected = selection.ranges.some((range) => intersectsRange(range, targetRange));
          if (!alreadySelected) {
            contextRangeRef.current = targetRange;
            onSelectionChange({
              ranges: [targetRange],
              primaryRangeIndex: 0,
              activeCell: { row: hitCell.row, column: hitCell.column },
              anchorCell: { row: hitCell.row, column: hitCell.column },
            });
          } else {
            contextRangeRef.current = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? targetRange;
          }
        } else {
          setContextHit(null);
          const cell = resolveMergedCell(sheet, hitCell);
          const merge = mergeAtCell(sheet, hitCell);
          const targetRange: RangeRef = merge?.range ?? {
            sheetId,
            startRow: cell.row,
            endRow: cell.row,
            startColumn: cell.column,
            endColumn: cell.column,
          };
          const alreadySelected = selection.ranges.some((range) => intersectsRange(range, targetRange));
          if (!alreadySelected) {
            contextRangeRef.current = expandRangeForMerges(sheet, targetRange);
            onSelectionChange({
              ranges: [contextRangeRef.current],
              primaryRangeIndex: 0,
              activeCell: cell,
              anchorCell: cell,
            });
          } else {
            contextRangeRef.current = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? targetRange;
          }
        }
      }
    }
    setContextMenu({ x: event.clientX, y: event.clientY, open: true });
  }, [canvasInteraction.localPointOf, onPivotContextHit, onSelectAll, onSelectionChange, phase, selection, sheet, sheetId, skeleton]);

  // ---------- 编辑器定位(随滚动更新) ----------

  const editingPivotContextHit = editingCell
    ? resolvePivotProjectionHit(sheet, editingCell.row, editingCell.column)
    : null;

  // 编辑器随滚动重定位:依赖 scrollTick 触发重算
  const editorRect = useMemo(() => {
    void scrollTick;
    const engine = engineRef.current;
    if (!engine || !editingCell || editingPivotContextHit) return null;
    // The engine owns the render geometry. Reading the local React skeleton
    // here can race with setSkeleton during a session refresh, and the main
    // pane is not the correct origin for frozen rows/columns.
    const rect = engine.skeleton.getCellRect(editingCell.row, editingCell.column);
    if (!rect) return null;
    const topLeft = engine.contentToMainScreen({ x: rect.x, y: rect.y }, editingCell);
    return {
      x: topLeft.x,
      y: topLeft.y,
      width: rect.width,
      height: rect.height,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editingCell, editingPivotContextHit, skeleton, scrollTick]);

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
            {sheet.pane.kind !== 'none' ? (
              <Text size="xs" tone="subtle">{sheet.pane.kind} {sheet.pane.xSplit}x{sheet.pane.ySplit}</Text>
            ) : null}
          </Inline>
          <Text size="xs" tone="subtle">{activeCell}</Text>
        </Inline>
        <Box className="relative min-h-0 flex-1">
          <Box
            ref={containerRef}
            role="grid"
            aria-label="Spreadsheet canvas"
            data-testid="sheet-canvas"
            aria-rowcount={sheet.rowCount}
            aria-colcount={sheet.columnCount}
            aria-rowindex={selection.activeCell.row + 1}
            aria-colindex={selection.activeCell.column + 1}
            tabIndex={0}
            className="absolute inset-0 outline-none"
            onPointerDown={canvasInteraction.handlePointerDown}
            onPointerMove={canvasInteraction.handlePointerMove}
            onPointerUp={canvasInteraction.handlePointerUp}
            onPointerCancel={canvasInteraction.handlePointerUp}
            onDoubleClick={canvasInteraction.handleDoubleClick}
            onWheel={canvasInteraction.handleWheel}
            onKeyDown={canvasInteraction.handleKeyDown}
            onContextMenu={handleContextMenu}
          >
            <CanvasRenderSurface
              onReady={(engine) => {
                engineRef.current = engine;
                setEngineReady(true);
                engine.setCellProvider(cellProvider);
                engine.setSkeleton(skeleton);
                engine.setFloating(floatables, selectedFloatingId);
                engine.setChrome(chromeState);
              }}
              className="absolute inset-0"
            />
            {engineReady ? pivotStatusProjections.map((projection) => (
              <PivotProjectionStatusNotice
                key={`${projection.pivotId}:${projection.refresh.status}:${projection.refresh.error ?? ""}`}
                engine={engineRef.current}
                projection={projection}
                scrollTick={scrollTick}
              />
            )) : null}
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
            <FillPreviewOverlay engine={engineRef.current} preview={fillPreview} />
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

function parseCellValue(cell: CanvasCellSnapshot): string | number | boolean | null {
  const numeric = Number(cell.value.replace(/[$,]/g, ""));
  if (cell.value !== "" && Number.isFinite(numeric) && /\d/.test(cell.value)) return numeric;
  if (cell.value === "TRUE") return true;
  if (cell.value === "FALSE") return false;
  return cell.value;
}

function pivotProjectionStatusMessage(projection: PivotGridProjection): string | null {
  if (projection.collision.status === "collision") {
    const reason = projection.collision.reasons.length > 0
      ? projection.collision.reasons.join(", ")
      : "target range is occupied";
    return `PivotTable ${projection.pivotId} cannot expand: ${reason}`;
  }
  if (projection.refresh.status === "error") {
    return `PivotTable ${projection.pivotId} error: ${projection.refresh.error ?? "refresh failed"}`;
  }
  if (projection.refresh.status === "refreshing") return `PivotTable ${projection.pivotId}: Loading PivotTable…`;
  if (projection.refresh.status === "stale") return `PivotTable ${projection.pivotId}: Refresh required`;
  return null;
}

function PivotProjectionStatusNotice({
  engine,
  projection,
  scrollTick,
}: {
  engine: CanvasRenderEngine | null;
  projection: PivotGridProjection;
  scrollTick: number;
}): React.ReactElement | null {
  void scrollTick;
  const message = pivotProjectionStatusMessage(projection);
  if (!engine || !message) return null;
  const rect = engine.contentRangeToScreenRects(projection.occupiedRange)[0];
  if (!rect) return null;
  const collision = projection.collision.status === "collision";
  return (
    <Box
      role="status"
      aria-live="polite"
      className={`pointer-events-none absolute z-10 max-w-[min(32rem,calc(100%-1rem))] rounded border px-2 py-1 text-[11px] shadow-sm ${collision ? "border-amber-300 bg-amber-50 text-amber-800" : "border-red-200 bg-red-50 text-red-700"}`}
      style={{ left: rect.x + 2, top: rect.y + 2 }}
    >
      {message}
    </Box>
  );
}

function FillPreviewOverlay({
  engine,
  preview,
}: {
  engine: CanvasRenderEngine | null;
  preview: { startRow: number; endRow: number; startColumn: number; endColumn: number };
}): React.ReactElement | null {
  if (!engine) return null;
  const rects = engine.contentRangeToScreenRects(preview);
  if (rects.length === 0) return null;
  return (
    <>
      {rects.map((screen, index) => (
        <Box
          key={`${screen.x}:${screen.y}:${index}`}
          className="pointer-events-none absolute z-10 border-2 border-dashed border-blue-500 bg-blue-500/5"
          style={{ left: screen.x, top: screen.y, width: screen.width, height: screen.height }}
        />
      ))}
    </>
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
