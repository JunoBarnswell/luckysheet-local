import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Box,
  ContextMenu,
  type ContextMenuItem,
  Panel,
  Stack,
  StatePanel,
  Button,
  ScrollBar,
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
  PivotFilter,
  PivotHitTest,
  PivotProjectionCell,
  PivotSourceRowPath,
  PivotResultTree,
  PivotSort,
  RangeRef,
  SparklineModel,
  WorkbookTableModel,
} from "@react-sheets/core-model";
import { CellEditor } from "./CellEditor";
import { FilterPopover, type FilterPatch } from "./FilterPopover";
import { resolveContextHit, type PeerCursor, type ResolvedContextHit, type SelectionState, type CanvasSheetSnapshot, type AppPhase } from "@react-sheets/spreadsheet-app";
import type { CanvasCellSnapshot } from "@react-sheets/spreadsheet-app";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { createCanvasFloatingDrawables } from "./canvas/drawing-renderers";
import { useCanvasInteraction } from "./canvas/useCanvasInteraction";
import type { ColumnDimensionController } from '../editor/column-dimension-controller';
import type { Locale } from '../i18n';
import { pivotTemplate, pivotText } from './pivot/pivot-localization';
import { PivotHeaderFilterPopover } from './pivot/PivotHeaderFilterPopover';
import { createMergeSpatialIndex } from './canvas/merge-spatial-index';
import { GanttViewOverlay } from './GanttViewOverlay';
import { ReportViewOverlay } from './ReportViewOverlay';

export interface SheetCanvasProps {
  locale: Locale;
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
  tables?: readonly WorkbookTableModel[];
  selectedFloatingId: string | null;
  showFormulas?: boolean;
  /** Notifies the host when a visible Pivot projection becomes/leaves the active context. */
  onPivotContextHit?: (hit: ResolvedContextHit | null) => void;
  /** Lets the host add/replace Pivot-specific right-click commands. */
  getPivotContextMenuItems?: (hit: ResolvedContextHit) => readonly ContextMenuItem[];
  /** Opens a real details-sheet flow for a Pivot value/double-click or menu action. */
  onPivotShowDetails: (request: PivotShowDetailsRequest) => void;
  onApplyPivotFilter: (pivotId: string, fieldId: string, filter: PivotFilter | undefined, sort: PivotSort | undefined, scope: 'report' | 'field') => void;
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
  onClearSelection: (mode: "contents" | "formats") => void;
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
  onApplyFilter: (column: number, patch: FilterPatch) => void;
  onSortFilterColumn: (column: number, ascending: boolean) => void;
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
export function pivotProjectionCellRenderData(cell: PivotProjectionCell, locale: Locale = 'en-US'): CellRenderData {
  const localizedText = cell.filterSummary
    ? `${cell.filterSummary.fieldName}: ${cell.filterSummary.mode === 'all' ? pivotText(locale, 'allItems') : pivotTemplate(locale, 'selectedItems', { count: cell.filterSummary.count })}`
    : cell.captionKey === 'row-labels'
    ? pivotText(locale, 'rowLabels')
    : cell.captionKey === 'grand-total'
      ? pivotText(locale, 'grandTotal')
      : cell.captionKey === 'loading'
        ? pivotText(locale, 'loadingPivot')
        : cell.text;
  const text = cell.kind === "expand-toggle"
    ? `${cell.expanded ? "▾" : "▸"} ${localizedText}`
    : localizedText;
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
  locale,
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
  tables = [],
  selectedFloatingId,
  showFormulas = false,
  onPivotContextHit,
  getPivotContextMenuItems,
  onPivotShowDetails,
  onApplyPivotFilter,
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
  onSortFilterColumn,
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
  const [filterPopover, setFilterPopover] = useState<{ column: number; x: number; y: number } | null>(null);
  const [pivotFilterPopover, setPivotFilterPopover] = useState<{ pivotId: string; fieldId: string; scope: 'report' | 'field'; x: number; y: number } | null>(null);
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

  const findMerge = useMemo(() => createMergeSpatialIndex(sheet.merges), [sheet.merges]);

  const cellProvider = useCallback(({ row, column }: { row: number; column: number }): CellRenderData | undefined => {
    const pivotCell = findPivotProjectionCell(sheet, row, column);
    if (pivotCell) return pivotProjectionCellRenderData(pivotCell.cell, locale);

    const cell = sheet.getCell(row, column);
    const merge = findMerge(row, column);
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
      editor: cell.editor,
      presentation: cell.presentation,
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
  }, [findMerge, locale, sheet, showFormulas]);

  const pivotStatusProjections = useMemo(
    () => Object.values(sheet.pivotProjections).filter((projection) =>
      projection.collision.status === "collision"
      || projection.refresh.status === "error"
      || projection.refresh.status === "refreshing"
      || projection.refresh.status === "stale"),
    [sheet.pivotProjections],
  );
  const pivotHeaderFilterCells = useMemo(() => Object.values(sheet.pivotProjections).flatMap((projection) => {
    if (projection.collision.status !== 'clear') return [];
    return projection.cells
      .filter((cell) => (cell.kind === 'column-header' || cell.kind === 'filter') && Boolean(cell.fieldId))
      .map((cell) => ({ projection, cell }));
  }), [sheet.pivotProjections]);

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
    tables,
  }), [allSheets, drawingPayloads, drawings, pivotResults, sparklines, skeleton, sheet, tables]);

  // ---------- 引擎生命周期与 chrome 同步 ----------

  const chromeState = useMemo<ChromeState>(() => {
    const state = createEmptyChromeState();
    state.selection = toChromeSelection(selection);
    state.editing = editingCell ? { row: editingCell.row, column: editingCell.column } : null;
    state.filterColumns = sheet.activeFilterColumns;
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
  }, [editingCell, peers, selectedFloatingId, selection, sheet.activeFilterColumns, sheet.filterButtons, sheet.outlineControls, sheet.sheetTables]);

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
    onResizeColumn: (column, widthPx) => columnDimensions.resizeBoundary(column, widthPx),
    onAutoFitColumn: (column) => columnDimensions.autoFit(columnDimensions.columnsForBoundary(column)),
    onAutoFitRow: (row) => columnDimensions.autoFitRows([row]),
    formatColumnWidthPreview: (widthPx) => columnDimensions.previewPixels(widthPx),
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
    engine.setPane(sheet.pane.kind === 'none' ? null : sheet.pane);
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
          disabled: sourceRowPaths.length === 0,
          onSelect: () => onPivotShowDetails({
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
    if (contextHit?.kind === 'column-header') {
      const columns = columnDimensions.selectedColumns();
      return [
        { id: 'column-width', label: 'Column Width…', onSelect: () => onOpenColumnWidthDialog(columns) },
        { id: 'column-autofit', label: 'AutoFit Column Width', onSelect: () => { void columnDimensions.autoFit(columns); } },
        { id: 'column-hide', label: 'Hide Columns', onSelect: () => columnDimensions.setHidden(columns, true) },
        { id: 'column-unhide', label: 'Unhide Columns', onSelect: () => columnDimensions.setHidden(columns, false) },
      ];
    }
    if (contextHit?.kind === 'row-header') {
      const range = getContextRange();
      const startRow = range?.startRow ?? selection.activeCell.row;
      const rowCount = range ? range.endRow - range.startRow + 1 : 1;
      return [
        { id: 'row-insert', label: 'Insert rows above', onSelect: () => onCommand({ commandId: 'sheet.rows.insert', params: { sheetId, at: startRow, count: rowCount } }) },
        { id: 'row-delete', label: 'Delete rows', danger: true, onSelect: () => onCommand({ commandId: 'sheet.rows.delete', params: { sheetId, at: startRow, count: rowCount } }) },
        { id: 'row-hide', label: 'Hide rows', onSelect: () => onCommand({ commandId: 'sheet.row.hide', params: { sheetId, index: startRow } }) },
        { id: 'row-unhide', label: 'Unhide rows', onSelect: () => onCommand({ commandId: 'sheet.rows.unhide.all', params: { sheetId } }) },
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
      { id: "clear", label: "Clear contents", onSelect: () => onClearSelection("contents") },
      { id: "clear-formats", label: "Clear formats", onSelect: () => onClearSelection("formats") },
      { id: "comment-add", label: "Add comment", onSelect: onOpenInspector },
    ];
    return items;
  }, [columnDimensions, contextHit, getPivotContextMenuItems, onClearSelection, onCommand, onCopy, onCut, onOpenColumnWidthDialog, onOpenInspector, onPaste, onPivotShowDetails, selection, sheetId]);

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
    setContextHit(headerHit?.kind === 'row' || headerHit?.kind === 'col'
      ? resolveContextHit({
        sheetId,
        header: headerHit.kind === 'row' ? 'row' : 'column',
        cell: { row: headerHit.kind === 'row' ? headerHit.index : 0, column: headerHit.kind === 'col' ? headerHit.index : 0 },
      })
      : null);
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
    const topLeft = engine.contentToScreen({ x: rect.x, y: rect.y }, editingCell);
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
            {engineReady && engineRef.current ? (() => {
              const viewport = engineRef.current.viewport.getSnapshot();
              const content = engineRef.current.skeleton.contentSize;
              return (
                <>
                  <ScrollBar
                    contentSize={content.width}
                    offset={viewport.scrollX}
                    onChange={(offset) => engineRef.current?.scrollTo(offset, viewport.scrollY)}
                    orientation="horizontal"
                    viewportSize={viewport.width}
                  />
                  <ScrollBar
                    contentSize={content.height}
                    offset={viewport.scrollY}
                    onChange={(offset) => engineRef.current?.scrollTo(viewport.scrollX, offset)}
                    orientation="vertical"
                    viewportSize={viewport.height}
                  />
                </>
              );
            })() : null}
            {engineReady ? pivotStatusProjections.map((projection) => (
              <PivotProjectionStatusNotice
                key={`${projection.pivotId}:${projection.refresh.status}:${projection.refresh.error ?? ""}`}
                engine={engineRef.current}
                projection={projection}
                locale={locale}
                scrollTick={scrollTick}
              />
            )) : null}
            {engineReady && sheet.kind === 'gantt-sheet' && sheet.ganttSheet ? (
              <GanttViewOverlay engine={engineRef.current} sheet={sheet} tables={tables ?? []} scrollTick={scrollTick} />
            ) : null}
            {engineReady && sheet.kind === 'report-sheet' && sheet.reportSheet ? (
              <ReportViewOverlay engine={engineRef.current} sheet={sheet} tables={tables ?? []} sourceSheets={allSheets ?? []} scrollTick={scrollTick} />
            ) : null}
            {engineReady ? pivotHeaderFilterCells.map(({ cell, projection }) => {
              const row = projection.target.anchor.row + cell.row;
              const column = projection.target.anchor.column + cell.column;
              const rect = engineRef.current?.contentRangeToScreenRects({ startRow: row, endRow: row, startColumn: column, endColumn: column })[0];
              if (!rect || !cell.fieldId) return null;
              const fieldName = sheet.pivots.find((pivot) => pivot.id === projection.pivotId)?.fieldCatalog.fields.find((field) => field.fieldId === cell.fieldId)?.name ?? cell.fieldId;
              return <Button key={`${projection.pivotId}:${cell.id}:filter`} aria-label={`${pivotText(locale, 'filterValues')}: ${fieldName}`} icon="chevron-down" iconOnly size="xs" variant="ghost" className="absolute z-20 !h-4 !min-h-0 !w-4 rounded-none border border-[#9ba8b6] bg-white p-0 text-[#50606e]" style={{ left: rect.x + rect.width - 18, top: rect.y + 2 }} onClick={() => setPivotFilterPopover({ pivotId: projection.pivotId, fieldId: cell.fieldId!, scope: cell.kind === 'filter' ? 'report' : 'field', x: Math.max(2, Math.min(rect.x, (containerRef.current?.clientWidth ?? 320) - 304)), y: rect.y + rect.height })} />;
            }) : null}
          </Box>

          {editorRect && editingCell ? (
            <Box
              className="absolute z-20 overflow-hidden rounded-none border border-[#5292f7] bg-white"
              style={{ left: editorRect.x, top: editorRect.y, width: editorRect.width, height: editorRect.height }}
            >
              <CellEditor
                cellStyle={sheet.getCell(editingCell.row, editingCell.column)?.style}
                initialText={formulaDraft}
                onCancel={onCancelEdit}
                onChange={onFormulaDraftChange}
                onCommit={onCommitEdit}
                onInsertRef={onInsertRef}
              />
            </Box>
          ) : null}

          {pivotFilterPopover ? (() => {
            const pivot = sheet.pivots.find((candidate) => candidate.id === pivotFilterPopover.pivotId);
            const field = pivot?.fieldCatalog.fields.find((candidate) => candidate.fieldId === pivotFilterPopover.fieldId);
            const placement = pivot ? [...pivot.layout.rows, ...pivot.layout.columns].find((candidate) => candidate.fieldId === pivotFilterPopover.fieldId) : undefined;
            const currentFilter = pivot?.layout.filters.find((candidate) => candidate.fieldId === pivotFilterPopover.fieldId && (candidate.scope ?? 'report') === pivotFilterPopover.scope);
            const valueFieldId = pivot?.layout.values[0]?.fieldId;
            return pivot && field ? <PivotHeaderFilterPopover locale={locale} scope={pivotFilterPopover.scope} x={pivotFilterPopover.x} y={pivotFilterPopover.y} field={field} valueFieldId={valueFieldId} currentFilter={currentFilter} currentSort={placement?.sort} onClose={() => setPivotFilterPopover(null)} onApply={(filter, sort) => { onApplyPivotFilter(pivot.id, field.fieldId, filter, sort, pivotFilterPopover.scope); setPivotFilterPopover(null); }} /> : null;
          })() : null}

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
              x={filterPopover.x}
              y={filterPopover.y}
              sheet={sheet}
              onApply={(patch) => {
                onApplyFilter(filterPopover.column, patch);
                setFilterPopover(null);
              }}
              onSort={(ascending) => {
                onSortFilterColumn(filterPopover.column, ascending);
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

function pivotProjectionStatusMessage(projection: PivotGridProjection, locale: Locale): string | null {
  if (projection.collision.status === "collision") {
    const reason = projection.collision.reasons.length > 0
      ? projection.collision.reasons.join(", ")
      : "target range is occupied";
    return `${pivotText(locale, 'pivotCollision')}: ${reason}`;
  }
  if (projection.refresh.status === "error") {
    return `${pivotText(locale, 'pivotRefreshFailed')}: ${projection.refresh.error ?? ''}`;
  }
  if (projection.refresh.status === "refreshing") return pivotText(locale, 'loadingPivot');
  if (projection.refresh.status === "stale") return pivotText(locale, 'pivotRefreshRequired');
  return null;
}

function PivotProjectionStatusNotice({
  engine,
  locale,
  projection,
  scrollTick,
}: {
  engine: CanvasRenderEngine | null;
  locale: Locale;
  projection: PivotGridProjection;
  scrollTick: number;
}): React.ReactElement | null {
  void scrollTick;
  const message = pivotProjectionStatusMessage(projection, locale);
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
