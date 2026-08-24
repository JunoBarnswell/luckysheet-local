import { useMemo, useState } from "react";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import type {
  ChartDrawingPayload,
  DrawingObject,
  PivotAggregateFunction,
  PivotFieldDefinition,
  PivotLayout,
  PivotSource,
  ShapeDrawingPayload,
  SparklineModel,
} from "@react-sheets/core-model";
import {
  type SelectionState,
  type SidebarPanelId,
  type UiSessionIntent,
  type UiSnapshot,
  type WorkbookSession,
} from "@react-sheets/spreadsheet-app";
import { parseRangeInput } from "../domain/range-input";
import type { PivotPanelCallbacks, PivotPanelState, PivotSlicerControl, PivotTimelineControl } from "../components/pivot/pivot-contract";

export interface EditorCommandControllerOptions {
  session: WorkbookSession;
  state: UiSnapshot;
  dispatchCommand: (descriptor: CommandDescriptor) => void;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
}

export interface EditorCommandController {
  selectedRange: SelectionState["ranges"][number] | undefined;
  currentDataRange: ReturnType<WorkbookSession["getCurrentRegion"]>;
  sortColumns: UiSnapshot["selectedSheet"]["columns"];
  pivotSourceRange: ReturnType<WorkbookSession["getCurrentRegion"]>;
  activePivotId: string | undefined;
  setActivePivotId: (pivotId: string | undefined) => void;
  activePivot: UiSnapshot["selectedSheet"]["pivots"][number] | undefined;
  activePivotSheetId: string;
  activePivotSourceRange: ReturnType<WorkbookSession["getCurrentRegion"]> | undefined;
  pivotFields: PivotFieldDefinition[];
  pivotSlicerControls: PivotSlicerControl[];
  pivotTimelineControls: PivotTimelineControl[];
  pivotPanelState: PivotPanelState;
  pivotCallbacks: PivotPanelCallbacks;
  pivotSourceOptions: Array<{ id: string; label: string; source: PivotSource }>;
  selectedDrawing: UiSnapshot["selectedSheet"]["drawings"][number] | undefined;
  buildQuickChartCommand: () => CommandDescriptor;
  buildQuickSparklineCommand: () => CommandDescriptor;
  buildQuickShapeCommand: () => CommandDescriptor;
  buildDrawingCommand: (commandId: "drawing.zorder" | "drawing.remove", direction?: "forward" | "backward") => CommandDescriptor | undefined;
  buildTotalRowCommand: () => CommandDescriptor | undefined;
  buildSubtotalCommand: () => CommandDescriptor;
  buildRemoveDuplicatesCommand: () => CommandDescriptor;
  buildTextToColumnsCommand: () => CommandDescriptor;
  buildOutlineCommand: (axis: "row" | "column", action: "add" | "remove") => CommandDescriptor | undefined;
  buildFilterSelectionCommand: () => CommandDescriptor;
  buildClearFilterCommand: () => CommandDescriptor;
  buildSortDescriptor: (ascending: boolean) => CommandDescriptor | undefined;
  createPivotFromDialog: (request: { sourceId: string; destination: "new-sheet" | "existing-sheet"; targetReference?: string }) => void;
  executeShortcut: (id: string) => boolean;
  selectPanel: (panel: SidebarPanelId) => void;
  applySelection: (selection: SelectionState) => void;
}

type RangeLike = ReturnType<WorkbookSession["getCurrentRegion"]>;

function createWebId(prefix: string): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return `${prefix}-${crypto.randomUUID()}`;
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function cloneLayout(layout: PivotLayout): PivotLayout {
  return structuredClone(layout);
}

function removePivotField(layout: PivotLayout, fieldId: string): PivotLayout {
  const next = cloneLayout(layout);
  next.filters = next.filters.filter((filter) => filter.fieldId !== fieldId);
  next.rows = next.rows.filter((field) => field.fieldId !== fieldId);
  next.columns = next.columns.filter((field) => field.fieldId !== fieldId);
  next.values = next.values.filter((value) => value.fieldId !== fieldId);
  return next;
}

/**
 * Owns selection-derived command construction and Pivot panel actions.
 * The route only wires these typed actions into visual hosts; it never creates
 * chart, shape or pivot drawing domain objects itself.
 */
export function useEditorCommandController({
  session,
  state,
  dispatchCommand,
  dispatchSessionIntent,
}: EditorCommandControllerOptions): EditorCommandController {
  const [activePivotId, setActivePivotId] = useState<string>();
  const selectedRange = state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0];
  const currentDataRange = session.getCurrentRegion();
  const sortColumns = state.selectedSheet.columns.slice(currentDataRange.startColumn, currentDataRange.endColumn + 1);
  const pivotSourceRange = selectedRange && (selectedRange.endRow > selectedRange.startRow || selectedRange.endColumn > selectedRange.startColumn)
    ? selectedRange
    : state.selectedSheet.usedRange;

  const selectedDrawing = state.selectedFloatingId
    ? state.selectedSheet.drawings.find((drawing) => drawing.id === state.selectedFloatingId)
    : undefined;

  const activePivot = state.selectedSheet.pivots.find((pivot) => pivot.id === activePivotId) ?? state.selectedSheet.pivots[0];
  const pivotTree = activePivot ? state.selectedSheet.pivotResults[activePivot.id] : undefined;
  const pivotFields: PivotFieldDefinition[] = pivotTree?.fields.fields ?? session.getPivotFieldCatalog(pivotSourceRange);
  const activePivotSheetId = activePivot ? activePivot.target.sheetId : state.activeSheetId;
  const activePivotSourceRange = activePivot?.source.kind === "worksheet-range" ? activePivot.source.range : undefined;
  const pivotControlRecords = activePivot ? session.listPivotControls(activePivot.id) : [];
  const pivotSlicerControls = pivotControlRecords.flatMap((record) => record.payload.kind === "slicer"
    ? [{ id: record.drawing.id, pivotId: record.payload.pivotId, fieldId: record.payload.fieldId, mode: record.payload.filter.mode, memberKeys: record.payload.filter.memberKeys, connectedPivotIds: record.payload.connectedPivotIds }]
    : []);
  const pivotTimelineControls = pivotControlRecords.flatMap((record) => record.payload.kind === "timeline"
    ? [{ id: record.drawing.id, pivotId: record.payload.pivotId, fieldId: record.payload.fieldId, start: record.payload.period.start, end: record.payload.period.end, connectedPivotIds: record.payload.connectedPivotIds }]
    : []);

  const buildQuickChartCommand = (): CommandDescriptor => {
    const chartId = createWebId("chart");
    const drawing: DrawingObject = {
      id: createWebId("drawing"),
      sheetId: state.activeSheetId,
      kind: "chart",
      payloadId: chartId,
      anchor: { kind: "absolute" },
      transform: { x: 96, y: 96, width: 480, height: 280, rotation: 0 },
      zIndex: 0,
    };
    const payload: ChartDrawingPayload = {
      kind: "chart",
      chartId,
      chartType: "column",
      title: "Chart",
      sourceRanges: [{ ...pivotSourceRange, sheetId: state.activeSheetId }],
      legendPosition: "bottom",
      showDataLabels: false,
    };
    return { commandId: "chart.insert", params: { sheetId: state.activeSheetId, drawing, payload } };
  };

  const buildQuickSparklineCommand = (): CommandDescriptor => {
    const sparklineId = createWebId("sparkline");
    return {
      commandId: "sparkline.insertDataLocation",
      params: {
        sheetId: state.activeSheetId,
        sparklineId,
        dataRange: { ...pivotSourceRange, sheetId: state.activeSheetId },
        location: { row: pivotSourceRange.startRow, column: pivotSourceRange.endColumn + 1 },
        type: "line" as SparklineModel["type"],
        highlightMax: true,
        highlightMin: true,
      },
    };
  };

  const buildQuickShapeCommand = (): CommandDescriptor => {
    const payloadId = createWebId("shape");
    const drawingId = createWebId("drawing");
    const payload: ShapeDrawingPayload = {
      kind: "shape",
      type: "rectangle",
      fill: "#dbeafe",
      stroke: "#2563eb",
      strokeWidth: 2,
      textColor: "#1e3a8a",
      fontSize: 13,
    };
    return {
      commandId: "drawing.add.shape",
      params: {
        sheetId: state.activeSheetId,
        drawing: { id: drawingId, sheetId: state.activeSheetId, kind: "shape", payloadId, anchor: { kind: "absolute" }, transform: { x: 96, y: 96, width: 160, height: 60, rotation: 0 }, zIndex: 0 },
        payload,
      },
    };
  };

  const buildDrawingCommand = (commandId: "drawing.zorder" | "drawing.remove", direction?: "forward" | "backward"): CommandDescriptor | undefined => {
    if (!selectedDrawing) return undefined;
    return { commandId, params: { sheetId: state.activeSheetId, drawingId: selectedDrawing.id, ...(direction ? { direction } : {}) } };
  };

  const buildTotalRowCommand = (): CommandDescriptor | undefined => {
    const table = state.selectedSheet.sheetTables.find((entry) =>
      pivotSourceRange.startRow >= entry.range.startRow
      && pivotSourceRange.startRow <= entry.range.endRow
      && pivotSourceRange.startColumn >= entry.range.startColumn
      && pivotSourceRange.startColumn <= entry.range.endColumn);
    return table ? { commandId: "sheetTable.toggleTotalRow", params: { sheetId: state.activeSheetId, tableId: table.id, enabled: !table.hasTotalRow } } : undefined;
  };
  const buildSubtotalCommand = (): CommandDescriptor => ({
    commandId: "data.subtotal",
    params: { sheetId: state.activeSheetId, range: pivotSourceRange, groupColumn: pivotSourceRange.startColumn, valueColumn: pivotSourceRange.startColumn + 1, functionName: "SUM" },
  });
  const buildRemoveDuplicatesCommand = (): CommandDescriptor => ({
    commandId: "data.removeDuplicates",
    params: { sheetId: state.activeSheetId, range: pivotSourceRange, columns: Array.from({ length: pivotSourceRange.endColumn - pivotSourceRange.startColumn + 1 }, (_, index) => pivotSourceRange.startColumn + index), hasHeader: true },
  });
  const buildTextToColumnsCommand = (): CommandDescriptor => ({
    commandId: "data.textToColumns",
    params: { sheetId: state.activeSheetId, range: { ...pivotSourceRange, endColumn: pivotSourceRange.startColumn }, delimiter: ",", maxColumns: 8 },
  });
  const buildOutlineCommand = (axis: "row" | "column", action: "add" | "remove"): CommandDescriptor | undefined => {
    const start = axis === "row" ? pivotSourceRange.startRow : pivotSourceRange.startColumn;
    const end = axis === "row" ? pivotSourceRange.endRow : pivotSourceRange.endColumn;
    if (action === "add") {
      if (end <= start) return undefined;
      return { commandId: "outline.group.add", params: { sheetId: state.activeSheetId, group: { id: createWebId("outline"), axis, start, end, level: 1, collapsed: false } } };
    }
    const group = state.selectedSheet.outlineGroups.find((entry) => entry.axis === axis && entry.start >= start && entry.end <= end);
    return group ? { commandId: "outline.group.remove", params: { sheetId: state.activeSheetId, groupId: group.id } } : undefined;
  };
  const buildFilterSelectionCommand = (): CommandDescriptor => ({ commandId: "sheet.filter.toggle", params: { sheetId: state.activeSheetId, range: currentDataRange } });
  const buildClearFilterCommand = (): CommandDescriptor => ({ commandId: "sheet.filter.clearCriteria", params: { sheetId: state.activeSheetId, range: currentDataRange } });
  const buildSortDescriptor = (ascending: boolean): CommandDescriptor | undefined => {
    void ascending;
    const range = session.getCurrentRegion();
    if (range.endRow <= range.startRow) {
      session.notify("Select a data region with at least one data row before sorting");
      return undefined;
    }
    return { commandId: "data.sort.quick", params: { sheetId: state.activeSheetId, range, sortColumn: state.selection.activeCell.column, hasHeader: true } };
  };

  const pivotSourceOptions = useMemo(() => {
    const options: Array<{ id: string; label: string; source: PivotSource }> = [{
      id: "current-region",
      label: `Current region · ${state.selectedSheet.name}`,
      source: { kind: "worksheet-range", range: currentDataRange },
    }];
    for (const table of state.selectedSheet.sheetTables) options.push({ id: `sheet-table:${table.id}`, label: `Table · ${table.name}`, source: { kind: "table", tableId: table.id } });
    for (const table of state.tables) options.push({ id: `data-table:${table.id}`, label: `Data table · ${table.name}`, source: { kind: "table", tableId: table.id } });
    for (const name of state.definedNameModels) options.push({ id: `name:${name.name}`, label: `Named range · ${name.name}`, source: { kind: "named-range", name: name.name } });
    for (const source of state.dataSources) options.push({ id: `source:${source.id}`, label: `Data source · ${source.name}`, source: { kind: "data-source", dataSourceId: source.id } });
    return options;
  }, [currentDataRange, state.dataSources, state.definedNameModels, state.selectedSheet.name, state.selectedSheet.sheetTables, state.tables, state.version]);

  const createPivotFromDialog = (request: { sourceId: string; destination: "new-sheet" | "existing-sheet"; targetReference?: string }) => {
    const source = pivotSourceOptions.find((option) => option.id === request.sourceId)?.source;
    if (!source) {
      session.notify("Select a valid PivotTable source");
      return;
    }
    if (request.destination === "new-sheet") {
      if (session.createPivotTable({ source, destination: { kind: "new-sheet" } })) session.closeCreatePivotDialog();
      return;
    }
    const target = parseRangeInput(request.targetReference ?? "", state.activeSheetId);
    if (!target) {
      session.notify("Enter a valid PivotTable destination such as A1");
      return;
    }
    if (session.createPivotTable({ source, destination: { kind: "existing-sheet", sheetId: state.activeSheetId, anchor: { row: target.startRow, column: target.startColumn } } })) session.closeCreatePivotDialog();
  };

  const updatePivotLayout = (nextLayout: PivotLayout) => {
    if (!activePivot) return;
    dispatchCommand({ commandId: "pivot.update", params: { sheetId: activePivotSheetId, pivotId: activePivot.id, layout: nextLayout } });
    session.notify("Pivot layout updated");
  };
  const pivotCallbacks: PivotPanelCallbacks = {
    onCreate: () => dispatchSessionIntent({ type: "dialog.open", dialog: "create-pivot" }),
    onPivotSelect: setActivePivotId,
    onFieldAreaChange: (fieldId, area, index) => {
      if (!activePivot) return;
      const next = removePivotField(activePivot.layout, fieldId);
      if (area === "values") {
        const field = pivotFields.find((entry) => entry.fieldId === fieldId);
        const summarizeBy: PivotAggregateFunction = field?.dataType === "number" ? "sum" : "count";
        next.values.splice(Math.max(0, index), 0, { fieldId, summarizeBy });
      } else if (area === "filters") next.filters.splice(Math.max(0, index), 0, { kind: "manual", fieldId, mode: "all", memberKeys: [] });
      else next[area].splice(Math.max(0, index), 0, { fieldId });
      updatePivotLayout(next);
    },
    onRemoveField: (fieldId, area) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      if (area === "values") next.values = next.values.filter((value) => value.fieldId !== fieldId);
      else if (area === "filters") next.filters = next.filters.filter((filter) => filter.fieldId !== fieldId);
      else next[area] = next[area].filter((field) => field.fieldId !== fieldId);
      updatePivotLayout(next);
    },
    onValueChange: (value) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const index = next.values.findIndex((entry) => entry.fieldId === value.fieldId);
      if (index >= 0) {
        next.values[index] = structuredClone(value);
        updatePivotLayout(next);
      }
    },
    onCalculatedFieldsChange: (fields) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), calculatedFields: fields.map((field) => ({ ...field })) }); },
    onCalculatedItemsChange: (items) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), calculatedItems: items.map((item) => ({ ...item })) }); },
    onFilterChange: (fieldId, filter) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const existing = next.filters.find((entry) => entry.fieldId === fieldId);
      if (existing?.kind === "manual") {
        existing.mode = filter.mode;
        existing.memberKeys = [...filter.memberKeys];
      } else next.filters.push({ kind: "manual", fieldId, mode: filter.mode, memberKeys: [...filter.memberKeys] });
      updatePivotLayout(next);
      session.notify("Pivot filter updated");
    },
    onSortChange: (fieldId, sort) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), rows: activePivot.layout.rows.map((field) => field.fieldId === fieldId ? { ...field, sort } : field), columns: activePivot.layout.columns.map((field) => field.fieldId === fieldId ? { ...field, sort } : field) }); },
    onGroupChange: (fieldId, group) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), rows: activePivot.layout.rows.map((field) => field.fieldId === fieldId ? { ...field, group } : field), columns: activePivot.layout.columns.map((field) => field.fieldId === fieldId ? { ...field, group } : field) }); },
    onRefresh: () => { if (activePivot) dispatchCommand({ commandId: "pivot.refresh", params: { sheetId: activePivotSheetId, pivotId: activePivot.id } }); },
    onLayoutChange: (layout) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), compact: layout === "compact", repeatLabels: layout === "tabular" }); },
    onSlicerChange: (fieldId, enabled) => {
      if (!activePivot) return;
      if (enabled) session.createPivotSlicerControl(activePivot.id, fieldId);
      else {
        const control = pivotControlRecords.find((record) => record.payload.kind === "slicer" && record.payload.fieldId === fieldId);
        if (control) session.removePivotControl(control.drawing.id);
      }
    },
    onTimelineChange: (fieldId) => {
      if (!activePivot) return;
      if (!fieldId) {
        for (const control of pivotControlRecords.filter((record) => record.payload.kind === "timeline")) session.removePivotControl(control.drawing.id);
      } else session.createPivotTimelineControl(activePivot.id, fieldId);
    },
    onSlicerFilterChange: (slicerId, filter) => session.setPivotSlicerFilter(slicerId, filter.mode, filter.memberKeys),
    onTimelineRangeChange: (timelineId, start, end) => session.setPivotTimelinePeriod(timelineId, start || undefined, end || undefined),
    onPivotChartChange: (chart) => {
      if (!activePivot || !chart) return;
      const chartId = `pivot-chart-${activePivot.id}-${Date.now().toString(36)}`;
      const drawing: DrawingObject = { id: `drawing-${chartId}`, sheetId: activePivotSheetId, kind: "chart", payloadId: chartId, anchor: { kind: "absolute" }, transform: { x: 80, y: 80, width: 480, height: 280, rotation: 0 }, zIndex: 0 };
      const payload: ChartDrawingPayload = { kind: "chart", chartId, pivotId: activePivot.id, chartType: chart.type, title: chart.title, sourceRanges: activePivotSourceRange ? [activePivotSourceRange] : [] };
      dispatchCommand({ commandId: "pivot.chart.create", params: { sheetId: activePivotSheetId, pivotId: activePivot.id, drawing, payload } });
    },
  };

  const pivotPanelState: PivotPanelState = { disabled: state.phase !== "ready", loading: state.phase === "loading", error: state.phase === "error" ? "Pivot data could not be loaded" : undefined, empty: pivotFields.length === 0 };

  const executeShortcut = (id: string): boolean => {
    switch (id) {
      case "history.undo": session.undo(); return true;
      case "history.redo": session.redo(); return true;
      case "history.repeat": session.repeatLastCommand(); return true;
      case "clipboard.copy": session.copy(); return true;
      case "clipboard.cut": session.cut(); return true;
      case "clipboard.paste": session.paste(); return true;
      case "clipboard.pasteSpecial": dispatchSessionIntent({ type: "dialog.open", dialog: "paste-special" }); return true;
      case "clipboard.cancel": session.clearClipboard(); session.cancelFormatPainter(); return true;
      case "workbook.save": void session.saveWorkbook("Keyboard shortcut"); return true;
      case "format.bold": dispatchCommand({ commandId: "sheet.style.set", params: { style: { bold: !state.homeRibbon.style.bold } } }); return true;
      case "format.italic": dispatchCommand({ commandId: "sheet.style.set", params: { style: { italic: !state.homeRibbon.style.italic } } }); return true;
      case "format.underline": dispatchCommand({ commandId: "sheet.style.set", params: { style: { underline: !state.homeRibbon.style.underline } } }); return true;
      case "range.fillDown": session.fillSelection("down"); return true;
      case "range.clearContents": session.clearSelection("contents"); return true;
      case "cells.insert": case "cells.delete": dispatchSessionIntent({ type: "dialog.open", dialog: "shift-cells" }); return true;
      case "filter.toggle": session.applyFilterSelection(); return true;
      case "find.open": case "replace.open": dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace" }); return true;
      case "name.goto": case "navigation.goto": dispatchSessionIntent({ type: "dialog.open", dialog: "goto" }); return true;
      case "ribbon.home.keyTips": session.setRibbonTab("home"); session.notify("Home shortcuts are active"); return true;
      case "format.cells": dispatchSessionIntent({ type: "dialog.open", dialog: "format-cells" }); return true;
      case "hyperlink.insert": dispatchSessionIntent({ type: "panel.open", panel: "inspector", notice: "Use the Inspector to insert a hyperlink." }); return true;
      case "row.select": session.selectActiveRow(); return true;
      case "column.select": session.selectActiveColumn(); return true;
      case "sheet.previous": session.selectAdjacentSheet("previous"); return true;
      case "sheet.next": session.selectAdjacentSheet("next"); return true;
      case "formula.autoSum": session.autoSum(); return true;
      case "edit.begin": session.beginEdit(); return true;
      case "drawing.remove": session.removeSelectedDrawing(); return true;
      case "formula.functionWizard": dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" }); return true;
      case "formula.calculate": void session.recalculateFormulas(); return true;
      case "pivot.refresh": {
        const pivotId = state.activeContext.kind === "pivot" ? state.activeContext.pivotId : undefined;
        if (!pivotId) return false;
        session.refreshPivot(pivotId);
        return true;
      }
      default: return false;
    }
  };

  const selectPanel = (panel: SidebarPanelId) => {
    const panelNotice = panel === "inspector" ? "Select a cell and use Review tools for comments." : undefined;
    if (panel === "print") {
      dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" });
      return;
    }
    if (panelNotice) {
      dispatchSessionIntent({ type: "panel.open", panel, notice: panelNotice });
      return;
    }
    if (panel === "selectionPane") session.setDrawingSelectionMode(true);
    dispatchSessionIntent({ type: "panel.open", panel });
  };

  return {
    selectedRange,
    currentDataRange,
    sortColumns,
    pivotSourceRange,
    activePivotId,
    setActivePivotId,
    activePivot,
    activePivotSheetId,
    activePivotSourceRange,
    pivotFields,
    pivotSlicerControls,
    pivotTimelineControls,
    pivotPanelState,
    pivotCallbacks,
    pivotSourceOptions,
    selectedDrawing,
    buildQuickChartCommand,
    buildQuickSparklineCommand,
    buildQuickShapeCommand,
    buildDrawingCommand,
    buildTotalRowCommand,
    buildSubtotalCommand,
    buildRemoveDuplicatesCommand,
    buildTextToColumnsCommand,
    buildOutlineCommand,
    buildFilterSelectionCommand,
    buildClearFilterCommand,
    buildSortDescriptor,
    createPivotFromDialog,
    executeShortcut,
    selectPanel,
    applySelection: (selection) => session.applyCanvasSelection(selection),
  };
}

export type EditorRange = RangeLike;
