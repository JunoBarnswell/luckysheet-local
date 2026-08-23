import { AppShell, Box, Button, DropdownMenu, Inline, SidebarShell, Stack } from "@react-sheets/ui-system";
import { FormulaBar } from "./components/FormulaBar";
import { Ribbon } from "./components/Ribbon";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { WorkspaceErrorBoundary } from "./components/WorkspaceErrorBoundary";
import { parseRangeInput } from "./domain/range-input";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";
import {
  getInitialSessionPhase,
  useWorkbookSession,
  type SelectionState,
  type SidebarPanelId,
  type UiSessionIntent,
} from "@react-sheets/spreadsheet-app";
import { getInitialLocale, localeLabels, persistLocale, shellLabels, type Locale } from "./i18n";
import zhCN from "./locales/zh-CN.json";
import enUS from "./locales/en-US.json";
import type { ChartDrawingPayload, DrawingObject, PivotAggregateFunction, PivotFieldDefinition, PivotLayout, PivotSource, ShapeDrawingPayload, SparklineModel } from "@react-sheets/core-model";
import type { PivotPanelCallbacks, PivotPanelState } from "./components/pivot/pivot-contract";

const FeatureSidebar = lazy(() => import("./components/FeatureSidebar").then((module) => ({ default: module.FeatureSidebar })));
const FunctionWizardDialog = lazy(() => import("./components/dialogs/FunctionWizardDialog").then((module) => ({ default: module.FunctionWizardDialog })));
const SortDialog = lazy(() => import("./components/dialogs/SortDialog").then((module) => ({ default: module.SortDialog })));
const FindReplaceDialog = lazy(() => import("./components/dialogs/FindReplaceDialog").then((module) => ({ default: module.FindReplaceDialog })));
const GoToDialog = lazy(() => import("./components/dialogs/GoToDialog").then((module) => ({ default: module.GoToDialog })));
const PasteSpecialDialog = lazy(() => import("./components/dialogs/PasteSpecialDialog").then((module) => ({ default: module.PasteSpecialDialog })));
const FormatCellsDialog = lazy(() => import("./components/dialogs/FormatCellsDialog").then((module) => ({ default: module.FormatCellsDialog })));
const ShiftCellsDialog = lazy(() => import("./components/dialogs/ShiftCellsDialog").then((module) => ({ default: module.ShiftCellsDialog })));
const CreatePivotTableDialog = lazy(() => import("./components/dialogs/CreatePivotTableDialog").then((module) => ({ default: module.CreatePivotTableDialog })));
const PrintPreviewDialog = lazy(() => import("./components/dialogs/PrintPreviewDialog").then((module) => ({ default: module.PrintPreviewDialog })));
const WorkbookCatalog = lazy(() => import("./components/WorkbookCatalog").then((module) => ({ default: module.WorkbookCatalog })));
const SheetCanvas = lazy(() => import("./components/SheetCanvas").then((module) => ({ default: module.SheetCanvas })));

function WorkspaceApp() {
  const { session, snapshot: state } = useWorkbookSession({ initialPhase: getInitialSessionPhase() });
  const [locale, setLocaleState] = useState<Locale>(() => getInitialLocale());
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const previousPanelRef = useRef(state.activePanel);
  const isBusy = state.phase !== "ready";
  const setLocale = (nextLocale: Locale) => {
    setLocaleState(nextLocale);
    persistLocale(nextLocale);
  };

  const dispatchCommand = (descriptor: CommandDescriptor) => {
    session.dispatch(descriptor);
  };

  const dispatchSessionIntent = (intent: UiSessionIntent) => {
    if (intent.type === 'panel.open') setSidebarOpen(true);
    session.dispatchUiSessionIntent(intent);
  };

  useEffect(() => {
    if (previousPanelRef.current !== state.activePanel) {
      setSidebarOpen(true);
    }
    previousPanelRef.current = state.activePanel;
  }, [state.activePanel]);

  const copyWorkbookLink = () => { void session.createGuestShareLink('editor'); };

  const saveWorkbook = () => { void session.saveWorkbook("Ribbon save"); };

  const exportXlsx = async () => {
    const exported = await session.exportXlsxWorkbook();
    if (!exported) return;
    const href = URL.createObjectURL(new Blob([exported.buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const link = document.createElement("a");
    link.href = href;
    link.download = exported.fileName;
    link.click();
    URL.revokeObjectURL(href);
  };

  const importXlsx = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    input.onchange = () => {
      const file = input.files?.[0];
      if (!file) return;
      void file.arrayBuffer().then((buffer) => session.importXlsxBuffer(buffer, file.name));
    };
    input.click();
  };

  const selectedRange = state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0];
  const currentDataRange = session.getCurrentRegion();
  const sortColumns = state.selectedSheet.columns.slice(currentDataRange.startColumn, currentDataRange.endColumn + 1);
  const pivotSourceRange = selectedRange && (selectedRange.endRow > selectedRange.startRow || selectedRange.endColumn > selectedRange.startColumn)
    ? selectedRange
    : state.selectedSheet.usedRange;
  const createWebId = (prefix: string): string => {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return `${prefix}-${crypto.randomUUID()}`;
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  };
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
    return {
      commandId: "chart.insert",
      params: {
        sheetId: state.activeSheetId,
        drawing,
        payload,
      },
    };
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
  const selectedDrawing = state.selectedFloatingId ? state.selectedSheet.drawings.find((drawing) => drawing.id === state.selectedFloatingId) : undefined;
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
      return {
        commandId: "outline.group.add",
        params: { sheetId: state.activeSheetId, group: { id: createWebId("outline"), axis, start, end, level: 1, collapsed: false } },
      };
    }
    const group = state.selectedSheet.outlineGroups.find((entry) => entry.axis === axis && entry.start >= start && entry.end <= end);
    return group ? { commandId: "outline.group.remove", params: { sheetId: state.activeSheetId, groupId: group.id } } : undefined;
  };
  const buildFilterSelectionCommand = (): CommandDescriptor => ({
    commandId: "sheet.filter.toggle",
    params: { sheetId: state.activeSheetId, range: session.getCurrentRegion() },
  });
  const buildSortDescriptor = (ascending: boolean): CommandDescriptor | undefined => {
    const range = session.getCurrentRegion();
    if (range.endRow <= range.startRow) {
      session.notify("Select a data region with at least one data row before sorting");
      return undefined;
    }
    return {
      commandId: "data.sort.quick",
      params: {
        sheetId: state.activeSheetId,
        range,
        sortColumn: state.selection.activeCell.column,
        hasHeader: true,
      },
    };
  };
  const pivotSourceOptions = useMemo((): Array<{ id: string; label: string; source: PivotSource }> => {
    const options: Array<{ id: string; label: string; source: PivotSource }> = [{
      id: 'current-region',
      label: `Current region · ${state.selectedSheet.name}`,
      source: { kind: 'worksheet-range', range: session.getCurrentRegion() },
    }];
    for (const table of state.selectedSheet.sheetTables) {
      options.push({ id: `sheet-table:${table.id}`, label: `Table · ${table.name}`, source: { kind: 'table', tableId: table.id } });
    }
    for (const table of state.tables) {
      options.push({ id: `data-table:${table.id}`, label: `Data table · ${table.name}`, source: { kind: 'table', tableId: table.id } });
    }
    for (const name of state.definedNameModels) {
      options.push({ id: `name:${name.name}`, label: `Named range · ${name.name}`, source: { kind: 'named-range', name: name.name } });
    }
    for (const source of state.dataSources) {
      options.push({ id: `source:${source.id}`, label: `Data source · ${source.name}`, source: { kind: 'data-source', dataSourceId: source.id } });
    }
    return options;
  }, [session, state.dataSources, state.definedNameModels, state.selectedSheet.name, state.selectedSheet.sheetTables, state.tables, state.version]);
  const createPivotFromDialog = (request: { sourceId: string; destination: "new-sheet" | "existing-sheet"; targetReference?: string }) => {
    const source = pivotSourceOptions.find((option) => option.id === request.sourceId)?.source;
    if (!source) {
      session.notify("Select a valid PivotTable source");
      return;
    }
    if (request.destination === "new-sheet") {
      const created = session.createPivotTable({ source, destination: { kind: "new-sheet" } });
      if (created) session.closeCreatePivotDialog();
      return;
    }
    const target = parseRangeInput(request.targetReference ?? "", state.activeSheetId);
    if (!target) {
      session.notify("Enter a valid PivotTable destination such as A1");
      return;
    }
    const created = session.createPivotTable({
      source,
      destination: { kind: "existing-sheet", sheetId: state.activeSheetId, anchor: { row: target.startRow, column: target.startColumn } },
    });
    if (created) session.closeCreatePivotDialog();
  };
  const executeShortcut = (id: string): boolean => {
    switch (id) {
      case "history.undo": session.undo(); return true;
      case "history.redo": session.redo(); return true;
      case "history.repeat": session.repeatLastCommand(); return true;
      case "clipboard.copy": session.copy(); return true;
      case "clipboard.cut": session.cut(); return true;
      case "clipboard.paste": session.paste(); return true;
      case "workbook.save": void session.saveWorkbook("Keyboard shortcut"); return true;
      case "format.bold": dispatchCommand({ commandId: "sheet.style.set", params: { style: { bold: !selectedCellStyle.bold } } }); return true;
      case "format.italic": dispatchCommand({ commandId: "sheet.style.set", params: { style: { italic: !selectedCellStyle.italic } } }); return true;
      case "format.underline": dispatchCommand({ commandId: "sheet.style.set", params: { style: { underline: !selectedCellStyle.underline } } }); return true;
      case "find.open": dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace" }); return true;
      case "replace.open": dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace" }); return true;
      case "name.goto": dispatchSessionIntent({ type: "dialog.open", dialog: "goto" }); return true;
      case "format.cells": dispatchSessionIntent({ type: "dialog.open", dialog: "format-cells" }); return true;
      case "hyperlink.insert": dispatchSessionIntent({ type: "panel.open", panel: "inspector", notice: "Use the Inspector to insert a hyperlink." }); return true;
      case "row.select": session.selectActiveRow(); return true;
      case "column.select": session.selectActiveColumn(); return true;
      case "sheet.previous": session.selectAdjacentSheet("previous"); return true;
      case "sheet.next": session.selectAdjacentSheet("next"); return true;
      case "formula.autoSum": session.autoSum(); return true;
      case "formula.functionWizard": dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" }); return true;
      case "formula.calculate": void session.recalculateFormulas(); return true;
      case "pivot.refresh": {
        const pivotId = state.activeContext.kind === 'pivot' ? state.activeContext.pivotId : undefined;
        if (!pivotId) return false;
        session.refreshPivot(pivotId);
        return true;
      }
      case "drawing.remove": session.removeSelectedDrawing(); return true;
      default: return false;
    }
  };
  const buildClearFilterCommand = (): CommandDescriptor => ({ commandId: "sheet.filter.clearCriteria", params: { sheetId: state.activeSheetId, range: session.getCurrentRegion() } });
  const [activePivotId, setActivePivotId] = useState<string>();
  const activePivot = state.selectedSheet.pivots.find((pivot) => pivot.id === activePivotId) ?? state.selectedSheet.pivots[0];
  const pivotTree = activePivot ? state.selectedSheet.pivotResults[activePivot.id] : undefined;
  const corePivotFields = pivotTree?.fields.fields ?? session.getPivotFieldCatalog(pivotSourceRange);
  const pivotFields: PivotFieldDefinition[] = corePivotFields;
  const activePivotSheetId = activePivot ? activePivot.target.sheetId : state.activeSheetId;
  const activePivotSourceRange = activePivot?.source.kind === "worksheet-range" ? activePivot.source.range : undefined;
  const pivotControlRecords = activePivot ? session.listPivotControls(activePivot.id) : [];
  const pivotSlicerControls = pivotControlRecords.flatMap((record) => record.payload.kind === "slicer"
    ? [{ id: record.drawing.id, pivotId: record.payload.pivotId, fieldId: record.payload.fieldId, mode: record.payload.filter.mode, memberKeys: record.payload.filter.memberKeys, connectedPivotIds: record.payload.connectedPivotIds }]
    : []);
  const pivotTimelineControls = pivotControlRecords.flatMap((record) => record.payload.kind === "timeline"
    ? [{ id: record.drawing.id, pivotId: record.payload.pivotId, fieldId: record.payload.fieldId, start: record.payload.period.start, end: record.payload.period.end, connectedPivotIds: record.payload.connectedPivotIds }]
    : []);

  const cloneLayout = (layout: PivotLayout): PivotLayout => structuredClone(layout);
  const removeField = (layout: PivotLayout, fieldId: string): PivotLayout => {
    const next = cloneLayout(layout);
    next.filters = next.filters.filter((filter) => filter.fieldId !== fieldId);
    next.rows = next.rows.filter((field) => field.fieldId !== fieldId);
    next.columns = next.columns.filter((field) => field.fieldId !== fieldId);
    next.values = next.values.filter((value) => value.fieldId !== fieldId);
    return next;
  };
  const updatePivotLayout = (nextLayout: PivotLayout) => {
    if (!activePivot) return;
    dispatchCommand({ commandId: "pivot.update", params: { sheetId: activePivotSheetId, pivotId: activePivot.id, layout: nextLayout } });
    session.notify("Pivot layout updated");
  };

  const createPivot = () => dispatchSessionIntent({ type: "dialog.open", dialog: "create-pivot" });

  const pivotCallbacks: PivotPanelCallbacks = {
    onCreate: createPivot,
    onPivotSelect: setActivePivotId,
    onFieldAreaChange: (fieldId, area, index) => {
      if (!activePivot) return;
      const next = removeField(activePivot.layout, fieldId);
      if (area === "values") {
        const field = pivotFields.find((entry) => entry.fieldId === fieldId);
        const summarizeBy: PivotAggregateFunction = field?.dataType === "number" ? "sum" : "count";
        next.values.splice(Math.max(0, index), 0, { fieldId, summarizeBy });
      } else if (area === "filters") {
        next.filters.splice(Math.max(0, index), 0, { kind: "manual", fieldId, mode: "all", memberKeys: [] });
      } else {
        next[area].splice(Math.max(0, index), 0, { fieldId });
      }
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
      const valueFieldId = value.fieldId;
      const index = next.values.findIndex((entry) => entry.fieldId === valueFieldId);
      if (index < 0) return;
      next.values[index] = structuredClone(value);
      updatePivotLayout(next);
    },
    onCalculatedFieldsChange: (fields) => {
      if (!activePivot) return;
      updatePivotLayout({ ...cloneLayout(activePivot.layout), calculatedFields: fields.map((field) => ({ ...field })) });
    },
    onCalculatedItemsChange: (items) => {
      if (!activePivot) return;
      updatePivotLayout({
        ...cloneLayout(activePivot.layout),
        calculatedItems: items.map((item) => ({ ...item })),
      });
    },
    onFilterChange: (fieldId, filter) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const existing = next.filters.find((filter) => filter.fieldId === fieldId);
      if (existing?.kind === "manual") {
        existing.fieldId = fieldId;
        existing.mode = filter.mode;
        existing.memberKeys = [...filter.memberKeys];
      } else next.filters.push({ kind: "manual", fieldId, mode: filter.mode, memberKeys: [...filter.memberKeys] });
      updatePivotLayout(next);
      session.notify("Pivot filter updated");
    },
    onSortChange: (fieldId, sort) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      next.rows = next.rows.map((field) => field.fieldId === fieldId ? { ...field, fieldId, sort } : field);
      next.columns = next.columns.map((field) => field.fieldId === fieldId ? { ...field, fieldId, sort } : field);
      updatePivotLayout(next);
    },
    onGroupChange: (fieldId, group) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      next.rows = next.rows.map((entry) => entry.fieldId === fieldId ? { ...entry, fieldId, group } : entry);
      next.columns = next.columns.map((entry) => entry.fieldId === fieldId ? { ...entry, fieldId, group } : entry);
      updatePivotLayout(next);
    },
    onRefresh: () => {
      if (!activePivot) return;
      dispatchCommand({ commandId: "pivot.refresh", params: { sheetId: activePivotSheetId, pivotId: activePivot.id } });
    },
    onLayoutChange: (layout) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      next.compact = layout === "compact";
      next.repeatLabels = layout === "tabular";
      updatePivotLayout(next);
    },
    onSlicerChange: (fieldId, enabled) => {
      if (!activePivot) return;
      if (enabled) {
        session.createPivotSlicerControl(activePivot.id, fieldId);
        return;
      }
      const control = pivotControlRecords.find((record) => record.payload.kind === "slicer" && record.payload.fieldId === fieldId);
      if (control) session.removePivotControl(control.drawing.id);
    },
    onTimelineChange: (fieldId) => {
      if (!activePivot) return;
      if (!fieldId) {
        for (const control of pivotControlRecords.filter((record) => record.payload.kind === "timeline")) session.removePivotControl(control.drawing.id);
        return;
      }
      session.createPivotTimelineControl(activePivot.id, fieldId);
    },
    onSlicerFilterChange: (slicerId, filter) => {
      session.setPivotSlicerFilter(slicerId, filter.mode, filter.memberKeys);
    },
    onTimelineRangeChange: (timelineId, start, end) => {
      session.setPivotTimelinePeriod(timelineId, start || undefined, end || undefined);
    },
    onPivotChartChange: (chart) => {
      if (!activePivot || !chart) return;
      const chartId = `pivot-chart-${activePivot.id}-${Date.now().toString(36)}`;
      const drawing: DrawingObject = {
        id: `drawing-${chartId}`,
        sheetId: activePivotSheetId,
        kind: "chart",
        payloadId: chartId,
        anchor: { kind: "absolute" },
        transform: { x: 80, y: 80, width: 480, height: 280, rotation: 0 },
        zIndex: 0,
      };
      const payload: ChartDrawingPayload = {
        kind: "chart",
        chartId,
        pivotId: activePivot.id,
        chartType: chart.type,
        title: chart.title,
        sourceRanges: activePivotSourceRange ? [activePivotSourceRange] : [],
      };
      dispatchCommand({ commandId: "pivot.chart.create", params: { sheetId: activePivotSheetId, pivotId: activePivot.id, drawing, payload } });
    },
  };
  const pivotPanelState: PivotPanelState = {
    disabled: isBusy,
    loading: state.phase === "loading",
    error: state.phase === "error" ? "Pivot data could not be loaded" : undefined,
    empty: pivotFields.length === 0,
  };
  const selectPanel = (panel: SidebarPanelId) => {
    setSidebarOpen(true);
    const panelNotice = panel === "inspector" ? "Select a cell and use Review tools for comments." : undefined;
    if (panel === "print") {
      dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" });
      return;
    }
    if (panelNotice) {
      dispatchSessionIntent({ type: "panel.open", panel, notice: panelNotice });
      return;
    }
    if (panel === "inspector" || panel === "chart" || panel === "pivot" || panel === "shape" || panel === "selectionPane" || panel === "sparkline" || panel === "conditionalFormat" || panel === "dataValidation" || panel === "history" || panel === "data") {
      if (panel === 'selectionPane') session.setDrawingSelectionMode(true);
      dispatchSessionIntent({ type: "panel.open", panel });
      return;
    }
    if (panel === "query" || panel === "automate" || panel === "extended") {
      dispatchSessionIntent({ type: "panel.open", panel });
      return;
    }
    dispatchSessionIntent({ type: "panel.open", panel });
  };

  const applySelection = (selection: SelectionState) => {
    session.applyCanvasSelection(selection);
  };

  const formatCellsInitial = useMemo(() => {
    const style = state.homeRibbon.style;
    return {
      numberFormat: style.numberFormat ?? "general",
      style: { ...style },
    };
  }, [state.homeRibbon.style, state.showFormatCells, state.version]);
  const selectedCellStyle = state.homeRibbon.style;

  return (
    <>
      <AppShell
        formulaBar={
          <FormulaBar
            cellName={state.activeCell}
            disabled={isBusy}
            formula={state.formulaDraft}
            locale={locale}
            onCancel={session.cancelEdit.bind(session)}
            onChange={session.setFormulaDraft.bind(session)}
            onCommit={() => {
              if (state.editingCell) session.commitEdit("down");
              else session.commitFormula();
            }}
            onNameBoxCommit={(value) => session.selectAddress(value)}
            onOpenWizard={() => dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" })}
            phase={state.phase}
          />
        }
        isBusy={isBusy}
        labels={shellLabels(locale, state.saveState)}
        localeMenuLabel={localeLabels[locale]}
        notice={state.notice}
        onLocaleChange={setLocale}
        onSearch={(query) => dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace", findQuery: query })}
        onShare={copyWorkbookLink}
        peers={state.peers}
        workbookMenu={
          <DropdownMenu
            align="right"
            trigger={<Button aria-label="Open workbook menu" disabled={isBusy} icon="more-horizontal" iconOnly size="sm" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white" />}
          >
            {({ close }) => (
              <Stack gap="xs" className="min-w-44">
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => window.location.assign("/workbooks")}>
                  Open workbooks
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => {
                  const nextName = window.prompt('Enter workbook name:', state.workbookName);
                  if (nextName?.trim()) session.renameWorkbook(nextName.trim());
                  close();
                }}>
                  Rename workbook
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); copyWorkbookLink(); }}>
                  Copy workbook link
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); void exportXlsx(); }}>
                  Export .xlsx
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); importXlsx(); }}>
                  Import .xlsx
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" }); }}>
                  Print / Save as PDF
                </Button>
              </Stack>
            )}
          </DropdownMenu>
        }
        ribbon={
          <Ribbon
            activeTab={state.ribbonTab}
            activePivot={state.activeContext.kind === 'pivot'
              ? { sheetId: state.activeContext.sheetId, pivotId: state.activeContext.pivotId }
              : undefined}
            locale={locale}
            onCommand={dispatchCommand}
            onSessionIntent={dispatchSessionIntent}
            onCopy={() => session.copy()}
            onCut={() => session.cut()}
            onPaste={() => session.paste()}
            onUndo={() => session.undo()}
            onRedo={() => session.redo()}
            onSave={saveWorkbook}
            onExportXlsx={() => void exportXlsx()}
            onImportXlsx={importXlsx}
            onRecalculate={() => session.recalculateFormulas()}
            onTracePrecedents={() => session.showFormulaPrecedents()}
            onTraceDependents={() => session.showFormulaDependents()}
            onRemoveArrows={() => session.removeFormulaAuditArrows()}
            onToggleShowFormulas={() => session.setShowFormulas(!state.formulaAudit.showFormulas)}
            onScanFormulaErrors={() => session.scanFormulaErrors()}
            onEvaluateFormula={() => session.evaluateFormulaStep()}
            onOpenPrintLayout={() => session.openPrintLayout()}
            onSetPrintArea={() => session.setCurrentPrintArea()}
            onClearPrintArea={() => session.clearPrintArea()}
            onSetPrintTitleRows={() => session.setPrintTitles("rows")}
            onSetPrintTitleColumns={() => session.setPrintTitles("columns")}
            onSetPrintScale={(scale) => session.setPrintScale(scale)}
            onToggleViewGridlines={() => session.toggleViewGridlines()}
            onTogglePrintGridlines={() => session.togglePrintGridlines()}
            onToggleViewHeadings={() => session.toggleViewHeadings()}
            onTogglePrintHeadings={() => session.togglePrintHeadings()}
            onAutoSum={() => session.autoSum()}
            onFreezeAtPrimary={() => session.freezeAtPrimary()}
            onCreatePivotDialog={() => dispatchSessionIntent({ type: "dialog.open", dialog: "create-pivot" })}
            buildSortDescriptor={buildSortDescriptor}
            onCreatePivot={() => session.buildQuickPivotDescriptor()}
            onCreateChart={buildQuickChartCommand}
            onCreateSparkline={buildQuickSparklineCommand}
            onCreateShape={buildQuickShapeCommand}
            onBringDrawingForward={() => buildDrawingCommand("drawing.zorder", "forward")}
            onSendDrawingBackward={() => buildDrawingCommand("drawing.zorder", "backward")}
            onRemoveDrawing={() => buildDrawingCommand("drawing.remove")}
            onCreateSheetTable={() => session.createSheetTableFromSelection()}
            onCreateDataTable={() => session.createDataTableFromSelection()}
            onToggleSheetTableTotalRow={buildTotalRowCommand}
            onApplyFilterSelection={buildFilterSelectionCommand}
            onClearFilter={buildClearFilterCommand}
            onGroupRows={() => buildOutlineCommand("row", "add")}
            onUngroupRows={() => buildOutlineCommand("row", "remove")}
            onGroupColumns={() => buildOutlineCommand("column", "add")}
            onUngroupColumns={() => buildOutlineCommand("column", "remove")}
            onSubtotal={buildSubtotalCommand}
            onRemoveDuplicates={buildRemoveDuplicatesCommand}
            onTextToColumns={buildTextToColumnsCommand}
            onResolveComment={() => session.resolveComment()}
            onProtectSelection={() => session.protectSelection()}
            onUnprotectSelection={() => session.unprotectSelection()}
            onShowOutlineLevel={(level) => session.showOutlineLevel(level)}
            onTransposeSelection={() => session.transposeSelection()}
            onFlipSelection={(axis) => session.flipSelection(axis)}
            onSplitByDelimiter={() => session.splitByDelimiter(',')}
            onToggleBandedRows={() => session.toggleBandedRows()}
            onSetRecalculationMode={(mode) => session.setRecalculationMode(mode)}
            onOpenDefinedNames={() => dispatchSessionIntent({ type: 'panel.open', panel: 'definedNames' })}
            onTabChange={(tab) => session.setRibbonTab(tab)}
            phase={state.phase}
            homeState={state.homeRibbon}
            formatPainterActive={state.formatPainter !== null}
            onBeginFormatPainter={(locked) => session.beginFormatPainter(Boolean(locked))}
            canExecute={session.canExecute.bind(session)}
          />
        }
        saveState={state.saveState}
        sheetTabs={
          <SheetTabs
            activeSheetId={state.activeSheetId}
            locale={locale}
            disabled={isBusy}
            onAdd={session.addSheet.bind(session)}
            onSelect={session.selectSheet.bind(session)}
            onRenameSheet={session.renameSheet.bind(session)}
            onDeleteSheet={session.deleteSheet.bind(session)}
            onDuplicateSheet={session.duplicateSheet.bind(session)}
            onHideSheet={session.hideSheet.bind(session)}
            onSetTabColor={session.setSheetTabColor.bind(session)}
            onMoveSheet={session.moveSheet.bind(session)}
            sheets={state.sheets}
          />
        }
        statusBar={
          <StatusBar
            activeCell={state.activeCell}
            locale={locale}
            onOpenShortcuts={() => session.notify("Shortcuts: Arrows / Tab / Enter / F2 / F4 / Ctrl+C/X/V/Z/Y/B/I/U")}
            onZoomChange={session.setZoom.bind(session)}
            phase={state.phase}
            saveState={state.saveState}
            sheetCount={state.sheets.length}
            zoom={state.zoom}
            collabStatus={state.collabStatus}
            pendingChangeSetCount={state.pendingChangeSetCount}
            collabRevision={state.collabRevision}
            hasPendingOperations={state.hasPendingOperations}
          />
        }
        title={state.workbookName}
        workspacePhase={state.phase}
      >
        <Inline gap="none" className="h-full min-h-0 w-full flex-nowrap">
          <Box className="h-full min-h-0 min-w-0 flex-1">
            <Suspense fallback={<Box className="h-full min-h-0 w-full bg-canvas" />}>
            <SheetCanvas
              sheet={state.selectedSheet}
              sheetId={state.activeSheetId}
              selection={state.selection}
              activeCell={state.activeCell}
              formulaDraft={state.formulaDraft}
              editingCell={state.editingCell}
              phase={state.phase}
              zoom={state.zoom}
              peers={state.peers}
              cellStyle={selectedCellStyle}
              selectedFloatingId={state.selectedFloatingId}
              showFormulas={state.formulaAudit.showFormulas}
              onPivotContextHit={(hit) => {
                const pivotId = hit?.pivot?.pivotId ?? hit?.objectId;
                if (pivotId) {
                  session.setActivePivotContext(pivotId, state.activeSheetId);
                  setActivePivotId(pivotId);
                  setSidebarOpen(true);
                  dispatchSessionIntent({ type: "panel.open", panel: "pivot" });
                } else {
                  session.setActivePivotContext(null);
                }
              }}
              getPivotContextMenuItems={(hit) => {
                const pivotId = hit.pivot?.pivotId ?? hit.objectId;
                if (!pivotId) return [];
                const sourceRowPaths = hit.pivot?.sourceRowPaths ?? [];
                return [
                  {
                    id: "pivot-refresh",
                    label: "Refresh PivotTable",
                    onSelect: () => dispatchCommand({ commandId: "pivot.refresh", params: { sheetId: state.activeSheetId, pivotId } }),
                  },
                  {
                    id: "pivot-show-details",
                    label: "Show Details",
                    disabled: sourceRowPaths.length === 0,
                    onSelect: () => session.showPivotDetails(pivotId, sourceRowPaths),
                  },
                ];
              }}
              onPivotShowDetails={({ pivotId, sourceRowPaths }) => session.showPivotDetails(pivotId, sourceRowPaths)}
              drawings={state.selectedSheet.drawings}
              drawingPayloads={state.selectedSheet.drawingPayloads}
              allSheets={state.sheets}
              pivotResults={state.selectedSheet.pivotResults}
              sparklines={state.selectedSheet.sparklines}
              onSelectionChange={applySelection}
              onExtendSelection={(row, column) => session.extendSelectionTo(row, column)}
              onMovePrimary={(rowDelta, columnDelta, opts) => session.movePrimary(rowDelta, columnDelta, opts)}
              onCommitCell={(value) => session.commitFormula(value)}
              onBeginEdit={(initialText) => session.beginEdit(initialText)}
              onCancelEdit={session.cancelEdit.bind(session)}
              onCommitEdit={(moveAfter) => session.commitEdit(moveAfter ?? "down")}
              onFormulaDraftChange={session.setFormulaDraft.bind(session)}
              onAppendFormulaDraft={session.appendFormulaDraft.bind(session)}
              onInsertRef={session.insertRefIntoDraft.bind(session)}
              onToggleAbsolute={session.toggleAbsoluteReference.bind(session)}
              onJumpEdge={(direction, extend) => session.jumpEdge(direction, extend)}
              onSelectAll={session.selectAll.bind(session)}
              onResizeRow={session.resizeRow.bind(session)}
              onResizeColumn={session.resizeColumn.bind(session)}
              onFillRange={session.fillRange.bind(session)}
              drawingSelectionMode={state.drawingSelectionMode}
              onExitDrawingSelectionMode={() => session.setDrawingSelectionMode(false)}
              onFloatingSelect={(hit, mode) => session.setDrawingSelection(hit ? [hit.id] : [], mode)}
              onFloatingMove={(drawingId, bounds, rotation) => dispatchCommand({
                commandId: "drawing.move",
                params: { sheetId: state.activeSheetId, drawingId, transform: { ...bounds, rotation } },
              })}
              onFloatingRemove={(drawingId) => dispatchCommand({
                commandId: "drawing.remove",
                params: { sheetId: state.activeSheetId, drawingId },
              })}
              onCommand={dispatchCommand}
              onClearSelection={(mode) => session.clearSelection(mode)}
              formatPainterActive={state.formatPainter !== null}
              onCancelFormatPainter={() => session.cancelFormatPainter()}
              onCopy={() => session.copy()}
              onCut={() => session.cut()}
              onPaste={() => session.paste()}
              onUndo={() => session.undo()}
              onRedo={() => session.redo()}
              onShortcut={executeShortcut}
              canRepeat={session.canRepeatLastCommand()}
              onOpenInspector={() => dispatchSessionIntent({ type: "panel.open", panel: "inspector", notice: "Select a cell and use Review tools for comments." })}
              onApplyFilter={(column, patch) => session.applyFilter(column, patch)}
              onToggleOutline={(groupId) => session.toggleOutlineGroup(groupId)}
              getValidationList={session.getValidationAt.bind(session)}
              onRetry={session.retry.bind(session)}
              onCreateSheet={session.addSheet.bind(session)}
            />
            </Suspense>
          </Box>
          <SidebarShell
            open={sidebarOpen}
            onOpenChange={setSidebarOpen}
            title={locale === 'zh-CN' ? zhCN.sidebar.title : enUS.sidebar.title}
          >
          <Suspense fallback={<Box className="h-full min-h-0" />}>
          <FeatureSidebar
            activeCell={state.activeCell}
            activePanel={state.activePanel}
            locale={locale}
            selectedRange={selectedRange}
            onPanelChange={selectPanel}
            onRetry={session.retry.bind(session)}
            phase={state.phase}
            sheet={state.selectedSheet}
            sheetId={state.activeSheetId}
            drawings={state.selectedSheet.drawings}
            drawingPayloads={state.selectedSheet.drawingPayloads}
            selectedDrawingIds={state.selectedDrawingIds}
            onSelectDrawing={(drawingId, mode) => session.setDrawingSelection([drawingId], mode === 'extend' ? 'add' : mode)}
            onSetDrawingVisibility={(drawingId, visible) => session.setDrawingVisibility(drawingId, visible)}
            onRenameDrawing={(drawingId, name) => session.renameDrawing(drawingId, name)}
            onReorderDrawing={(drawingId, direction) => dispatchCommand({
              commandId: 'drawing.zorder',
              params: { sheetId: state.activeSheetId, drawingId, direction },
            })}
            pivot={activePivot}
            pivotList={state.selectedSheet.pivots.map((pivot) => ({ id: pivot.id, label: pivot.id }))}
            activePivotId={activePivot?.id}
            pivotFieldCatalog={pivotFields}
            pivotSlicerControls={pivotSlicerControls}
            pivotTimelineControls={pivotTimelineControls}
            pivotPanelState={pivotPanelState}
            pivotCallbacks={pivotCallbacks}
            formulaAudit={state.formulaAudit}
            formulaAuditState={state.phase === "loading" ? "loading" : state.phase === "error" ? "error" : "ready"}
            formulaAuditError={state.phase === "error" ? "Formula audit is unavailable while the workbook is in an error state." : undefined}
            formulaAuditCallbacks={{
              onShowPrecedents: () => session.showFormulaPrecedents(),
              onShowDependents: () => session.showFormulaDependents(),
              onRemoveArrows: () => session.removeFormulaAuditArrows(),
              onSetShowFormulas: (enabled) => session.setShowFormulas(enabled),
              onScanErrors: () => session.scanFormulaErrors(),
              onEvaluateFormula: () => session.evaluateFormulaStep(),
              onRetry: () => session.retry(),
            }}
            definedNames={state.definedNameModels}
            onSaveDefinedName={(input) => session.setDefinedName(input)}
            onRemoveDefinedName={(input) => session.removeDefinedName(input.name, input.scope, input.sheetId)}
            sparklines={state.selectedSheet.sparklines}
            conditionalFormats={state.selectedSheet.conditionalFormats}
            dataValidations={state.selectedSheet.dataValidations}
            historyEntries={state.historyEntries}
            remoteRevisions={state.remoteRevisions}
            historyPreviewRevision={state.historyPreviewRevision}
            canRestoreHistory={state.permissions.restore}
            onUndoToHistory={session.undoToHistoryIndex.bind(session)}
            onRestoreRevision={(revision) => { void session.restoreToRevision(revision); }}
            onPreviewRevision={(revision) => { void session.previewRevision(revision); }}
            onClearHistoryPreview={session.clearHistoryPreview.bind(session)}
            onRefreshRevisions={() => { void session.refreshRevisionLog(); }}
            compatibilityReport={state.compatibilityReport}
            onClearCompatibilityReport={session.clearCompatibilityReport.bind(session)}
            tables={state.tables}
            onReadDataRows={session.readDataTable.bind(session)}
            onRemoveDataTable={session.removeDataTable.bind(session)}
            onCommand={dispatchCommand}
            onAddSparkline={session.addSparkline.bind(session)}
            onRemoveSparkline={session.removeSparkline.bind(session)}
            onAddConditionalFormat={session.addConditionalFormat.bind(session)}
            onRemoveConditionalFormat={session.removeConditionalFormat.bind(session)}
            onAddDataValidation={session.addDataValidation.bind(session)}
            onRemoveDataValidation={session.removeDataValidation.bind(session)}
            onPrint={session.printWorkbook.bind(session)}
            onExportPdf={session.exportPdf.bind(session)}
            printPageCount={state.printPageCount}
            queryConnectors={state.queryConnectors}
            loadedQueries={state.loadedQueries}
            lastQueryResult={state.lastQueryResult}
            canQuery={state.permissions.query}
            onLoadQuery={session.loadQuery.bind(session)}
            onRefreshQuery={session.refreshQuery.bind(session)}
            onTestQueryConnection={session.testQueryConnection.bind(session)}
            automationRecording={state.automationRecording}
            recordedScript={state.recordedScript}
            lastScriptResult={state.lastScriptResult}
            canRunScripts={state.permissions.script}
            onRunAutomationScript={session.runAutomationScript.bind(session)}
            onStartAutomationRecording={session.startAutomationRecording.bind(session)}
            onStopAutomationRecording={session.stopAutomationRecording.bind(session)}
            lastWhatIfMessage={
              state.lastWhatIfResult && 'message' in state.lastWhatIfResult
                ? state.lastWhatIfResult.message
                : state.lastWhatIfResult && 'status' in state.lastWhatIfResult
                  ? `${state.lastWhatIfResult.kind}: ${state.lastWhatIfResult.status}`
                  : null
            }
            canRunExtended={state.permissions.script}
            onGoalSeek={(params) => {
              session.runGoalSeek({
                setCell: { row: params.setRow, column: params.setColumn },
                toValue: params.targetValue,
                byChangingCell: { row: params.changingRow, column: params.changingColumn },
              });
            }}
            onRunDataTable={(params) => {
              session.runDataTableAnalysis({
                tableRange: params.tableRange,
                ...(params.inputMode === 'column'
                  ? { columnInputCell: params.inputCell }
                  : { rowInputCell: params.inputCell }),
              });
            }}
            onRunScenario={(params) => {
              session.runScenarioAnalysis({
                id: `scenario-${Date.now()}`,
                name: params.name,
                changingCells: [{
                  row: params.changingCell.row,
                  column: params.changingCell.column,
                  value: params.changingValue,
                }],
                resultCells: [{ row: params.resultCell.row, column: params.resultCell.column }],
              });
            }}
            onAddComment={session.addComment.bind(session)}
            onReplyComment={session.replyComment.bind(session)}
            onResolveComment={session.resolveComment.bind(session)}
            onRemoveComment={session.removeComment.bind(session)}
            onAddNote={session.addNote.bind(session)}
            onRemoveNote={session.removeNote.bind(session)}
            onSetHyperlink={session.setHyperlink.bind(session)}
            onRemoveHyperlink={session.removeHyperlink.bind(session)}
          />
          </Suspense>
        </SidebarShell>
        </Inline>
      </AppShell>

      <Suspense fallback={null}>
      <FunctionWizardDialog
        open={state.showFunctionWizard}
        onClose={session.closeFunctionWizard.bind(session)}
        onInsertFormula={(formula) => {
          session.setFormulaDraft(formula);
          session.commitFormula(formula);
        }}
      />

      <SortDialog
        open={state.showSortDialog}
        columns={sortColumns}
        locale={locale}
        onClose={session.closeSortDialog.bind(session)}
        onSort={(criteria, hasHeader) => session.sortRange(criteria, hasHeader)}
      />

      <FindReplaceDialog
        open={state.showFindReplace}
        initialFind={state.findQuery}
        locale={locale}
        onClose={session.closeFindReplace.bind(session)}
        onReplaceAll={(params) => session.findReplace(params)}
      />

      <GoToDialog
        open={state.showGoTo}
        onClose={session.closeGoTo.bind(session)}
        onGoTo={(reference) => session.selectAddress(reference)}
        onGoToSpecial={(kind) => session.goToSpecial(kind)}
      />

      <PasteSpecialDialog
        open={state.showPasteSpecial}
        locale={locale}
        onClose={session.closePasteSpecial.bind(session)}
        onPaste={(mode) => session.pasteSpecial(mode)}
      />

      <FormatCellsDialog
        open={state.showFormatCells}
        initial={formatCellsInitial}
        locale={locale}
        onClose={session.closeFormatCells.bind(session)}
        onApply={(draft) => session.formatCells({ numberFormat: draft.numberFormat, style: draft.style })}
      />

      <ShiftCellsDialog
        open={state.showShiftCells}
        locale={locale}
        onClose={session.closeShiftCells.bind(session)}
        onShift={(direction) => session.shiftCells(direction)}
      />

      <CreatePivotTableDialog
        open={state.showCreatePivotDialog}
        sourceRegion={session.getCurrentRegion()}
        sourceOptions={pivotSourceOptions.map(({ id, label }) => ({ id, label }))}
        activeSheetName={state.selectedSheet.name}
        onClose={session.closeCreatePivotDialog.bind(session)}
        onCreate={createPivotFromDialog}
      />

      <PrintPreviewDialog
        open={state.showPrintPreview}
        onClose={() => session.setShowPrintPreview(false)}
        sheetId={state.activeSheetId}
        rowCount={state.selectedSheet.rowCount}
        columnCount={state.selectedSheet.columnCount}
        columns={state.selectedSheet.columns}
        rows={state.selectedSheet.previewRows}
        layout={state.printLayout}
        pages={state.printPages}
        getRow={(row) => state.selectedSheet.previewRows[row] ?? {
          rowNumber: row + 1,
          cells: Array.from({ length: state.selectedSheet.columnCount }, (_, column) => ({
            value: state.selectedSheet.getCell(row, column)?.value ?? "",
          })),
        }}
      />
      </Suspense>
    </>
  );
}

export default function App() {
  if (typeof window !== "undefined" && window.location.pathname === "/workbooks") {
    return <Suspense fallback={<Box as="main" className="min-h-screen bg-canvas" />}><WorkbookCatalog /></Suspense>;
  }
  return <WorkspaceErrorBoundary><WorkspaceApp /></WorkspaceErrorBoundary>;
}
