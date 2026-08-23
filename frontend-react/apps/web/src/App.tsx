import { AppShell, Box, Button, DropdownMenu, Inline, SidebarShell, Stack } from "@react-sheets/ui-system";
import { FormulaBar } from "./components/FormulaBar";
import { Ribbon } from "./components/Ribbon";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { WorkspaceErrorBoundary } from "./components/WorkspaceErrorBoundary";
import { parseRangeInput } from "./domain/range-input";
import type { CommandDescriptor } from "./domain/command-descriptor";
import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";
import {
  getInitialAppPhase,
  useSpreadsheetApp,
  type SelectionState,
  type SidebarPanelId,
  type UiSessionIntent,
} from "@react-sheets/spreadsheet-app";
import { getInitialLocale, localeLabels, persistLocale, shellLabels, type Locale } from "./i18n";
import zhCN from "./locales/zh-CN.json";
import enUS from "./locales/en-US.json";
import type { ChartDrawingPayload, DrawingObject, PivotAggregateFunction, PivotFieldDefinition, PivotLayout, PivotModel, ShapeDrawingPayload, SparklineModel } from "@react-sheets/core-model";
import type { PivotPanelCallbacks, PivotPanelResult, PivotPanelState } from "./components/pivot/pivot-contract";

const FeatureSidebar = lazy(() => import("./components/FeatureSidebar").then((module) => ({ default: module.FeatureSidebar })));
const FunctionWizardDialog = lazy(() => import("./components/dialogs/FunctionWizardDialog").then((module) => ({ default: module.FunctionWizardDialog })));
const SortDialog = lazy(() => import("./components/dialogs/SortDialog").then((module) => ({ default: module.SortDialog })));
const FindReplaceDialog = lazy(() => import("./components/dialogs/FindReplaceDialog").then((module) => ({ default: module.FindReplaceDialog })));
const GoToDialog = lazy(() => import("./components/dialogs/GoToDialog").then((module) => ({ default: module.GoToDialog })));
const PasteSpecialDialog = lazy(() => import("./components/dialogs/PasteSpecialDialog").then((module) => ({ default: module.PasteSpecialDialog })));
const FormatCellsDialog = lazy(() => import("./components/dialogs/FormatCellsDialog").then((module) => ({ default: module.FormatCellsDialog })));
const ShiftCellsDialog = lazy(() => import("./components/dialogs/ShiftCellsDialog").then((module) => ({ default: module.ShiftCellsDialog })));
const PrintPreviewDialog = lazy(() => import("./components/dialogs/PrintPreviewDialog").then((module) => ({ default: module.PrintPreviewDialog })));
const WorkbookCatalog = lazy(() => import("./components/WorkbookCatalog").then((module) => ({ default: module.WorkbookCatalog })));
const SheetCanvas = lazy(() => import("./components/SheetCanvas").then((module) => ({ default: module.SheetCanvas })));

function WorkspaceApp() {
  const { app, snapshot: state } = useSpreadsheetApp({ initialPhase: getInitialAppPhase() });
  const [locale, setLocaleState] = useState<Locale>(() => getInitialLocale());
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const previousPanelRef = useRef(state.activePanel);
  const isBusy = state.phase !== "ready";
  const setLocale = (nextLocale: Locale) => {
    setLocaleState(nextLocale);
    persistLocale(nextLocale);
  };

  const executeCommand = (descriptor: CommandDescriptor) => {
    app.execute(descriptor.commandId, descriptor.params);
  };

  const dispatchSessionIntent = (intent: UiSessionIntent) => {
    app.dispatchUiSessionIntent(intent);
  };

  useEffect(() => {
    if (previousPanelRef.current !== state.activePanel) {
      setSidebarOpen(true);
    }
    previousPanelRef.current = state.activePanel;
  }, [state.activePanel]);

  const copyWorkbookLink = () => { void app.createGuestShareLink('editor'); };

  const saveWorkbook = () => { void app.saveWorkbook("Ribbon save"); };

  const exportXlsx = async () => {
    const exported = await app.exportXlsxWorkbook();
    if (!exported) return;
    const binary = window.atob(exported.base64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    const href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
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
      void file.arrayBuffer().then((buffer) => {
        const bytes = new Uint8Array(buffer);
        let binary = "";
        for (const byte of bytes) binary += String.fromCharCode(byte);
        return app.importXlsxBase64(window.btoa(binary), file.name);
      });
    };
    input.click();
  };

  const selectedRange = state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0];
  const pivotSourceRange = selectedRange && (selectedRange.endRow > selectedRange.startRow || selectedRange.endColumn > selectedRange.startColumn)
    ? selectedRange
    : state.selectedSheet.usedRange;
  const createWebId = (prefix: string): string => {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return `${prefix}-${crypto.randomUUID()}`;
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  };
  const buildQuickPivotCommand = (): CommandDescriptor | undefined => {
    const fields = app.getPivotFieldCatalog(pivotSourceRange);
    const rowField = fields.find((field) => field.dataType !== "number")?.name ?? fields[0]?.name;
    const valueField = fields.find((field) => field.dataType === "number")?.name ?? fields[0]?.name;
    if (!rowField || !valueField) return undefined;
    const pivotId = createWebId("pivot");
    setActivePivotId(pivotId);
    const pivot: PivotModel = {
      id: pivotId,
      sheetId: state.activeSheetId,
      sourceRange: { ...pivotSourceRange, sheetId: state.activeSheetId },
      refreshPolicy: { mode: "on-change", preserveFormatting: true, refreshOnLoad: true },
      layout: {
        rows: [{ field: rowField }],
        columns: [],
        filters: [],
        values: [{ field: valueField, summarizeBy: fields.find((field) => field.name === valueField)?.dataType === "number" ? "sum" : "count" }],
        calculatedFields: [],
        calculatedItems: [],
        showSubtotals: true,
        showGrandTotals: true,
        compact: true,
        repeatLabels: false,
      },
    };
    return { commandId: "pivot.add", params: pivot };
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
  const selectedDrawing = state.selectedFloatingId ? state.selectedSheet.drawings.find((drawing) => drawing.payloadId === state.selectedFloatingId) : undefined;
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
    commandId: "sheet.filter.set",
    params: { sheetId: state.activeSheetId, filter: { sheetId: state.activeSheetId, range: state.selectedSheet.usedRange, criteria: {} } },
  });
  const buildClearFilterCommand = (): CommandDescriptor => ({ commandId: "sheet.filter.remove", params: { sheetId: state.activeSheetId } });
  const [activePivotId, setActivePivotId] = useState<string>();
  const activePivot = state.selectedSheet.pivots.find((pivot) => pivot.id === activePivotId) ?? state.selectedSheet.pivots[0];
  const pivotTree = activePivot ? state.selectedSheet.pivotResults[activePivot.id] : undefined;
  const corePivotFields = pivotTree?.fields.fields ?? app.getPivotFieldCatalog(pivotSourceRange);
  const pivotFields: PivotFieldDefinition[] = corePivotFields;
  const pivotResult: PivotPanelResult | undefined = pivotTree
    ? { rowCount: pivotTree.rows.length, columnCount: pivotTree.columnPaths.length, tree: pivotTree, summary: `${pivotTree.rows.length} row groups × ${pivotTree.columnPaths.length || 1} column groups` }
    : undefined;

  const cloneLayout = (layout: PivotLayout): PivotLayout => structuredClone(layout);
  const removeField = (layout: PivotLayout, fieldId: string): PivotLayout => {
    const next = cloneLayout(layout);
    next.filters = next.filters.filter((filter) => filter.field !== fieldId);
    next.rows = next.rows.filter((field) => field.field !== fieldId);
    next.columns = next.columns.filter((field) => field.field !== fieldId);
    next.values = next.values.filter((value) => value.field !== fieldId);
    return next;
  };
  const updatePivotLayout = (nextLayout: PivotLayout) => {
    if (!activePivot) return;
    executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, layout: nextLayout } });
    app.notify("Pivot layout updated");
  };

  const createPivot = () => {
    const id = app.insertQuickPivot();
    if (id) setActivePivotId(id);
  };

  const pivotCallbacks: PivotPanelCallbacks = {
    onCreate: createPivot,
    onPivotSelect: setActivePivotId,
    onFieldAreaChange: (fieldId, area, index) => {
      if (!activePivot) return;
      const next = removeField(activePivot.layout, fieldId);
      if (area === "values") {
        const field = pivotFields.find((entry) => entry.id === fieldId);
        const summarizeBy: PivotAggregateFunction = field?.dataType === "number" ? "sum" : "count";
        next.values.splice(Math.max(0, index), 0, { field: fieldId, summarizeBy });
      } else if (area === "filters") {
        next.filters.splice(Math.max(0, index), 0, { kind: "manual", field: fieldId, selected: [] });
      } else {
        next[area].splice(Math.max(0, index), 0, { field: fieldId });
      }
      updatePivotLayout(next);
    },
    onRemoveField: (fieldId, area) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      if (area === "values") next.values = next.values.filter((value) => value.field !== fieldId);
      else if (area === "filters") next.filters = next.filters.filter((filter) => filter.field !== fieldId);
      else next[area] = next[area].filter((field) => field.field !== fieldId);
      updatePivotLayout(next);
    },
    onValueChange: (value) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const index = next.values.findIndex((entry) => entry.field === value.field);
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
    onFilterChange: (fieldId, selectedValues) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const existing = next.filters.find((filter) => filter.field === fieldId);
      if (existing?.kind === "manual") existing.selected = [...selectedValues];
      else next.filters.push({ kind: "manual", field: fieldId, selected: [...selectedValues] });
      updatePivotLayout(next);
      app.notify("Pivot filter updated");
    },
    onSortChange: (fieldId, direction) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      next.rows = next.rows.map((field) => field.field === fieldId ? { ...field, sort: direction === "none" ? undefined : { direction } } : field);
      next.columns = next.columns.map((field) => field.field === fieldId ? { ...field, sort: direction === "none" ? undefined : { direction } } : field);
      updatePivotLayout(next);
    },
    onGroupChange: (fieldId, grouped) => {
      if (!activePivot) return;
      const field = pivotFields.find((entry) => entry.id === fieldId);
      const group = !grouped ? undefined : field?.dataType === "date"
        ? { kind: "date" as const, unit: "month" as const }
        : field?.dataType === "number"
          ? { kind: "number" as const, interval: 10 }
          : undefined;
      const next = cloneLayout(activePivot.layout);
      next.rows = next.rows.map((entry) => entry.field === fieldId ? { ...entry, group } : entry);
      next.columns = next.columns.map((entry) => entry.field === fieldId ? { ...entry, group } : entry);
      updatePivotLayout(next);
    },
    onRefresh: () => {
      if (!activePivot) return;
      executeCommand({ commandId: "pivot.refresh", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, refreshRevision: (activePivot.refreshRevision ?? 0) + 1, lastRefreshedAt: new Date().toISOString() } });
    },
    onSourceRangeChange: (sourceRange) => {
      if (!activePivot) return;
      const parsed = parseRangeInput(sourceRange, activePivot.sheetId);
      if (!parsed) {
        app.notify("Invalid pivot source range");
        return;
      }
      executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, sourceRange: { sheetId: activePivot.sheetId, ...parsed } } });
    },
    onLayoutChange: (layout) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      next.compact = layout === "compact";
      next.repeatLabels = layout === "tabular";
      updatePivotLayout(next);
    },
    onExpandedChange: (fieldId, expanded) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const expandedFieldIds = new Set(next.expandedFieldIds ?? next.rows.map((field) => field.field));
      if (expanded) expandedFieldIds.add(fieldId);
      else expandedFieldIds.delete(fieldId);
      next.expandedFieldIds = [...expandedFieldIds];
      updatePivotLayout(next);
    },
    onSlicerChange: (fieldId, enabled) => {
      if (!activePivot) return;
      const connectedPivotIds = app.getConnectedPivotIds(activePivot.sourceRange);
      if (enabled) {
        executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, slicers: [...(activePivot.slicers ?? []), { id: `slicer-${fieldId}`, field: fieldId, selected: [], connectedPivotIds }] } });
        return;
      }
      executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, slicers: (activePivot.slicers ?? []).filter((slicer) => slicer.field !== fieldId) } });
    },
    onTimelineChange: (fieldId) => {
      if (!activePivot) return;
      const connectedPivotIds = app.getConnectedPivotIds(activePivot.sourceRange);
      if (!fieldId) {
        executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, timelines: [] } });
        return;
      }
      executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, timelines: [{ id: `timeline-${fieldId}`, field: fieldId, connectedPivotIds }] } });
    },
    onTimelineRangeChange: (start, end) => {
      if (!activePivot) return;
      const timelineFieldId = activePivot.timelines?.[0]?.field;
      if (!timelineFieldId) return;
      const connectedPivotIds = app.getConnectedPivotIds(activePivot.sourceRange);
      executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, timelines: [{ id: `timeline-${timelineFieldId}`, field: timelineFieldId, start: start || undefined, end: end || undefined, connectedPivotIds }] } });
    },
    onPivotChartChange: (chart) => {
      if (!activePivot || !chart) return;
      const chartId = `pivot-chart-${activePivot.id}-${Date.now().toString(36)}`;
      const drawing: DrawingObject = {
        id: `drawing-${chartId}`,
        sheetId: activePivot.sheetId,
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
        sourceRanges: [activePivot.sourceRange],
      };
      executeCommand({ commandId: "chart.insert", params: { sheetId: activePivot.sheetId, drawing, payload } });
      executeCommand({ commandId: "pivot.update", params: { sheetId: activePivot.sheetId, pivotId: activePivot.id, chartReferences: [...(activePivot.chartReferences ?? []), { chartId, role: "linked" }] } });
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
    if (panel === "inspector" || panel === "chart" || panel === "pivot" || panel === "shape" || panel === "sparkline" || panel === "conditionalFormat" || panel === "dataValidation" || panel === "history" || panel === "data") {
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
    const range = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
    if (!range) return;
    app.selectRange(
      { startRow: range.startRow, endRow: range.endRow, startColumn: range.startColumn, endColumn: range.endColumn },
      selection.ranges.length > 1 ? "add" : "replace",
    );
  };

  const formatCellsInitial = useMemo(() => {
    const cell = state.selectedSheet.getCell(state.selection.primaryRowIndex, state.selection.primaryColumnIndex);
    const style = cell?.style ?? {};
    return {
      numberFormat: style.numberFormat ?? "general",
      style: { ...style },
    };
  }, [state.showFormatCells, state.selection.primaryRowIndex, state.selection.primaryColumnIndex, state.selectedSheet, state.version]);
  const selectedCellStyle = state.selectedSheet.getCell(state.selection.primaryRowIndex, state.selection.primaryColumnIndex)?.style ?? {};

  return (
    <>
      <AppShell
        formulaBar={
          <FormulaBar
            cellName={state.activeCell}
            disabled={isBusy}
            formula={state.formulaDraft}
            locale={locale}
            onCancel={app.cancelEdit}
            onChange={app.setFormulaDraft}
            onCommit={() => {
              if (state.editingCell) app.commitEdit("down");
              else app.commitFormula();
            }}
            onNameBoxCommit={(value) => app.selectAddress(value)}
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
                  if (nextName?.trim()) app.renameWorkbook(nextName.trim());
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
            locale={locale}
            onCommand={executeCommand}
            onSessionIntent={dispatchSessionIntent}
            onCopy={() => app.copy()}
            onCut={() => app.cut()}
            onPaste={() => app.paste()}
            onUndo={() => app.undo()}
            onRedo={() => app.redo()}
            onSave={saveWorkbook}
            onExportXlsx={() => void exportXlsx()}
            onImportXlsx={importXlsx}
            onRecalculate={() => app.recalculateFormulas()}
            onAutoSum={() => dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" })}
            onFreezeAtPrimary={() => app.freezeAtPrimary()}
            onCreatePivot={buildQuickPivotCommand}
            onCreateChart={buildQuickChartCommand}
            onCreateSparkline={buildQuickSparklineCommand}
            onCreateShape={buildQuickShapeCommand}
            onBringDrawingForward={() => buildDrawingCommand("drawing.zorder", "forward")}
            onSendDrawingBackward={() => buildDrawingCommand("drawing.zorder", "backward")}
            onRemoveDrawing={() => buildDrawingCommand("drawing.remove")}
            onCreateSheetTable={() => app.createSheetTableFromSelection()}
            onCreateDataTable={() => app.createDataTableFromSelection()}
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
            onTabChange={(tab) => app.setRibbonTab(tab)}
            phase={state.phase}
            canExecute={app.canExecute.bind(app)}
          />
        }
        saveState={state.saveState}
        sheetTabs={
          <SheetTabs
            activeSheetId={state.activeSheetId}
            locale={locale}
            disabled={isBusy}
            onAdd={app.addSheet}
            onSelect={app.selectSheet}
            onRenameSheet={app.renameSheet}
            onDeleteSheet={app.deleteSheet}
            onDuplicateSheet={app.duplicateSheet}
            onHideSheet={app.hideSheet}
            onSetTabColor={app.setSheetTabColor}
            onMoveSheet={app.moveSheet}
            sheets={state.sheets}
          />
        }
        statusBar={
          <StatusBar
            activeCell={state.activeCell}
            locale={locale}
            onOpenShortcuts={() => app.notify("Shortcuts: Arrows / Tab / Enter / F2 / F4 / Ctrl+C/X/V/Z/Y/B/I/U")}
            onZoomChange={app.setZoom}
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
              drawings={state.selectedSheet.drawings}
              drawingPayloads={state.selectedSheet.drawingPayloads}
              allSheets={state.sheets}
              pivotResults={state.selectedSheet.pivotResults}
              sparklines={state.selectedSheet.sparklines}
              onSelectionChange={applySelection}
              onExtendSelection={(row, column) => app.extendSelectionTo(row, column)}
              onMovePrimary={(rowDelta, columnDelta, opts) => app.movePrimary(rowDelta, columnDelta, opts)}
              onCommitCell={(value) => app.commitFormula(value)}
              onBeginEdit={(initialText) => app.beginEdit(initialText)}
              onCancelEdit={app.cancelEdit}
              onCommitEdit={(moveAfter) => app.commitEdit(moveAfter ?? "down")}
              onFormulaDraftChange={app.setFormulaDraft}
              onAppendFormulaDraft={app.appendFormulaDraft.bind(app)}
              onInsertRef={app.insertRefIntoDraft}
              onToggleAbsolute={app.toggleAbsoluteReference}
              onJumpEdge={(direction, extend) => app.jumpEdge(direction, extend)}
              onSelectAll={app.selectAll}
              onResizeRow={app.resizeRow}
              onResizeColumn={app.resizeColumn}
              onFillRange={app.fillRange}
              onFloatingSelect={(hit) => app.setSelectedFloatingId(hit ? hit.id : null)}
              onFloatingMove={(drawingId, bounds, rotation) => executeCommand({
                commandId: "drawing.move",
                params: { sheetId: state.activeSheetId, drawingId, transform: { ...bounds, rotation } },
              })}
              onFloatingRemove={(drawingId) => executeCommand({
                commandId: "drawing.remove",
                params: { sheetId: state.activeSheetId, drawingId },
              })}
              onCommand={executeCommand}
              onCopy={() => app.copy()}
              onCut={() => app.cut()}
              onPaste={() => app.paste()}
              onUndo={() => app.undo()}
              onRedo={() => app.redo()}
              onOpenInspector={() => dispatchSessionIntent({ type: "panel.open", panel: "inspector", notice: "Select a cell and use Review tools for comments." })}
              onApplyFilter={(column, patch) => app.applyFilter(column, patch)}
              onToggleOutline={(groupId) => app.toggleOutlineGroup(groupId)}
              getValidationList={app.getValidationAt}
              onRetry={app.retry}
              onCreateSheet={app.addSheet}
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
            onRetry={app.retry}
            phase={state.phase}
            sheet={state.selectedSheet}
            sheetId={state.activeSheetId}
            drawings={state.selectedSheet.drawings}
            drawingPayloads={state.selectedSheet.drawingPayloads}
            pivot={activePivot}
            pivotList={state.selectedSheet.pivots.map((pivot) => ({ id: pivot.id, label: pivot.id }))}
            activePivotId={activePivot?.id}
            pivotFieldCatalog={pivotFields}
            pivotResult={pivotResult}
            onShowPivotDetails={(paths) => {
              if (activePivot) app.showPivotDetails(activePivot.id, paths);
            }}
            pivotPanelState={pivotPanelState}
            pivotCallbacks={pivotCallbacks}
            sparklines={state.selectedSheet.sparklines}
            conditionalFormats={state.selectedSheet.conditionalFormats}
            dataValidations={state.selectedSheet.dataValidations}
            historyEntries={state.historyEntries}
            remoteRevisions={state.remoteRevisions}
            historyPreviewRevision={state.historyPreviewRevision}
            canRestoreHistory={state.permissions.restore}
            onUndoToHistory={app.undoToHistoryIndex.bind(app)}
            onRestoreRevision={(revision) => { void app.restoreToRevision(revision); }}
            onPreviewRevision={(revision) => { void app.previewRevision(revision); }}
            onClearHistoryPreview={app.clearHistoryPreview.bind(app)}
            onRefreshRevisions={() => { void app.refreshRevisionLog(); }}
            compatibilityReport={state.compatibilityReport}
            onClearCompatibilityReport={app.clearCompatibilityReport.bind(app)}
            tables={state.tables}
            onReadDataRows={app.readDataTable.bind(app)}
            onRemoveDataTable={app.removeDataTable.bind(app)}
            onCommand={executeCommand}
            onAddSparkline={app.addSparkline}
            onRemoveSparkline={app.removeSparkline}
            onAddConditionalFormat={app.addConditionalFormat}
            onRemoveConditionalFormat={app.removeConditionalFormat}
            onAddDataValidation={app.addDataValidation}
            onRemoveDataValidation={app.removeDataValidation}
            onPrint={app.printWorkbook}
            onExportPdf={app.exportPdf}
            printPageCount={state.printPageCount}
            queryConnectors={state.queryConnectors}
            loadedQueries={state.loadedQueries}
            lastQueryResult={state.lastQueryResult}
            canQuery={state.permissions.query}
            onLoadQuery={app.loadQuery.bind(app)}
            onRefreshQuery={app.refreshQuery.bind(app)}
            onTestQueryConnection={app.testQueryConnection.bind(app)}
            automationRecording={state.automationRecording}
            recordedScript={state.recordedScript}
            lastScriptResult={state.lastScriptResult}
            canRunScripts={state.permissions.script}
            onRunAutomationScript={app.runAutomationScript.bind(app)}
            onStartAutomationRecording={app.startAutomationRecording.bind(app)}
            onStopAutomationRecording={app.stopAutomationRecording.bind(app)}
            platformCapabilities={state.platformCapabilities}
            lastWhatIfMessage={
              state.lastWhatIfResult && 'message' in state.lastWhatIfResult
                ? state.lastWhatIfResult.message
                : state.lastWhatIfResult && 'status' in state.lastWhatIfResult
                  ? `${state.lastWhatIfResult.kind}: ${state.lastWhatIfResult.status}`
                  : null
            }
            canRunExtended={state.permissions.script}
            onGoalSeek={(params) => {
              app.runGoalSeek({
                setCell: { row: params.setRow, column: params.setColumn },
                toValue: params.targetValue,
                byChangingCell: { row: params.changingRow, column: params.changingColumn },
              });
            }}
            onRunDataTable={(params) => {
              app.runDataTableAnalysis({
                tableRange: params.tableRange,
                ...(params.inputMode === 'column'
                  ? { columnInputCell: params.inputCell }
                  : { rowInputCell: params.inputCell }),
              });
            }}
            onRunScenario={(params) => {
              app.runScenarioAnalysis({
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
            onEvaluateCapability={async (capability) => {
              const result = app.evaluatePlatformCapability(capability as import('@react-sheets/spreadsheet-app').PlatformCapability);
              return { ok: result.canEnable, message: result.reason };
            }}
            onAddComment={app.addComment}
            onReplyComment={app.replyComment}
            onResolveComment={app.resolveComment}
            onRemoveComment={app.removeComment}
            onAddNote={app.addNote}
            onRemoveNote={app.removeNote}
            onSetHyperlink={app.setHyperlink}
            onRemoveHyperlink={app.removeHyperlink}
          />
          </Suspense>
        </SidebarShell>
        </Inline>
      </AppShell>

      <Suspense fallback={null}>
      <FunctionWizardDialog
        open={state.showFunctionWizard}
        onClose={app.closeFunctionWizard}
        onInsertFormula={(formula) => {
          app.setFormulaDraft(formula);
          app.commitFormula(formula);
        }}
      />

      <SortDialog
        open={state.showSortDialog}
        columns={state.selectedSheet.columns}
        onClose={app.closeSortDialog}
        onSort={(criteria, hasHeader) => app.sortRange(criteria, hasHeader)}
      />

      <FindReplaceDialog
        open={state.showFindReplace}
        initialFind={state.findQuery}
        onClose={app.closeFindReplace}
        onReplaceAll={(params) => app.findReplace(params)}
      />

      <GoToDialog
        open={state.showGoTo}
        onClose={app.closeGoTo}
        onGoTo={(reference) => app.selectAddress(reference)}
        onGoToSpecial={(kind) => app.goToSpecial(kind)}
      />

      <PasteSpecialDialog
        open={state.showPasteSpecial}
        onClose={app.closePasteSpecial}
        onPaste={(mode) => app.pasteSpecial(mode)}
      />

      <FormatCellsDialog
        open={state.showFormatCells}
        initial={formatCellsInitial}
        onClose={app.closeFormatCells}
        onApply={(draft) => app.formatCells({ numberFormat: draft.numberFormat, style: draft.style })}
      />

      <ShiftCellsDialog
        open={state.showShiftCells}
        onClose={app.closeShiftCells}
        onShift={(direction) => app.shiftCells(direction)}
      />

      <PrintPreviewDialog
        open={state.showPrintPreview}
        onClose={() => app.setShowPrintPreview(false)}
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
