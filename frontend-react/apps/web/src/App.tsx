import { Box, Button, DropdownMenu, Inline, Stack } from "@react-sheets/ui-system";
import { AppShell } from "./components/AppShell";
import { FeatureSidebar } from "./components/FeatureSidebar";
import { FormulaBar } from "./components/FormulaBar";
import { Ribbon } from "./components/Ribbon";
import { SheetCanvas } from "./components/SheetCanvas";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { FunctionWizardDialog } from "./components/dialogs/FunctionWizardDialog";
import { SortDialog } from "./components/dialogs/SortDialog";
import { FindReplaceDialog } from "./components/dialogs/FindReplaceDialog";
import { PrintPreviewDialog } from "./components/dialogs/PrintPreviewDialog";
import { WorkbookCatalog } from "./components/WorkbookCatalog";
import { useState } from "react";
import { getInitialWorkspacePhase, useWorkspaceState, type SelectionState, type SidebarPanelId } from "./state/workspace";
import type { PivotAggregateFunction, PivotFieldDefinition as CorePivotFieldDefinition, PivotFilter, PivotLayout, PivotModel, PivotShowAs as CorePivotShowAs, PivotValueField } from "@react-sheets/core-model";
import type { PivotCalculatedFieldDefinition, PivotCalculatedItemDefinition, PivotDefinition, PivotFieldDefinition, PivotPanelCallbacks, PivotPanelState, PivotResult, PivotShowAs as UiPivotShowAs } from "./components/pivot/types";

const emptyPivotDefinition: PivotDefinition = {
  filters: [], columns: [], rows: [], values: [], calculatedFields: [], calculatedItems: [], filterSelections: {}, sort: {}, groupedFields: [], layout: "compact", showGrandTotals: true, showSubtotals: true, expandedFieldIds: [], slicers: [],
};

function toUiShowAs(showAs?: CorePivotShowAs): UiPivotShowAs {
  if (!showAs || showAs.kind === "normal") return "normal";
  if (showAs.kind === "grand-percentage") return "percent-of-total";
  if (showAs.kind === "row-percentage") return "percent-of-row";
  if (showAs.kind === "column-percentage") return "percent-of-column";
  if (showAs.kind === "parent-percentage") return "percent-of-parent";
  if (showAs.kind === "percentage-difference") return "percent-difference-from";
  return showAs.kind === "difference" ? "difference-from" : showAs.kind;
}

function fromUiShowAs(showAs: UiPivotShowAs): CorePivotShowAs {
  if (showAs === "percent-of-total") return { kind: "grand-percentage" };
  if (showAs === "percent-of-row") return { kind: "row-percentage" };
  if (showAs === "percent-of-column") return { kind: "column-percentage" };
  if (showAs === "percent-of-parent") return { kind: "parent-percentage" };
  if (showAs === "difference-from") return { kind: "difference", base: "grand" };
  if (showAs === "percent-difference-from") return { kind: "percentage-difference", base: "grand" };
  if (showAs === "running-total") return { kind: "running-total", axis: "row" };
  if (showAs === "rank") return { kind: "rank", axis: "row", direction: "descending" };
  if (showAs === "index") return { kind: "index" };
  return { kind: "normal" };
}

function pivotToUiDefinition(pivot?: PivotModel): PivotDefinition {
  if (!pivot) return emptyPivotDefinition;
  const layout = pivot.layout;
  const rows = layout?.rows.map((field) => field.field) ?? pivot.rowFields;
  const columns = layout?.columns.map((field) => field.field) ?? pivot.columnFields;
  const filters = layout?.filters.map((filter) => filter.field) ?? pivot.filterFields;
  const values = (layout?.values ?? pivot.valueFields).map((value, index) => ({
    id: `${value.field}-${index}`,
    fieldId: value.field,
    summary: value.summarizeBy,
    displayName: value.displayName ?? `${value.summarizeBy.toUpperCase()} of ${value.field}`,
    numberFormat: "",
    showAs: toUiShowAs(value.showAs),
  }));
  const filterSelections: Record<string, string[]> = {};
  for (const filter of layout?.filters ?? []) if (filter.kind === "manual") filterSelections[filter.field] = filter.selected.map((value) => String(value));
  const sort: Record<string, "none" | "ascending" | "descending"> = {};
  for (const field of [...(layout?.rows ?? []), ...(layout?.columns ?? [])]) sort[field.field] = field.sort?.direction ?? "none";
  return {
    filters,
    columns,
    rows,
    values,
    calculatedFields: (layout?.calculatedFields ?? []).map((field) => ({ ...field })),
    calculatedItems: (layout?.calculatedItems ?? []).map((item) => ({ fieldId: item.field, name: item.name, formula: item.formula })),
    filterSelections,
    sort,
    groupedFields: [...(layout?.rows ?? []), ...(layout?.columns ?? [])].filter((field) => Boolean(field.group)).map((field) => field.field),
    layout: layout?.compact ? "compact" : "outline",
    showGrandTotals: layout?.showGrandTotals ?? true,
    showSubtotals: layout?.showSubtotals ?? true,
    expandedFieldIds: layout?.expandedFieldIds ?? [...rows],
    slicers: pivot.slicers?.map((slicer) => slicer.field) ?? [],
    timelineFieldId: pivot.timelines?.[0]?.field,
    timelineStart: pivot.timelines?.[0]?.start,
    timelineEnd: pivot.timelines?.[0]?.end,
  };
}

function uiToPivotLayout(definition: PivotDefinition, fields: readonly PivotFieldDefinition[]): PivotLayout {
  const placement = (field: string) => ({
    field,
    sort: definition.sort[field] && definition.sort[field] !== "none" ? { direction: definition.sort[field] as "ascending" | "descending" } : undefined,
    group: definition.groupedFields.includes(field)
      ? fields.find((candidate) => candidate.id === field)?.type === "date" ? { kind: "date" as const, unit: "month" as const }
        : fields.find((candidate) => candidate.id === field)?.type === "number" ? { kind: "number" as const, interval: 10 }
          : undefined
      : undefined,
  });
  const filters: PivotFilter[] = [...new Set([...definition.filters, ...definition.slicers])].flatMap((field) => {
    const selected = definition.filterSelections[field] ?? [];
    return selected.length ? [{ kind: "manual", field, selected: selected.slice() }] : [];
  });
  const values: PivotValueField[] = definition.values.map((value) => ({ field: value.fieldId, summarizeBy: value.summary as PivotAggregateFunction, displayName: value.displayName, showAs: fromUiShowAs(value.showAs) }));
  return {
    rows: definition.rows.map(placement),
    columns: definition.columns.map(placement),
    filters,
    values,
    calculatedFields: definition.calculatedFields.map((field) => ({ ...field })),
    calculatedItems: definition.calculatedItems.map((item) => ({ field: item.fieldId, name: item.name, formula: item.formula })),
    showSubtotals: definition.showSubtotals,
    showGrandTotals: definition.showGrandTotals,
    compact: definition.layout === "compact",
    repeatLabels: definition.layout === "tabular",
    expandedFieldIds: [...definition.expandedFieldIds],
  };
}

function WorkspaceApp() {
  const { actions, state } = useWorkspaceState({ initialPhase: getInitialWorkspacePhase() });
  const isBusy = state.phase !== "ready";
  const copyWorkbookLink = () => {
    const link = `${window.location.origin}/workbooks/${encodeURIComponent(state.unitId)}`;
    const clipboard = navigator.clipboard;
    if (!clipboard) {
      actions.notify("Clipboard access is unavailable");
      return;
    }
    void clipboard.writeText(link)
      .then(() => actions.notify("Workbook link copied"))
      .catch(() => actions.notify("Could not copy workbook link"));
  };

  const selectedRange = state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0];
  const pivotSourceRange = selectedRange && (selectedRange.endRow > selectedRange.startRow || selectedRange.endColumn > selectedRange.startColumn)
    ? selectedRange
    : { sheetId: state.activeSheetId, startRow: 0, endRow: Math.max(1, state.selectedSheet.rows.length - 1), startColumn: 0, endColumn: Math.max(1, state.selectedSheet.columns.length - 1) };
  const [activePivotId, setActivePivotId] = useState<string>();
  const activePivot = state.selectedSheet.pivots.find((pivot) => pivot.id === activePivotId) ?? state.selectedSheet.pivots[0];
  const pivotTree = activePivot ? state.selectedSheet.pivotResults[activePivot.id] : undefined;
  const corePivotFields = pivotTree?.fields.fields ?? actions.getPivotFieldCatalog(pivotSourceRange);
  const pivotFields: PivotFieldDefinition[] = corePivotFields.map((field: CorePivotFieldDefinition) => ({ id: field.id, label: field.name, type: field.dataType === "mixed" ? "text" : field.dataType, values: field.values?.map((value) => String(value)) }));
  const pivotDefinition = pivotToUiDefinition(activePivot);
  const pivotResult: PivotResult | undefined = pivotTree ? { rowCount: pivotTree.rows.length, columnCount: pivotTree.columnPaths.length, tree: pivotTree, summary: `${pivotTree.rows.length} row groups × ${pivotTree.columnPaths.length || 1} column groups` } : undefined;
  const updatePivotDefinition = (next: PivotDefinition) => {
    if (activePivot) {
      actions.updatePivotLayout(activePivot.id, uiToPivotLayout(next, pivotFields));
      actions.notify("Pivot layout updated");
    }
  };
  const clonePivotDefinition = (source: PivotDefinition): PivotDefinition => ({
    ...source,
    filters: [...source.filters],
    columns: [...source.columns],
    rows: [...source.rows],
    values: source.values.map((value) => ({ ...value })),
    calculatedFields: source.calculatedFields.map((field) => ({ ...field })),
    calculatedItems: source.calculatedItems.map((item) => ({ ...item })),
    filterSelections: Object.fromEntries(Object.entries(source.filterSelections).map(([key, values]) => [key, [...values]])),
    sort: { ...source.sort },
    groupedFields: [...source.groupedFields],
    expandedFieldIds: [...source.expandedFieldIds],
    slicers: [...source.slicers],
  });
  const createPivot = () => {
    const rowField = pivotFields.find((field) => field.type !== "number")?.id ?? pivotFields[0]?.id;
    const valueField = pivotFields.find((field) => field.type === "number")?.id ?? pivotFields[0]?.id;
    if (!rowField || !valueField) return;
    const id = `pivot-${Math.random().toString(36).slice(2, 8)}`;
    const summarizeBy = pivotFields.find((field) => field.id === valueField)?.type === "number" ? "sum" : "count";
    actions.addPivot({ id, sheetId: state.activeSheetId, sourceRange: pivotSourceRange, rowFields: [rowField], columnFields: [], valueFields: [{ field: valueField, summarizeBy }], filterFields: [], layout: { rows: [{ field: rowField }], columns: [], filters: [], values: [{ field: valueField, summarizeBy }], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false, calculatedFields: [], calculatedItems: [] } });
    setActivePivotId(id);
  };
  const pivotCallbacks: PivotPanelCallbacks = {
    onCreate: createPivot,
    onPivotSelect: setActivePivotId,
    onDefinitionChange: updatePivotDefinition,
    onFieldAreaChange: (fieldId, area, index) => {
      const next = clonePivotDefinition(pivotDefinition);
      next.filters = next.filters.filter((field) => field !== fieldId);
      next.columns = next.columns.filter((field) => field !== fieldId);
      next.rows = next.rows.filter((field) => field !== fieldId);
      next.values = next.values.filter((value) => value.fieldId !== fieldId);
      if (area === "values") next.values.splice(Math.max(0, index), 0, { id: `${fieldId}-${Date.now()}`, fieldId, summary: pivotFields.find((field) => field.id === fieldId)?.type === "number" ? "sum" : "count", displayName: fieldId, numberFormat: "", showAs: "normal" });
      else next[area].splice(Math.max(0, index), 0, fieldId);
      updatePivotDefinition(next);
    },
    onRemoveField: (fieldId, area) => {
      const next = clonePivotDefinition(pivotDefinition);
      if (area === "values") next.values = next.values.filter((value) => value.fieldId !== fieldId);
      else next[area] = next[area].filter((field) => field !== fieldId);
      updatePivotDefinition(next);
    },
    onValueChange: (value) => {
      const next = clonePivotDefinition(pivotDefinition);
      const index = next.values.findIndex((entry) => entry.id === value.id);
      if (index >= 0) next.values[index] = value;
      updatePivotDefinition(next);
    },
    onCalculatedFieldsChange: (fields) => updatePivotDefinition({ ...clonePivotDefinition(pivotDefinition), calculatedFields: fields }),
    onCalculatedItemsChange: (items) => updatePivotDefinition({ ...clonePivotDefinition(pivotDefinition), calculatedItems: items }),
    onFilterChange: (fieldId, selectedValues) => {
      const next = clonePivotDefinition(pivotDefinition);
      next.filterSelections[fieldId] = [...selectedValues];
      updatePivotDefinition(next);
      actions.notify("Pivot filter updated");
    },
    onSortChange: (fieldId, direction) => {
      const next = clonePivotDefinition(pivotDefinition);
      next.sort[fieldId] = direction;
      updatePivotDefinition(next);
    },
    onGroupChange: (fieldId, grouped) => {
      const next = clonePivotDefinition(pivotDefinition);
      next.groupedFields = grouped ? [...new Set([...next.groupedFields, fieldId])] : next.groupedFields.filter((field) => field !== fieldId);
      updatePivotDefinition(next);
    },
    onRefresh: () => { if (activePivot) actions.refreshPivot(activePivot.id); },
    onLayoutChange: (layout) => updatePivotDefinition({ ...clonePivotDefinition(pivotDefinition), layout }),
    onExpandedChange: (fieldId, expanded) => {
      const next = clonePivotDefinition(pivotDefinition);
      next.expandedFieldIds = expanded ? [...new Set([...next.expandedFieldIds, fieldId])] : next.expandedFieldIds.filter((field) => field !== fieldId);
      updatePivotDefinition(next);
    },
    onSlicerChange: (fieldId, enabled) => {
      if (!activePivot) return;
      const next = clonePivotDefinition(pivotDefinition);
      next.slicers = enabled ? [...new Set([...next.slicers, fieldId])] : next.slicers.filter((field) => field !== fieldId);
      updatePivotDefinition(next);
      actions.updatePivotConfiguration(activePivot.id, { slicers: next.slicers.map((field) => ({ id: `slicer-${field}`, field, selected: [], connectedPivotIds: state.selectedSheet.pivots.map((pivot) => pivot.id) })) });
    },
    onTimelineChange: (fieldId) => {
      if (activePivot) actions.updatePivotConfiguration(activePivot.id, { timelines: fieldId ? [{ id: `timeline-${fieldId}`, field: fieldId, connectedPivotIds: state.selectedSheet.pivots.map((pivot) => pivot.id) }] : [] });
    },
    onTimelineRangeChange: (start, end) => {
      if (activePivot && pivotDefinition.timelineFieldId) actions.updatePivotConfiguration(activePivot.id, { timelines: [{ id: `timeline-${pivotDefinition.timelineFieldId}`, field: pivotDefinition.timelineFieldId, start: start || undefined, end: end || undefined, connectedPivotIds: state.selectedSheet.pivots.map((pivot) => pivot.id) }] });
    },
    onPivotChartChange: (chart) => {
      if (!activePivot || !chart) return;
      actions.addChart({ id: `pivot-chart-${activePivot.id}`, pivotId: activePivot.id, sheetId: activePivot.sheetId, type: chart.type, title: chart.title, sourceRanges: [activePivot.sourceRange], bounds: { x: 80, y: 80, width: 480, height: 280 } });
      actions.updatePivotConfiguration(activePivot.id, { chartReferences: [{ chartId: `pivot-chart-${activePivot.id}`, role: "linked" }] });
    },
  };
  const pivotPanelState: PivotPanelState = { disabled: isBusy, loading: state.phase === "loading", error: state.phase === "error" ? "Pivot data could not be loaded" : undefined, empty: pivotFields.length === 0 };
  const selectPanel = (panel: SidebarPanelId) => actions.handleRibbonAction(panel === "inspector" ? "open-comments" : `open-${panel.replace("conditionalFormat", "conditional-format").replace("dataValidation", "data-validation")}`);

  /** 画布选区 → workspace 多选区模型 */
  const applySelection = (selection: SelectionState) => {
    const range = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
    if (!range) return;
    actions.selectRange(
      { startRow: range.startRow, endRow: range.endRow, startColumn: range.startColumn, endColumn: range.endColumn },
      selection.ranges.length > 1 ? "add" : "replace",
    );
  };

  return (
    <>
      <AppShell
        formulaBar={
          <FormulaBar
            cellName={state.activeCell}
            disabled={isBusy}
            formula={state.formulaDraft}
            onCancel={() => actions.notify("Formula edit cancelled")}
            onChange={actions.setFormulaDraft}
            onCommit={() => actions.commitFormula()}
            onOpenWizard={() => actions.handleRibbonAction("function-wizard")}
            phase={state.phase}
          />
        }
        isBusy={isBusy}
        notice={state.notice}
        onShare={copyWorkbookLink}
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
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); copyWorkbookLink(); }}>
                  Copy workbook link
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); actions.handleRibbonAction("export-xlsx"); }}>
                  Export .xlsx
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); actions.handleRibbonAction("import-xlsx"); }}>
                  Import .xlsx
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); actions.handleRibbonAction("open-print"); }}>
                  Print / Save as PDF
                </Button>
              </Stack>
            )}
          </DropdownMenu>
        }
        ribbon={
          <Ribbon
            activeTab={state.ribbonTab}
            onAction={actions.handleRibbonAction}
            onTabChange={actions.setRibbonTab}
            phase={state.phase}
          />
        }
        saveState={state.saveState}
        sheetTabs={
          <SheetTabs
            activeSheetId={state.activeSheetId}
            disabled={isBusy}
            onAdd={actions.addSheet}
            onSelect={actions.selectSheet}
            onRenameSheet={actions.renameSheet}
            onDeleteSheet={actions.deleteSheet}
            sheets={state.sheets}
          />
        }
        statusBar={
          <StatusBar
            activeCell={state.activeCell}
            onOpenShortcuts={() => actions.notify("Shortcuts: Arrows / Tab / Enter / F2 / F4 / Ctrl+C/X/V/Z/Y/B/I/U")}
            onZoomChange={actions.setZoom}
            phase={state.phase}
            saveState={state.saveState}
            sheetCount={state.sheets.length}
            zoom={state.zoom}
          />
        }
        title={state.workbookName}
        workspacePhase={state.phase}
      >
        <Inline gap="none" className="h-full min-h-0 w-full flex-nowrap">
          <Box className="h-full min-h-0 min-w-0 flex-1">
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
              selectedFloatingId={state.selectedFloatingId}
              charts={state.selectedSheet.charts}
              pivotResults={state.selectedSheet.pivotResults}
              shapes={state.selectedSheet.shapes}
              sparklines={state.selectedSheet.sparklines}
              onSelectionChange={applySelection}
              onCommitCell={(value) => actions.commitFormula(value)}
              onBeginEdit={(initialText) => {
                if (initialText !== undefined) actions.setFormulaDraft(initialText);
                actions.beginEdit(initialText);
              }}
              onCancelEdit={actions.cancelEdit}
              onCommitEdit={(moveAfter) => actions.commitEdit(moveAfter ?? "down")}
              onFormulaDraftChange={actions.setFormulaDraft}
              onInsertRef={actions.insertRefIntoDraft}
              onToggleAbsolute={actions.toggleAbsoluteReference}
              onJumpEdge={(direction, extend) => actions.jumpEdge(direction, extend)}
              onSelectAll={actions.selectAll}
              onSelectRows={(startRow, endRow, additive) => actions.selectRowHeader(startRow, endRow, additive ? "add" : "replace")}
              onSelectColumns={(startColumn, endColumn, additive) => actions.selectColumnHeader(startColumn, endColumn, additive ? "add" : "replace")}
              onResizeRow={actions.resizeRow}
              onResizeColumn={actions.resizeColumn}
              onFillRange={actions.fillRange}
              onFloatingSelect={(hit) => actions.setSelectedFloatingId(hit ? hit.id : null)}
              onFloatingMove={(kind, id, bounds) => {
                const chart = state.selectedSheet.charts.find((entry) => entry.id === id);
                if (chart) actions.updateChartBounds(id, bounds);
                else actions.updateShapeBounds(id, bounds);
              }}
              onFloatingRemove={(kind, id) =>
                actions.removeFloatingObject(kind === "chart" ? "chart" : "shape", id)
              }
              onAction={actions.handleRibbonAction}
              onApplyFilter={(column, patch) => actions.applyFilter(column, patch)}
              getValidationList={actions.getValidationAt}
              onRetry={actions.retry}
              onCreateSheet={actions.addSheet}
            />
          </Box>
          <FeatureSidebar
            activeCell={state.activeCell}
            activePanel={state.activePanel}
            selectedRange={selectedRange}
            getRangeMatrix={actions.getRangeMatrix}
            getRangeNumbers={actions.getRangeNumbers}
            onPanelChange={selectPanel}
            onRetry={actions.retry}
            phase={state.phase}
            sheet={state.selectedSheet}
            sheetId={state.activeSheetId}
            charts={state.selectedSheet.charts}
            pivotDefinition={pivotDefinition}
            pivotList={state.selectedSheet.pivots.map((pivot) => ({ id: pivot.id, label: pivot.id }))}
            activePivotId={activePivot?.id}
            pivotFieldCatalog={pivotFields}
            pivotResult={pivotResult}
            onShowPivotDetails={actions.showPivotDetails}
            pivotPanelState={pivotPanelState}
            pivotCallbacks={pivotCallbacks}
            shapes={state.selectedSheet.shapes}
            sparklines={state.selectedSheet.sparklines}
            conditionalFormats={state.selectedSheet.conditionalFormats}
            dataValidations={state.selectedSheet.dataValidations}
            historyEntries={state.historyEntries}
            onAddChart={actions.addChart}
            onRemoveChart={actions.removeChart}
            onAddShape={actions.addShape}
            onRemoveShape={actions.removeShape}
            onAddSparkline={actions.addSparkline}
            onRemoveSparkline={actions.removeSparkline}
            onAddConditionalFormat={actions.addConditionalFormat}
            onRemoveConditionalFormat={actions.removeConditionalFormat}
            onAddDataValidation={actions.addDataValidation}
            onRemoveDataValidation={actions.removeDataValidation}
            onPrint={actions.printWorkbook}
            onExportPdf={actions.exportPdf}
            onAddComment={actions.addComment}
            onRemoveComment={actions.removeComment}
            onSetHyperlink={actions.setHyperlink}
            onRemoveHyperlink={actions.removeHyperlink}
          />
        </Inline>
      </AppShell>

      {/* Dialogs */}
      <FunctionWizardDialog
        open={state.showFunctionWizard}
        onClose={actions.closeFunctionWizard}
        onInsertFormula={(formula) => {
          actions.setFormulaDraft(formula);
          actions.commitFormula(formula);
        }}
      />

      <SortDialog
        open={state.showSortDialog}
        columns={state.selectedSheet.columns}
        onClose={actions.closeSortDialog}
        onSort={(criteria, hasHeader) => actions.sortRange(criteria, hasHeader)}
      />

      <FindReplaceDialog
        open={state.showFindReplace}
        onClose={actions.closeFindReplace}
        onReplaceAll={(params) => actions.findReplace(params)}
      />

      <PrintPreviewDialog
        open={state.showPrintPreview}
        onClose={() => actions.setShowPrintPreview(false)}
        sheetId={state.activeSheetId}
        rowCount={state.selectedSheet.rowCount}
        columnCount={state.selectedSheet.columns.length}
        columns={state.selectedSheet.columns}
        rows={state.selectedSheet.rows}
      />
    </>
  );
}

export default function App() {
  if (typeof window !== "undefined" && window.location.pathname === "/workbooks") return <WorkbookCatalog />;
  return <WorkspaceApp />;
}
