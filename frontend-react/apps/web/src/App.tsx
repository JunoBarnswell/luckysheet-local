import { AppShell, Box, Button, DropdownMenu, Inline, SidebarShell, Stack } from "@react-sheets/ui-system";
import { FeatureSidebar } from "./components/FeatureSidebar";
import { FormulaBar } from "./components/FormulaBar";
import { Ribbon } from "./components/Ribbon";
import { SheetCanvas } from "./components/SheetCanvas";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { FunctionWizardDialog } from "./components/dialogs/FunctionWizardDialog";
import { SortDialog } from "./components/dialogs/SortDialog";
import { FindReplaceDialog } from "./components/dialogs/FindReplaceDialog";
import { GoToDialog } from "./components/dialogs/GoToDialog";
import { PasteSpecialDialog } from "./components/dialogs/PasteSpecialDialog";
import { FormatCellsDialog } from "./components/dialogs/FormatCellsDialog";
import { ShiftCellsDialog } from "./components/dialogs/ShiftCellsDialog";
import { PrintPreviewDialog } from "./components/dialogs/PrintPreviewDialog";
import { WorkbookCatalog } from "./components/WorkbookCatalog";
import { WorkspaceErrorBoundary } from "./components/WorkspaceErrorBoundary";
import { parseRangeInput } from "./domain/range-input";
import { mapRibbonAction } from "./domain/ribbon-command-map";
import {
  addFieldToLayout,
  clonePivotLayout,
  setExpandedField,
  setFieldGrouped,
  setFieldSort,
  setFilterSelection,
  setLayoutMode,
  updateValueInLayout,
} from "./domain/pivot-layout-ops";
import type { RibbonAction } from "./domain/ribbon-actions";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  getInitialAppPhase,
  useSpreadsheetApp,
  type SelectionState,
  type SidebarPanelId,
} from "@react-sheets/spreadsheet-app";
import { getInitialLocale, localeLabels, persistLocale, shellLabels, type Locale } from "./i18n";
import zhCN from "./locales/zh-CN.json";
import enUS from "./locales/en-US.json";
import type { PivotFieldDefinition as CorePivotFieldDefinition } from "@react-sheets/core-model";
import type { PivotFieldDefinition, PivotPanelCallbacks, PivotPanelState, PivotResult } from "./components/pivot/types";

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

  const executeRibbonAction = (action: RibbonAction, payload?: unknown) => {
    const { commandId, params } = mapRibbonAction(action, payload);
    app.execute(commandId, params);
  };

  useEffect(() => {
    if (previousPanelRef.current !== state.activePanel) {
      setSidebarOpen(true);
    }
    previousPanelRef.current = state.activePanel;
  }, [state.activePanel]);

  const copyWorkbookLink = () => {
    const link = `${window.location.origin}/workbooks/${encodeURIComponent(state.unitId)}`;
    const clipboard = navigator.clipboard;
    if (!clipboard) {
      app.notify("Clipboard access is unavailable");
      return;
    }
    void clipboard.writeText(link)
      .then(() => app.notify("Workbook link copied"))
      .catch(() => app.notify("Could not copy workbook link"));
  };

  const selectedRange = state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0];
  const pivotSourceRange = selectedRange && (selectedRange.endRow > selectedRange.startRow || selectedRange.endColumn > selectedRange.startColumn)
    ? selectedRange
    : state.selectedSheet.usedRange;
  const [activePivotId, setActivePivotId] = useState<string>();
  const activePivot = state.selectedSheet.pivots.find((pivot) => pivot.id === activePivotId) ?? state.selectedSheet.pivots[0];
  const pivotTree = activePivot ? state.selectedSheet.pivotResults[activePivot.id] : undefined;
  const corePivotFields = pivotTree?.fields.fields ?? app.getPivotFieldCatalog(pivotSourceRange);
  const pivotFields: PivotFieldDefinition[] = corePivotFields.map((field: CorePivotFieldDefinition) => ({
    id: field.id,
    label: field.name,
    type: field.dataType === "mixed" ? "text" : field.dataType,
    values: field.values?.map((value) => String(value)),
  }));
  const pivotResult: PivotResult | undefined = pivotTree
    ? { rowCount: pivotTree.rows.length, columnCount: pivotTree.columnPaths.length, tree: pivotTree, summary: `${pivotTree.rows.length} row groups × ${pivotTree.columnPaths.length || 1} column groups` }
    : undefined;

  const updatePivotLayout = (nextLayout: ReturnType<typeof clonePivotLayout>) => {
    if (!activePivot) return;
    app.updatePivotLayout(activePivot.id, nextLayout);
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
      updatePivotLayout(addFieldToLayout(activePivot.layout, fieldId, area, index, pivotFields));
    },
    onRemoveField: (fieldId, area) => {
      if (!activePivot) return;
      const next = clonePivotLayout(activePivot.layout);
      if (area === "values") next.values = next.values.filter((value) => value.field !== fieldId);
      else if (area === "filters") next.filters = next.filters.filter((filter) => filter.field !== fieldId);
      else next[area] = next[area].filter((field) => field.field !== fieldId);
      updatePivotLayout(next);
    },
    onValueChange: (value) => {
      if (!activePivot) return;
      updatePivotLayout(updateValueInLayout(activePivot.layout, value));
    },
    onCalculatedFieldsChange: (fields) => {
      if (!activePivot) return;
      updatePivotLayout({ ...clonePivotLayout(activePivot.layout), calculatedFields: fields.map((field) => ({ ...field })) });
    },
    onCalculatedItemsChange: (items) => {
      if (!activePivot) return;
      updatePivotLayout({
        ...clonePivotLayout(activePivot.layout),
        calculatedItems: items.map((item) => ({ field: item.fieldId, name: item.name, formula: item.formula })),
      });
    },
    onFilterChange: (fieldId, selectedValues) => {
      if (!activePivot) return;
      updatePivotLayout(setFilterSelection(activePivot.layout, fieldId, selectedValues));
      app.notify("Pivot filter updated");
    },
    onSortChange: (fieldId, direction) => {
      if (!activePivot) return;
      updatePivotLayout(setFieldSort(activePivot.layout, fieldId, direction));
    },
    onGroupChange: (fieldId, grouped) => {
      if (!activePivot) return;
      updatePivotLayout(setFieldGrouped(activePivot.layout, fieldId, grouped, pivotFields));
    },
    onRefresh: () => { if (activePivot) app.refreshPivot(activePivot.id); },
    onSourceRangeChange: (sourceRange) => {
      if (!activePivot) return;
      const parsed = parseRangeInput(sourceRange, activePivot.sheetId);
      if (!parsed) {
        app.notify("Invalid pivot source range");
        return;
      }
      app.updatePivotConfiguration(activePivot.id, { sourceRange: { sheetId: activePivot.sheetId, ...parsed } });
    },
    onLayoutChange: (layout) => {
      if (!activePivot) return;
      updatePivotLayout(setLayoutMode(activePivot.layout, layout));
    },
    onExpandedChange: (fieldId, expanded) => {
      if (!activePivot) return;
      updatePivotLayout(setExpandedField(activePivot.layout, fieldId, expanded));
    },
    onSlicerChange: (fieldId, enabled) => {
      if (!activePivot) return;
      const connectedPivotIds = app.getConnectedPivotIds(activePivot.sourceRange);
      if (enabled) {
        app.setPivotSlicer(
          activePivot.id,
          { id: `slicer-${fieldId}`, field: fieldId, selected: [], connectedPivotIds },
        );
        return;
      }
      app.updatePivotConfiguration(activePivot.id, {
        slicers: (activePivot.slicers ?? []).filter((slicer) => slicer.field !== fieldId),
      });
    },
    onTimelineChange: (fieldId) => {
      if (!activePivot) return;
      const connectedPivotIds = app.getConnectedPivotIds(activePivot.sourceRange);
      if (!fieldId) {
        app.updatePivotConfiguration(activePivot.id, { timelines: [] });
        return;
      }
      app.setPivotTimeline(activePivot.id, { id: `timeline-${fieldId}`, field: fieldId, connectedPivotIds });
    },
    onTimelineRangeChange: (start, end) => {
      if (!activePivot) return;
      const timelineFieldId = activePivot.timelines?.[0]?.field;
      if (!timelineFieldId) return;
      const connectedPivotIds = app.getConnectedPivotIds(activePivot.sourceRange);
      app.setPivotTimeline(
        activePivot.id,
        { id: `timeline-${timelineFieldId}`, field: timelineFieldId, start: start || undefined, end: end || undefined, connectedPivotIds },
      );
    },
    onPivotChartChange: (chart) => {
      if (!activePivot || !chart) return;
      const chartId = `pivot-chart-${activePivot.id}-${Date.now().toString(36)}`;
      app.addChart({
        id: chartId,
        pivotId: activePivot.id,
        sheetId: activePivot.sheetId,
        type: chart.type,
        title: chart.title,
        sourceRanges: [activePivot.sourceRange],
        bounds: { x: 80, y: 80, width: 480, height: 280 },
      });
      app.updatePivotConfiguration(activePivot.id, {
        chartReferences: [...(activePivot.chartReferences ?? []), { chartId, role: "linked" }],
      });
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
    const actionByPanel: Partial<Record<SidebarPanelId, RibbonAction>> = {
      inspector: "open-comments",
      chart: "open-chart",
      pivot: "open-pivot",
      shape: "open-shape",
      sparkline: "open-sparkline",
      conditionalFormat: "open-conditional-format",
      dataValidation: "open-data-validation",
      print: "open-print",
      history: "open-history",
      data: "open-data-table",
    };
    const action = actionByPanel[panel];
    if (action) executeRibbonAction(action);
    else app.setActivePanel(panel);
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
            onOpenWizard={() => executeRibbonAction("function-wizard")}
            phase={state.phase}
          />
        }
        isBusy={isBusy}
        labels={shellLabels(locale, state.saveState)}
        localeMenuLabel={localeLabels[locale]}
        notice={state.notice}
        onLocaleChange={setLocale}
        onSearch={(query) => executeRibbonAction("find-replace", query)}
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
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); executeRibbonAction("export-xlsx"); }}>
                  Export .xlsx
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); executeRibbonAction("import-xlsx"); }}>
                  Import .xlsx
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); executeRibbonAction("open-print"); }}>
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
            onExecute={app.execute.bind(app)}
            onTabChange={app.setRibbonTab}
            phase={state.phase}
            canExecute={app.canExecute.bind(app)}
            shareRole={state.shareRole}
            onRoleChange={(role) => app.execute('permission.role.set', { role })}
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
            hasLocalDraft={state.hasLocalDraft}
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
              images={state.selectedSheet.images}
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
              onFloatingMove={(kind, id, bounds) => {
                if (kind === "chart") app.updateChartBounds(id, bounds);
                else if (kind === "image") app.updateImageBounds(id, bounds);
                else app.updateShapeBounds(id, bounds);
              }}
              onFloatingRemove={(kind, id) => app.removeFloatingObject(kind, id)}
              onAction={executeRibbonAction}
              onApplyFilter={(column, patch) => app.applyFilter(column, patch)}
              onToggleOutline={(groupId) => app.toggleOutlineGroup(groupId)}
              getValidationList={app.getValidationAt}
              onRetry={app.retry}
              onCreateSheet={app.addSheet}
            />
          </Box>
          <SidebarShell
            open={sidebarOpen}
            onOpenChange={setSidebarOpen}
            title={locale === 'zh-CN' ? zhCN.sidebar.title : enUS.sidebar.title}
          >
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
            charts={state.selectedSheet.charts}
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
            shapes={state.selectedSheet.shapes}
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
            onAddChart={app.addChart}
            onRemoveChart={app.removeChart}
            onAddShape={app.addShape}
            onRemoveShape={app.removeShape}
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
        </SidebarShell>
        </Inline>
      </AppShell>

      {/* Dialogs */}
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
    </>
  );
}

export default function App() {
  if (typeof window !== "undefined" && window.location.pathname === "/workbooks") return <WorkbookCatalog />;
  return <WorkspaceErrorBoundary><WorkspaceApp /></WorkspaceErrorBoundary>;
}
