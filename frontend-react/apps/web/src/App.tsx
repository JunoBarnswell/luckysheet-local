import { Box, Inline } from "@react-sheets/ui-system";
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
import { getInitialWorkspacePhase, useWorkspaceState, type SelectionState } from "./state/workspace";

export default function App() {
  const { actions, state } = useWorkspaceState({ initialPhase: getInitialWorkspacePhase() });
  const isBusy = state.phase !== "ready";

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
        onMenu={() => actions.notify("Workbook menu opened")}
        onShare={() => actions.notify("Share link copied to clipboard")}
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
        title="Q3 Growth Planning"
        workspacePhase={state.phase}
      >
        <Inline gap="none" className="h-full min-h-0 w-full flex-nowrap">
          <Box className="min-w-0 flex-1">
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
              onSelectRows={(startRow, _endRow, additive) => actions.selectRowHeader(startRow, additive ? "add" : "replace")}
              onSelectColumns={(startColumn, _endColumn, additive) => actions.selectColumnHeader(startColumn, additive ? "add" : "replace")}
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
              onFilterColumn={() => undefined}
              getValidationList={actions.getValidationAt}
              onRetry={actions.retry}
              onCreateSheet={actions.addSheet}
            />
          </Box>
          <FeatureSidebar
            activeCell={state.activeCell}
            activePanel={state.activePanel}
            selectedRange={state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0]}
            getRangeMatrix={actions.getRangeMatrix}
            getRangeNumbers={actions.getRangeNumbers}
            onPanelChange={actions.setActivePanel}
            onRetry={actions.retry}
            phase={state.phase}
            sheet={state.selectedSheet}
            sheetId={state.activeSheetId}
            charts={state.selectedSheet.charts}
            pivots={state.selectedSheet.pivots}
            shapes={state.selectedSheet.shapes}
            sparklines={state.selectedSheet.sparklines}
            conditionalFormats={state.selectedSheet.conditionalFormats}
            dataValidations={state.selectedSheet.dataValidations}
            historyEntries={state.historyEntries}
            onAddChart={actions.addChart}
            onRemoveChart={actions.removeChart}
            onAddPivot={actions.addPivot}
            onRefreshPivot={actions.refreshPivot}
            onRemovePivot={actions.removePivot}
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
      />
    </>
  );
}