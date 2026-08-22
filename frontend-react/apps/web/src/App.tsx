import { Box, Inline } from '@react-sheets/ui-system';
import { AppShell } from './components/AppShell';
import { FeatureSidebar } from './components/FeatureSidebar';
import { FormulaBar } from './components/FormulaBar';
import { Ribbon } from './components/Ribbon';
import { SheetCanvas } from './components/SheetCanvas';
import { SheetTabs } from './components/SheetTabs';
import { StatusBar } from './components/StatusBar';
import { FunctionWizardDialog } from './components/dialogs/FunctionWizardDialog';
import { SortDialog } from './components/dialogs/SortDialog';
import { getInitialWorkspacePhase, useWorkspaceState } from './state/workspace';

export default function App() {
  const { actions, state } = useWorkspaceState({ initialPhase: getInitialWorkspacePhase() });
  const isBusy = state.phase !== 'ready';

  return (
    <>
      <AppShell
        formulaBar={
          <FormulaBar
            cellName={state.activeCell}
            disabled={isBusy}
            formula={state.formulaDraft}
            onCancel={() => actions.notify('Formula edit cancelled')}
            onChange={actions.setFormulaDraft}
            onCommit={() => actions.commitFormula()}
            onOpenWizard={() => actions.handleRibbonAction('function-wizard')}
            phase={state.phase}
          />
        }
        isBusy={isBusy}
        notice={state.notice}
        onMenu={() => actions.notify('Workbook menu opened')}
        onShare={() => actions.notify('Share link copied to clipboard')}
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
            onOpenShortcuts={() => actions.notify('Shortcuts: Arrows / Tab / Enter / F2 / Ctrl+C / Ctrl+V / Ctrl+Z')}
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
              selectedCell={state.activeCell}
              formulaDraft={state.formulaDraft}
              phase={state.phase}
              zoom={state.zoom}
              charts={state.selectedSheet.charts}
              shapes={state.selectedSheet.shapes}
              sparklines={state.selectedSheet.sparklines}
              onSelectCell={actions.selectCell}
              onCommitCell={(val) => actions.commitFormula(val)}
              onMoveCell={actions.moveCell}
              onAction={actions.handleRibbonAction}
              onRetry={actions.retry}
              onCreateSheet={actions.addSheet}
            />
          </Box>
          <FeatureSidebar
            activeCell={state.activeCell}
            activePanel={state.activePanel}
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
        onSort={actions.sortRange}
      />
    </>
  );
}
