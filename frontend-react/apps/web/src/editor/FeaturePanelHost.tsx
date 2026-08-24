import { lazy, Suspense, type ReactNode } from "react";
import { Box, SidebarShell } from "@react-sheets/ui-system";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import type { UiSessionIntent, UiSnapshot, WorkbookSession } from "@react-sheets/spreadsheet-app";
import type { Locale } from "../i18n";
import type { EditorCommandController } from "./command-controller";

const FeatureSidebar = lazy(() => import("../components/FeatureSidebar").then((module) => ({ default: module.FeatureSidebar })));

export interface FeaturePanelHostProps {
  state: UiSnapshot;
  session: WorkbookSession;
  locale: Locale;
  sidebarOpen: boolean;
  onSidebarOpenChange: (open: boolean) => void;
  selectedRange: EditorCommandController["selectedRange"];
  dispatchCommand: (descriptor: CommandDescriptor) => void;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
  commands: Pick<
    EditorCommandController,
    | "activePivot"
    | "activePivotId"
    | "pivotFields"
    | "pivotSlicerControls"
    | "pivotTimelineControls"
    | "pivotPanelState"
    | "pivotCallbacks"
    | "setActivePivotId"
  >;
  title: string;
}

/** Feature panel boundary. It owns panel-only wiring and keeps DesignerShell free of feature props. */
export function FeaturePanelHost({
  state,
  session,
  locale,
  sidebarOpen,
  onSidebarOpenChange,
  selectedRange,
  dispatchCommand,
  dispatchSessionIntent,
  commands,
  title,
}: FeaturePanelHostProps): ReactNode {
  return (
    <SidebarShell open={sidebarOpen} onOpenChange={onSidebarOpenChange} title={title} showHeader={state.panels.active !== 'pivot'} width={state.panels.active === 'pivot' ? 390 : state.panels.width} minWidth={state.panels.active === 'pivot' ? 360 : undefined} maxWidth={state.panels.active === 'pivot' ? 480 : undefined}>
      <Suspense fallback={<Box className="h-full min-h-0" />}>
        <FeatureSidebar
          activeCell={state.activeCell}
          activePanel={state.panels.active}
          locale={locale}
          selectedRange={selectedRange}
          onPanelChange={(panel) => {
            onSidebarOpenChange(true);
            if (panel === "print") {
              dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" });
              return;
            }
            if (panel === "selectionPane") session.setDrawingSelectionMode(true);
            dispatchSessionIntent({ type: "panel.open", panel });
          }}
          onClosePanel={() => onSidebarOpenChange(false)}
          onRetry={session.retry.bind(session)}
          phase={state.phase}
          sheet={state.selectedSheet}
          sheetId={state.activeSheetId}
          drawings={state.selectedSheet.drawings}
          drawingPayloads={state.selectedSheet.drawingPayloads}
          selectedDrawingIds={state.selectedDrawingIds}
          onSelectDrawing={(drawingId, mode) => session.setDrawingSelection([drawingId], mode === "extend" ? "add" : mode)}
          onSetDrawingVisibility={(drawingId, visible) => session.setDrawingVisibility(drawingId, visible)}
          onRenameDrawing={(drawingId, name) => session.renameDrawing(drawingId, name)}
          onReorderDrawing={(drawingId, direction) => dispatchCommand({ commandId: "drawing.zorder", params: { sheetId: state.activeSheetId, drawingId, direction } })}
          pivot={commands.activePivot}
          pivotList={state.selectedSheet.pivots.map((pivot) => ({ id: pivot.id, label: pivot.id }))}
          activePivotId={commands.activePivotId}
          pivotFieldCatalog={commands.pivotFields}
          pivotSlicerControls={commands.pivotSlicerControls}
          pivotTimelineControls={commands.pivotTimelineControls}
          pivotPanelState={commands.pivotPanelState}
          pivotCallbacks={commands.pivotCallbacks}
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
          lastWhatIfMessage={state.lastWhatIfResult && "message" in state.lastWhatIfResult
            ? state.lastWhatIfResult.message
            : state.lastWhatIfResult && "status" in state.lastWhatIfResult
              ? `${state.lastWhatIfResult.kind}: ${state.lastWhatIfResult.status}`
              : null}
          canRunExtended={state.permissions.script}
          onGoalSeek={(params) => session.runGoalSeek({ setCell: { row: params.setRow, column: params.setColumn }, toValue: params.targetValue, byChangingCell: { row: params.changingRow, column: params.changingColumn } })}
          onRunDataTable={(params) => session.runDataTableAnalysis({ tableRange: params.tableRange, ...(params.inputMode === "column" ? { columnInputCell: params.inputCell } : { rowInputCell: params.inputCell }) })}
          onRunScenario={(params) => session.runScenarioAnalysis({
            id: `scenario-${Date.now()}`,
            name: params.name,
            changingCells: [{ row: params.changingCell.row, column: params.changingCell.column, value: params.changingValue }],
            resultCells: [{ row: params.resultCell.row, column: params.resultCell.column }],
          })}
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
  );
}
