import type { ReactNode } from "react";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { Ribbon } from "../components/Ribbon";
import type { RibbonTabId, UiSessionIntent, UiSnapshot, WorkbookSession } from "@react-sheets/spreadsheet-app";
import type { DrawingConnectorType } from '@react-sheets/core-model';
import type { Locale } from "../i18n";
import type { EditorCommandController } from "./command-controller";
import type { ColumnDimensionController } from './column-dimension-controller';

export interface RibbonHostProps {
  state: UiSnapshot;
  session: WorkbookSession;
  locale: Locale;
  isBusy: boolean;
  dispatchCommand: (descriptor: CommandDescriptor) => void;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
  saveWorkbook: () => void;
  exportDocument: () => void | Promise<void>;
  importDocument: () => void;
  columnDimensions: ColumnDimensionController;
  selectedColumns: number[];
  selectedRows: number[];
  onOpenColumnWidthDialog: (columns: number[]) => void;
  onOpenDefaultColumnWidthDialog: () => void;
  onOpenRowHeightDialog: (rows: number[]) => void;
  commands: Pick<
    EditorCommandController,
    | "buildSortDescriptor"
    | "buildTotalRowCommand"
    | "buildFilterSelectionCommand"
    | "buildClearFilterCommand"
    | "buildOutlineCommand"
    | "buildSubtotalCommand"
    | "buildRemoveDuplicatesCommand"
    | "buildTextToColumnsCommand"
    | "pivotRibbonActions"
  >;
}

/** Ribbon-only orchestration. Domain command payload construction lives in command-controller. */
export function RibbonHost({
  state,
  session,
  locale,
  isBusy,
  dispatchCommand,
  dispatchSessionIntent,
  saveWorkbook,
  exportDocument,
  importDocument,
  columnDimensions,
  selectedColumns,
  selectedRows,
  onOpenColumnWidthDialog,
  onOpenDefaultColumnWidthDialog,
  onOpenRowHeightDialog,
  commands,
}: RibbonHostProps): ReactNode {
  const activeTableContext = state.activeContext.kind === 'table' ? state.activeContext : undefined;
  const activeChartDrawing = state.selectedSheet.drawings.find((drawing) => drawing.id === state.selectedFloatingId && drawing.kind === 'chart');
  const activePictureDrawing = state.selectedSheet.drawings.find((drawing) => drawing.id === state.selectedFloatingId && drawing.kind === 'image');
  const activeShapeDrawings = state.selectedSheet.drawings.filter((drawing) => state.selectedDrawingIds.includes(drawing.id) && (drawing.kind === 'shape' || drawing.kind === 'connector'));
  const activeShape = activeShapeDrawings.length > 0
    ? {
      sheetId: state.activeSheetId,
      drawingIds: activeShapeDrawings.map((drawing) => drawing.id),
      transforms: activeShapeDrawings.map((drawing) => ({ drawingId: drawing.id, transform: drawing.transform })),
    }
    : undefined;
  const activeSparklineContext = state.activeContext.kind === 'sparkline' ? state.activeContext : undefined;
  const runSessionAction = (action: () => void): void => {
    try { action(); }
    catch (error) { session.notify(error instanceof Error ? error.message : 'Ribbon action failed'); }
  };
  return (
    <Ribbon
      session={session}
      activeTab={state.ribbon.activeTab}
      activePivot={state.activeContext.kind === "pivot"
        ? { sheetId: state.activeContext.sheetId, pivotId: state.activeContext.pivotId }
        : undefined}
      pivotActions={commands.pivotRibbonActions}
      activeTableSheet={state.activeContext.kind === "table-sheet"
        ? { sheetId: state.activeContext.sheetId, viewId: state.activeContext.viewId }
        : undefined}
      activeGanttSheet={state.activeContext.kind === "gantt-sheet"
        ? { sheetId: state.activeContext.sheetId, viewId: state.activeContext.viewId }
        : undefined}
      activeReportSheet={state.activeContext.kind === "report-sheet"
        ? { sheetId: state.activeContext.sheetId, tableId: state.activeContext.tableId }
        : undefined}
      activeTable={activeTableContext
        ? (() => {
          const table = state.selectedSheet.sheetTables.find((entry) => entry.id === activeTableContext.tableId);
          const selection = state.selection.ranges[state.selection.primaryRangeIndex];
          const resizeRange = table && selection
            && (selection.startRow !== table.range.startRow || selection.endRow !== table.range.endRow
              || selection.startColumn !== table.range.startColumn || selection.endColumn !== table.range.endColumn)
            ? { ...selection }
            : undefined;
          return table ? { sheetId: activeTableContext.sheetId, tableId: table.id, table, resizeRange } : undefined;
        })()
        : undefined}
      activeChart={activeChartDrawing ? { sheetId: activeChartDrawing.sheetId, chartId: activeChartDrawing.payloadId } : undefined}
      activePicture={activePictureDrawing ? { sheetId: activePictureDrawing.sheetId, drawingId: activePictureDrawing.id } : undefined}
      activeShape={activeShape}
      activeSparkline={activeSparklineContext ? { sheetId: activeSparklineContext.sheetId, sparklineId: activeSparklineContext.sparklineId } : undefined}
      locale={locale}
      onCommand={dispatchCommand}
      onSessionIntent={dispatchSessionIntent}
      onCopy={() => session.copy()}
      onCut={() => session.cut()}
      onPaste={() => session.paste()}
      onUndo={() => session.undo()}
      onRedo={() => session.redo()}
      onSave={saveWorkbook}
      onExportDocument={() => { void exportDocument(); }}
      onImportDocument={importDocument}
      onRecalculate={() => { void session.recalculateFormulas(); }}
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
      onAutoSum={(functionName) => session.autoSum(functionName)}
      onFill={(direction, mode) => mode === 'series' ? dispatchSessionIntent({ type: 'dialog.open', dialog: 'fill-series' }) : session.fillSelection(direction)}
      onFreezeAtPrimary={() => session.freezeAtPrimary()}
      onOpenColumnWidth={() => onOpenColumnWidthDialog(selectedColumns)}
      onAutoFitColumns={() => { void columnDimensions.autoFit(selectedColumns).catch((error) => session.notify(error instanceof Error ? error.message : 'Column AutoFit failed')); }}
      onHideColumns={() => columnDimensions.setHidden(selectedColumns, true)}
      onUnhideColumns={() => columnDimensions.setHidden(selectedColumns, false)}
      onOpenDefaultColumnWidth={onOpenDefaultColumnWidthDialog}
      onOpenRowHeight={() => onOpenRowHeightDialog(selectedRows)}
      onAutoFitRows={() => { void columnDimensions.autoFitRows(selectedRows).catch((error) => session.notify(error instanceof Error ? error.message : 'Row AutoFit failed')); }}
      onHideRows={() => columnDimensions.setRowsHidden(selectedRows, true)}
      onUnhideRows={() => columnDimensions.setRowsHidden(selectedRows, false)}
      onCreatePivotDialog={() => dispatchSessionIntent({ type: "dialog.open", dialog: "create-pivot" })}
      buildSortDescriptor={commands.buildSortDescriptor}
      onCreateSheetTable={() => session.openCreateTableDialog()}
      onOpenTableSettings={() => session.openTableSettings()}
      onToggleTableOption={(option) => session.toggleActiveSheetTableOption(option)}
      onConvertActiveTableToRange={() => session.convertActiveSheetTableToRange()}
      onCreateDataSource={() => { void session.createDataSourceFromSelection().catch((error) => session.notify(error instanceof Error ? error.message : 'Data Source creation failed')); }}
      onToggleSheetTableTotalRow={commands.buildTotalRowCommand}
      onApplyFilterSelection={commands.buildFilterSelectionCommand}
      onClearFilter={commands.buildClearFilterCommand}
      onGroupRows={() => commands.buildOutlineCommand("row", "add")}
      onUngroupRows={() => commands.buildOutlineCommand("row", "remove")}
      onGroupColumns={() => commands.buildOutlineCommand("column", "add")}
      onUngroupColumns={() => commands.buildOutlineCommand("column", "remove")}
      onSubtotal={commands.buildSubtotalCommand}
      onRemoveDuplicates={commands.buildRemoveDuplicatesCommand}
      onTextToColumns={commands.buildTextToColumnsCommand}
      onResolveComment={() => session.resolveComment()}
      onProtectSelection={() => session.protectSelection()}
      onUnprotectSelection={() => session.unprotectSelection()}
      onShowOutlineLevel={(level: 1 | 2 | 3) => session.showOutlineLevel(level)}
      onTransposeSelection={() => session.transposeSelection()}
      onFlipSelection={(axis: "h" | "v") => session.flipSelection(axis)}
      onSplitByDelimiter={() => session.splitByDelimiter(",")}
      onToggleBandedRows={() => session.toggleBandedRows()}
      onSetRecalculationMode={(mode: "automatic" | "manual" | "partial") => session.setRecalculationMode(mode)}
      onOpenDefinedNames={() => dispatchSessionIntent({ type: "panel.open", panel: "definedNames" })}
      onCreateAdvancedSheet={(kind) => session.createAdvancedSheet(kind)}
      onApplyBarcode={(symbology) => session.openBarcodePanel(symbology)}
      onCreateCamera={() => runSessionAction(() => session.insertCamera())}
      onCaptureScreenshot={() => session.captureScreenshot()}
      onCreateFormControl={(type) => runSessionAction(() => session.insertFormControl(type))}
      onApplyCheckbox={() => runSessionAction(() => session.setCellEditor({ kind: 'checkbox' }))}
      onCreateTextBox={() => runSessionAction(() => session.insertTextBox())}
      onInsertChartType={(type, subtype) => runSessionAction(() => session.insertChart(type, subtype))}
      onInsertSparklineType={(type) => runSessionAction(() => session.insertSparkline(type))}
      onInsertShapeType={(type) => runSessionAction(() => session.insertShape(type))}
      onInsertConnectorType={(type: DrawingConnectorType) => runSessionAction(() => session.insertConnector(type))}
      onTabChange={(tab: RibbonTabId) => session.setRibbonTab(tab)}
      phase={state.phase}
      homeState={state.homeRibbon}
      commandPaletteOpen={state.dialogs.active === 'command-palette'}
      onCloseCommandPalette={session.closeCommandPalette}
      formatPainterActive={state.formatPainter !== null}
      onBeginFormatPainter={(locked) => session.beginFormatPainter(Boolean(locked))}
      onMergeCells={(operation) => session.requestMergeAction(operation)}
      canExecute={session.canExecute.bind(session)}
      featureSurfaceSchema={session.getFeatureSurfaceSchema()}
    />
  );
}
