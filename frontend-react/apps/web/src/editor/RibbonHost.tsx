import type { ReactNode } from "react";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { Ribbon } from "../components/Ribbon";
import type { RibbonTabId, UiSessionIntent, UiSnapshot, WorkbookSession } from "@react-sheets/spreadsheet-app";
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
  exportXlsx: () => void | Promise<void>;
  importXlsx: () => void;
  columnDimensions: ColumnDimensionController;
  onOpenColumnWidthDialog: (columns: number[]) => void;
  onOpenDefaultColumnWidthDialog: () => void;
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
  exportXlsx,
  importXlsx,
  columnDimensions,
  onOpenColumnWidthDialog,
  onOpenDefaultColumnWidthDialog,
  commands,
}: RibbonHostProps): ReactNode {
  const activeTableContext = state.activeContext.kind === 'table' ? state.activeContext : undefined;
  const activeChartDrawing = state.selectedSheet.drawings.find((drawing) => drawing.id === state.selectedFloatingId && drawing.kind === 'chart');
  return (
    <Ribbon
      activeTab={state.ribbon.activeTab}
      activePivot={state.activeContext.kind === "pivot"
        ? { sheetId: state.activeContext.sheetId, pivotId: state.activeContext.pivotId }
        : undefined}
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
      locale={locale}
      onCommand={dispatchCommand}
      onSessionIntent={dispatchSessionIntent}
      onCopy={() => session.copy()}
      onCut={() => session.cut()}
      onPaste={() => session.paste()}
      onUndo={() => session.undo()}
      onRedo={() => session.redo()}
      onSave={saveWorkbook}
      onExportXlsx={() => { void exportXlsx(); }}
      onImportXlsx={importXlsx}
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
      onAutoSum={() => session.autoSum()}
      onFill={(direction) => session.fillSelection(direction)}
      onFreezeAtPrimary={() => session.freezeAtPrimary()}
      onOpenColumnWidth={() => onOpenColumnWidthDialog(columnDimensions.selectedColumns())}
      onAutoFitColumns={() => { void columnDimensions.autoFit(columnDimensions.selectedColumns()); }}
      onHideColumns={() => columnDimensions.setHidden(columnDimensions.selectedColumns(), true)}
      onUnhideColumns={() => columnDimensions.setHidden(columnDimensions.selectedColumns(), false)}
      onOpenDefaultColumnWidth={onOpenDefaultColumnWidthDialog}
      onCreatePivotDialog={() => dispatchSessionIntent({ type: "dialog.open", dialog: "create-pivot" })}
      buildSortDescriptor={commands.buildSortDescriptor}
      onCreateSheetTable={() => session.openCreateTableDialog()}
      onOpenTableSettings={() => session.openTableSettings()}
      onToggleTableOption={(option) => session.toggleActiveSheetTableOption(option)}
      onConvertActiveTableToRange={() => session.convertActiveSheetTableToRange()}
      onCreateDataTable={() => session.createDataTableFromSelection()}
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
      onSetRecalculationMode={(mode: "automatic" | "manual") => session.setRecalculationMode(mode)}
      onOpenDefinedNames={() => dispatchSessionIntent({ type: "panel.open", panel: "definedNames" })}
      onCreateAdvancedSheet={(kind) => session.createAdvancedSheet(kind)}
      onApplyBarcode={(symbology) => session.openBarcodePanel(symbology)}
      onCreateDataChart={(type) => session.insertDataChart(type)}
      onCreateCamera={() => session.insertCamera()}
      onCreateFormControl={(type) => session.insertFormControl(type)}
      onApplyCheckbox={() => session.setCellEditor({ kind: 'checkbox' })}
      onCreateTextBox={() => session.insertTextBox()}
      onInsertChartType={(type) => session.insertChart(type)}
      onInsertSparklineType={(type) => { session.insertSparkline(type); }}
      onInsertShapeType={(type) => session.insertShape(type)}
      onTabChange={(tab: RibbonTabId) => session.setRibbonTab(tab)}
      phase={state.phase}
      homeState={state.homeRibbon}
      commandPaletteOpen={state.dialogs.active === 'command-palette'}
      onCloseCommandPalette={session.closeCommandPalette}
      formatPainterActive={state.formatPainter !== null}
      onBeginFormatPainter={(locked) => session.beginFormatPainter(Boolean(locked))}
      onMergeCells={() => session.requestMergeCells()}
      canExecute={session.canExecute.bind(session)}
    />
  );
}
