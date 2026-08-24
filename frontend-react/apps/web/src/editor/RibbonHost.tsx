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
    | "buildQuickChartCommand"
    | "buildQuickSparklineCommand"
    | "buildQuickShapeCommand"
    | "buildDrawingCommand"
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
  return (
    <Ribbon
      activeTab={state.ribbon.activeTab}
      activePivot={state.activeContext.kind === "pivot"
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
      onCreatePivot={() => session.buildQuickPivotDescriptor()}
      onCreateChart={commands.buildQuickChartCommand}
      onCreateSparkline={commands.buildQuickSparklineCommand}
      onCreateShape={commands.buildQuickShapeCommand}
      onBringDrawingForward={() => commands.buildDrawingCommand("drawing.zorder", "forward")}
      onSendDrawingBackward={() => commands.buildDrawingCommand("drawing.zorder", "backward")}
      onRemoveDrawing={() => commands.buildDrawingCommand("drawing.remove")}
      onCreateSheetTable={() => session.createSheetTableFromSelection()}
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
