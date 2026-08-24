import { lazy, Suspense, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { AppShell, Box, Button, DropdownMenu, Inline, Stack } from "@react-sheets/ui-system";
import { FormulaBar } from "../components/FormulaBar";
import { SheetTabs } from "../components/SheetTabs";
import { StatusBar } from "../components/StatusBar";
import { localeLabels, shellLabels, type Locale } from "../i18n";
import zhCN from "../locales/zh-CN.json";
import enUS from "../locales/en-US.json";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import type { UiSessionIntent, UiSnapshot, WorkbookSession } from "@react-sheets/spreadsheet-app";
import type { SelectionState } from "@react-sheets/spreadsheet-app";
import type { EditorCommandController } from "./command-controller";
import { RibbonHost } from "./RibbonHost";
import { FeaturePanelHost } from "./FeaturePanelHost";
import { EditorDialogHost } from "./EditorDialogHost";
import { ColumnDimensionController } from './column-dimension-controller';
import { ColumnWidthDialog } from '../components/dialogs/ColumnWidthDialog';

const SheetCanvas = lazy(() => import("../components/SheetCanvas").then((module) => ({ default: module.SheetCanvas })));

export interface EditorShellProps {
  state: UiSnapshot;
  session: WorkbookSession;
  locale: Locale;
  isBusy: boolean;
  sidebarOpen: boolean;
  onSidebarOpenChange: (open: boolean) => void;
  controller: EditorCommandController;
  dispatchCommand: (descriptor: CommandDescriptor) => void;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
  setLocale: (locale: Locale) => void;
  copyWorkbookLink: () => void;
  saveWorkbook: () => void;
  exportXlsx: () => void | Promise<void>;
  importXlsx: () => void;
  renameWorkbook: (name: string) => void | Promise<void>;
  onSetBackstageInfo: () => void;
  onOpenPrintPreview: () => void;
}

/**
 * Main editor surface. It owns visual composition (shell, canvas, feature
 * panels and dialogs); state transitions and domain payloads remain session or
 * command-controller owned.
 */
export function EditorShell({
  state,
  session,
  locale,
  isBusy,
  sidebarOpen,
  onSidebarOpenChange,
  controller,
  dispatchCommand,
  dispatchSessionIntent,
  setLocale,
  copyWorkbookLink,
  saveWorkbook,
  exportXlsx,
  importXlsx,
  renameWorkbook,
  onSetBackstageInfo,
  onOpenPrintPreview,
}: EditorShellProps): ReactNode {
  const sheetRef = useRef(state.selectedSheet);
  const selectionRef = useRef(state.selection);
  sheetRef.current = state.selectedSheet;
  selectionRef.current = state.selection;
  const columnDimensions = useMemo(
    () => new ColumnDimensionController(session, () => sheetRef.current, () => selectionRef.current),
    [session],
  );
  const [columnWidthDialog, setColumnWidthDialog] = useState<{ columns: number[]; defaultMode: boolean } | null>(null);
  useEffect(() => () => columnDimensions.cancelAutoFit(), [columnDimensions]);
  const selectedCellStyle = state.homeRibbon.style;
  const formatCellsInitial = {
    numberFormat: selectedCellStyle.numberFormat ?? "general",
    style: { ...selectedCellStyle },
  };

  const handleSelectionChange = (selection: SelectionState) => controller.applySelection(selection);

  return (
    <>
      <AppShell
        formulaBar={(
          <FormulaBar
            cellName={state.activeCell}
            disabled={isBusy}
            formula={state.formulaDraft}
            locale={locale}
            onCancel={session.cancelEdit.bind(session)}
            onChange={session.setFormulaDraft.bind(session)}
            onCommit={() => { if (state.editingCell) session.commitEdit("down"); else session.commitFormula(); }}
            onNameBoxCommit={(value) => session.selectAddress(value)}
            onOpenWizard={() => dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" })}
            phase={state.phase}
          />
        )}
        isBusy={isBusy}
        labels={shellLabels(locale, state.saveState)}
        localeMenuLabel={localeLabels[locale]}
        notice={state.notice}
        onLocaleChange={setLocale}
        onSearch={(query) => dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace", findQuery: query })}
        onShare={copyWorkbookLink}
        peers={state.peers}
        workbookMenu={(
          <DropdownMenu
            align="right"
            trigger={<Button aria-label="Open workbook menu" disabled={isBusy} icon="more-horizontal" iconOnly size="sm" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white" />}
          >
            {({ close }) => (
              <Stack gap="xs" className="min-w-44">
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { onSetBackstageInfo(); close(); }}>
                  File / 工作簿
                </Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => {
                  const nextName = window.prompt("Enter workbook name:", state.workbookName);
                  if (nextName?.trim()) void renameWorkbook(nextName.trim());
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
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onOpenPrintPreview(); }}>
                  Print / Save as PDF
                </Button>
              </Stack>
            )}
          </DropdownMenu>
        )}
        ribbon={(
          <RibbonHost
            state={state}
            session={session}
            locale={locale}
            isBusy={isBusy}
            dispatchCommand={dispatchCommand}
            dispatchSessionIntent={dispatchSessionIntent}
            saveWorkbook={saveWorkbook}
            exportXlsx={exportXlsx}
            importXlsx={importXlsx}
            commands={controller}
            columnDimensions={columnDimensions}
            onOpenColumnWidthDialog={(columns) => setColumnWidthDialog({ columns, defaultMode: false })}
            onOpenDefaultColumnWidthDialog={() => setColumnWidthDialog({ columns: columnDimensions.selectedColumns(), defaultMode: true })}
          />
        )}
        saveState={state.saveState}
        sheetTabs={(
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
        )}
        statusBar={(
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
        )}
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
                selectedFloatingId={state.selectedFloatingId}
                showFormulas={state.formulaAudit.showFormulas}
                onPivotContextHit={(hit) => {
                  const pivotId = hit?.pivot?.pivotId ?? hit?.objectId;
                  if (pivotId) {
                    session.setActivePivotContext(pivotId, state.activeSheetId);
                    controller.setActivePivotId(pivotId);
                    onSidebarOpenChange(true);
                    dispatchSessionIntent({ type: "panel.open", panel: "pivot" });
                  } else session.setActivePivotContext(null);
                }}
                getPivotContextMenuItems={(hit) => {
                  const pivotId = hit.pivot?.pivotId ?? hit.objectId;
                  if (!pivotId) return [];
                  const sourceRowPaths = hit.pivot?.sourceRowPaths ?? [];
                  return [
                    { id: "pivot-refresh", label: "Refresh PivotTable", onSelect: () => dispatchCommand({ commandId: "pivot.refresh", params: { sheetId: state.activeSheetId, pivotId } }) },
                    { id: "pivot-show-details", label: "Show Details", disabled: sourceRowPaths.length === 0, onSelect: () => session.showPivotDetails(pivotId, sourceRowPaths) },
                  ];
                }}
                onPivotShowDetails={({ pivotId, sourceRowPaths }) => session.showPivotDetails(pivotId, sourceRowPaths)}
                drawings={state.selectedSheet.drawings}
                drawingPayloads={state.selectedSheet.drawingPayloads}
                allSheets={state.sheets}
                pivotResults={state.selectedSheet.pivotResults}
                sparklines={state.selectedSheet.sparklines}
                onSelectionChange={handleSelectionChange}
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
                columnDimensions={columnDimensions}
                onOpenColumnWidthDialog={(columns) => setColumnWidthDialog({ columns, defaultMode: false })}
                onFillRange={session.fillRange.bind(session)}
                drawingSelectionMode={state.drawingSelectionMode}
                onExitDrawingSelectionMode={() => session.setDrawingSelectionMode(false)}
                onFloatingSelect={(hit, mode) => session.setDrawingSelection(hit ? [hit.id] : [], mode)}
                onFloatingMove={(drawingId, bounds, rotation) => dispatchCommand({ commandId: "drawing.move", params: { sheetId: state.activeSheetId, drawingId, transform: { ...bounds, rotation } } })}
                onFloatingRemove={(drawingId) => dispatchCommand({ commandId: "drawing.remove", params: { sheetId: state.activeSheetId, drawingId } })}
                onCommand={dispatchCommand}
                onClearSelection={(mode) => session.clearSelection(mode)}
                formatPainterActive={state.formatPainter !== null}
                onCancelFormatPainter={() => session.cancelFormatPainter()}
                onCopy={() => session.copy()}
                onCut={() => session.cut()}
                onPaste={() => session.paste()}
                onUndo={() => session.undo()}
                onRedo={() => session.redo()}
                onShortcut={controller.executeShortcut}
                canRepeat={session.canRepeatLastCommand()}
                onOpenInspector={() => dispatchSessionIntent({ type: "panel.open", panel: "inspector", notice: "Select a cell and use Review tools for comments." })}
                onApplyFilter={(column, patch) => session.applyFilter(column, patch)}
                onSortFilterColumn={(column, ascending) => session.sortFilterColumn(column, ascending)}
                onToggleOutline={(groupId) => session.toggleOutlineGroup(groupId)}
                getValidationList={session.getValidationAt.bind(session)}
                onRetry={session.retry.bind(session)}
                onCreateSheet={session.addSheet.bind(session)}
              />
            </Suspense>
          </Box>
          <FeaturePanelHost
            state={state}
            session={session}
            locale={locale}
            sidebarOpen={sidebarOpen}
            onSidebarOpenChange={onSidebarOpenChange}
            selectedRange={controller.selectedRange}
            dispatchCommand={dispatchCommand}
            dispatchSessionIntent={dispatchSessionIntent}
            commands={controller}
            title={locale === "zh-CN" ? zhCN.sidebar.title : enUS.sidebar.title}
          />
        </Inline>
      </AppShell>
      <EditorDialogHost
        state={state}
        session={session}
        locale={locale}
        sortColumns={controller.sortColumns}
        formatCellsInitial={formatCellsInitial}
        pivotSourceOptions={controller.pivotSourceOptions}
        createPivotFromDialog={controller.createPivotFromDialog}
      />
      <ColumnWidthDialog
        open={columnWidthDialog !== null}
        columnCount={columnWidthDialog?.columns.length ?? 0}
        defaultMode={columnWidthDialog?.defaultMode}
        maximumDigitWidthPx={state.selectedSheet.maximumDigitWidthPx}
        initialWidthPx={columnWidthDialog?.defaultMode ? state.selectedSheet.defaultColumnWidthPx : state.selectedSheet.columnWidthsPx[columnWidthDialog?.columns[0] ?? -1] ?? state.selectedSheet.defaultColumnWidthPx}
        onClose={() => setColumnWidthDialog(null)}
        onApply={(excelWidth) => {
          if (columnWidthDialog?.defaultMode) columnDimensions.setDefaultExcelWidth(excelWidth);
          else columnDimensions.setExcelWidth(columnWidthDialog?.columns ?? [], excelWidth);
        }}
      />
    </>
  );
}
