import { lazy, Suspense, useEffect, useMemo, useRef, type ReactNode } from "react";
import { DesignerShell, Box, Inline } from "@react-sheets/ui-system";
import { FormulaBar } from "../components/FormulaBar";
import { SheetTabs } from "../components/SheetTabs";
import { StatusBar } from "../components/StatusBar";
import { type Locale } from "../i18n";
import zhCN from "../locales/zh-CN.json";
import enUS from "../locales/en-US.json";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { selectedHeaderIndices, type ChartElementSelection, type UiSessionIntent, type UiSnapshot, type WorkbookSession } from "@react-sheets/spreadsheet-app";
import type { SelectionState } from "@react-sheets/spreadsheet-app";
import type { EditorCommandController } from "./command-controller";
import { RibbonHost } from "./RibbonHost";
import { FeaturePanelHost } from "./FeaturePanelHost";
import { EditorDialogHost } from "./EditorDialogHost";
import { ColumnDimensionController } from './column-dimension-controller';
import { ColumnWidthDialog } from '../components/dialogs/ColumnWidthDialog';
import { RowHeightDialog } from '../components/dialogs/RowHeightDialog';
import { buildPivotTimelineTiles, pivotMemberKey } from '@react-sheets/core-model';
import type { PivotControlAction } from '../components/canvas/drawing-renderers';

const SheetCanvas = lazy(() => import("../components/SheetCanvas").then((module) => ({ default: module.SheetCanvas })));

export interface EditorShellProps {
  state: UiSnapshot;
  session: WorkbookSession;
  locale: Locale;
  isBusy: boolean;
  controller: EditorCommandController;
  dispatchCommand: (descriptor: CommandDescriptor) => void;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
  setLocale: (locale: Locale) => void;
  copyWorkbookLink: () => void;
  saveWorkbook: () => void;
  exportDocument: () => void | Promise<void>;
  importDocument: () => void;
  renameWorkbook: (name: string) => void | Promise<void>;
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
  controller,
  dispatchCommand,
  dispatchSessionIntent,
  setLocale,
  copyWorkbookLink,
  saveWorkbook,
  exportDocument,
  importDocument,
  renameWorkbook,
  onOpenPrintPreview,
}: EditorShellProps): ReactNode {
  const sheetRef = useRef(state.selectedSheet);
  sheetRef.current = state.selectedSheet;
  const columnDimensions = useMemo(
    () => new ColumnDimensionController(session, () => sheetRef.current),
    [session],
  );
  useEffect(() => () => columnDimensions.cancelAutoFit(), [columnDimensions]);
  const dimensionBounds = { rowCount: state.selectedSheet.rowCount, columnCount: state.selectedSheet.columnCount };
  const selectedColumns = selectedHeaderIndices(state.selection, 'column', dimensionBounds, { includeOrdinaryCellRanges: true });
  const selectedRows = selectedHeaderIndices(state.selection, 'row', dimensionBounds, { includeOrdinaryCellRanges: true });
  const selectedCellStyle = state.homeRibbon.style;
  const formatCellsInitial = {
    numberFormat: selectedCellStyle.numberFormat ?? "general",
    style: { ...selectedCellStyle },
    mixedFontFamily: state.homeRibbon.mixedStyleKeys.includes('fontFamily'),
  };

  const handleSelectionChange = (selection: SelectionState) => controller.applySelection(selection);

  return (
    <>
      <DesignerShell
        formulaBar={(
          <FormulaBar
            cellName={state.activeCell}
            cellEdit={session.cellEdit}
            disabled={isBusy}
            readyFormula={state.formulaDraft}
            locale={locale}
            onCommitReady={() => session.commitFormula()}
            onNameBoxCommit={(value) => session.selectAddress(value)}
            onOpenNameManager={() => dispatchSessionIntent({ type: "panel.open", panel: "definedNames" })}
            onOpenWizard={() => dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" })}
            phase={state.phase}
          />
        )}
        formulaBarVisible={state.formulaBarVisible}
        ribbonVisible={state.ribbonVisible}
        isBusy={isBusy}
        ribbon={(
          <RibbonHost
            state={state}
            session={session}
            locale={locale}
            isBusy={isBusy}
            dispatchCommand={dispatchCommand}
            dispatchSessionIntent={dispatchSessionIntent}
            saveWorkbook={saveWorkbook}
            exportDocument={exportDocument}
            importDocument={importDocument}
            commands={controller}
            columnDimensions={columnDimensions}
            selectedColumns={selectedColumns}
            selectedRows={selectedRows}
            onOpenColumnWidthDialog={(columns) => dispatchSessionIntent({ type: "dialog.open", dialog: "column-width", columnWidth: { columns, defaultMode: false } })}
            onOpenDefaultColumnWidthDialog={() => dispatchSessionIntent({ type: "dialog.open", dialog: "column-width", columnWidth: { columns: selectedColumns, defaultMode: true } })}
            onOpenRowHeightDialog={(rows) => dispatchSessionIntent({ type: "dialog.open", dialog: "row-height", rowHeight: { rows } })}
          />
        )}
        sheetTabs={(
          <SheetTabs
            activeSheetId={state.activeSheetId}
            groupedSheetIds={state.groupedSheetIds}
            locale={locale}
            disabled={isBusy}
            onAdd={session.addSheet.bind(session)}
            onSelect={(sheetId, options) => session.selectSheet(sheetId, options)}
            onRenameSheet={session.renameSheet.bind(session)}
            onDeleteSheet={session.deleteSheet.bind(session)}
            onDuplicateSheet={session.duplicateSheet.bind(session)}
            onHideSheet={session.hideSheet.bind(session)}
            onSetTabColor={session.setSheetTabColor.bind(session)}
            onMoveSheet={session.moveSheet.bind(session)}
            dialog={state.dialogs.sheet}
            onOpenDialog={(sheetDialog) => dispatchSessionIntent({ type: "dialog.open", dialog: sheetDialog.kind === "rename" ? "sheet-rename" : sheetDialog.kind === "tab-color" ? "sheet-tab-color" : "sheet-delete", sheet: sheetDialog })}
            onUpdateDialog={(value) => dispatchSessionIntent({ type: "dialog.update", value })}
            onCloseDialog={session.closeActiveDialog.bind(session)}
            sheets={state.sheets}
          />
        )}
        statusBar={(
          <StatusBar
            activeCell={state.activeCell}
            cellEdit={session.cellEdit}
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
            fixedDecimalPlaces={state.editingOptions.fixedDecimalPlaces}
          />
        )}
        workspacePhase={state.phase}
      >
        <Inline gap="none" className="h-full min-h-0 w-full flex-nowrap">
          <Box className="h-full min-h-0 min-w-0 flex-1">
            <Suspense fallback={<Box className="h-full min-h-0 w-full bg-canvas" />}>
              <SheetCanvas
                locale={locale}
                sheet={state.selectedSheet}
                sheetId={state.activeSheetId}
                selection={state.selection}
                activeCell={state.activeCell}
                cellEdit={session.cellEdit}
                phase={state.phase}
                zoom={state.zoom}
                peers={state.peers}
                selectedFloatingId={state.selectedFloatingId}
                textBoxPlacementActive={state.textBoxPlacement}
                textBoxEdit={state.textBoxEdit}
                showFormulas={state.formulaAudit.showFormulas}
                onPivotContextHit={(hit) => {
                  const pivotId = hit?.pivot?.pivotId ?? hit?.objectId;
                  if (pivotId) {
                    session.setActivePivotContext(pivotId, state.activeSheetId);
                    controller.setActivePivotId(pivotId);
                    session.setPanelOpen(true);
                    dispatchSessionIntent({ type: "panel.open", panel: "pivot" });
                  } else session.setActivePivotContext(null);
                }}
                onPivotControlAction={(drawingId, action: PivotControlAction) => {
                  const drawing = state.selectedSheet.drawings.find((entry) => entry.id === drawingId);
                  const payload = drawing ? state.selectedSheet.drawingPayloads.get(drawing.payloadId) : undefined;
                  if (payload?.kind === 'slicer') {
                    if (action.kind === 'slicer-clear') {
                      session.setPivotSlicerFilter(drawingId, 'all', []);
                    } else if (action.kind === 'slicer-member') {
                      const existing = payload.filter.memberKeys;
                      if (payload.settings.multiSelect === false) {
                        session.setPivotSlicerFilter(drawingId, 'include', [action.memberKey]);
                      } else if (payload.filter.mode === 'exclude') {
                        const currentlySelected = !existing.some((entry) => pivotMemberKey(entry) === pivotMemberKey(action.memberKey));
                        const nextExcluded = currentlySelected
                          ? [...existing, action.memberKey]
                          : existing.filter((entry) => pivotMemberKey(entry) !== pivotMemberKey(action.memberKey));
                        session.setPivotSlicerFilter(drawingId, nextExcluded.length > 0 ? 'exclude' : 'all', nextExcluded);
                      } else {
                        const selected = payload.filter.mode === 'include' && existing.some((entry) => pivotMemberKey(entry) === pivotMemberKey(action.memberKey));
                        const memberKeys = selected
                          ? existing.filter((entry) => pivotMemberKey(entry) !== pivotMemberKey(action.memberKey))
                          : [...existing, action.memberKey];
                        session.setPivotSlicerFilter(drawingId, memberKeys.length > 0 ? 'include' : 'all', memberKeys);
                      }
                    }
                    return;
                  }
                  if (payload?.kind !== 'timeline') return;
                  const tree = state.selectedSheet.pivotResults[payload.pivotId];
                  const values = tree?.fields.fields.find((entry) => entry.fieldId === payload.fieldId)?.values ?? [];
                  const periods = buildPivotTimelineTiles(values, payload.level)
                    .filter((period) => (!payload.bounds.start || period.end >= payload.bounds.start) && (!payload.bounds.end || period.start <= payload.bounds.end));
                  if (action.kind === 'timeline-period') {
                    session.setPivotTimelinePeriod(drawingId, action.start, action.end);
                  } else if (action.kind === 'timeline-level') {
                    session.setPivotTimelineLevel(drawingId, action.level);
                  } else if (periods.length > 0 && (action.kind === 'timeline-scroll' || action.kind === 'timeline-handle')) {
                    if (action.kind === 'timeline-scroll') {
                      const currentStart = payload.scrollPosition ? Math.max(0, periods.findIndex((period) => period.start >= payload.scrollPosition!)) : 0;
                      const width = Math.min(8, periods.length);
                      const nextStart = Math.max(0, Math.min(periods.length - width, currentStart + action.direction * width));
                      const nextPosition = periods[nextStart]?.start;
                      if (nextPosition) session.setPivotTimelineWindow(drawingId, nextPosition);
                    } else if (action.edge === 'start') {
                      const currentStart = payload.period.start ? Math.max(0, periods.findIndex((period) => period.start >= payload.period.start!)) : 0;
                      const currentEnd = payload.period.end ? Math.max(currentStart, periods.findIndex((period) => period.end >= payload.period.end!)) : currentStart;
                      const nextStart = Math.max(0, Math.min(currentEnd, currentStart - 1));
                      session.setPivotTimelinePeriod(drawingId, periods[nextStart]?.start, periods[Math.max(nextStart, currentEnd)]?.end);
                    } else {
                      const startIndex = payload.period.start ? Math.max(0, periods.findIndex((period) => period.start >= payload.period.start!)) : 0;
                      const currentEnd = payload.period.end ? Math.max(startIndex, periods.findIndex((period) => period.end >= payload.period.end!)) : startIndex;
                      const nextEnd = Math.min(periods.length - 1, currentEnd + 1);
                      session.setPivotTimelinePeriod(drawingId, periods[startIndex]?.start, periods[nextEnd]?.end);
                    }
                  }
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
                onPivotExpansionToggle={(pivotId, nodeId) => { void session.togglePivotExpansion(pivotId, nodeId); }}
                onApplyPivotFilter={controller.applyPivotHeaderFilter}
                drawings={state.selectedSheet.drawings}
                drawingPayloads={state.selectedSheet.drawingPayloads}
                allSheets={state.projectionSheets}
                pivotResults={state.selectedSheet.pivotResults}
                sparklines={state.selectedSheet.sparklines}
                tables={state.tables}
                onSelectionChange={handleSelectionChange}
                onExtendSelection={(row, column) => session.extendSelectionTo(row, column)}
                onMovePrimary={(rowDelta, columnDelta, opts) => session.movePrimary(rowDelta, columnDelta, opts)}
                onEnsureSheetExtent={(rowCount, columnCount) => session.ensureSheetExtent(rowCount, columnCount)}
                onJumpEdge={(direction, extend) => session.jumpEdge(direction, extend)}
                onSelectAll={session.selectAll.bind(session)}
                onSelectAllDrawings={session.selectAllDrawings.bind(session)}
                onCycleDrawingSelection={session.cycleDrawingSelection.bind(session)}
                columnDimensions={columnDimensions}
                onOpenColumnWidthDialog={(columns) => dispatchSessionIntent({ type: "dialog.open", dialog: "column-width", columnWidth: { columns, defaultMode: false } })}
                onOpenFormatCells={() => dispatchSessionIntent({ type: "dialog.open", dialog: "format-cells" })}
                onOpenRowHeightDialog={(rows) => dispatchSessionIntent({ type: "dialog.open", dialog: "row-height", rowHeight: { rows } })}
                onFillRange={session.fillRange.bind(session)}
                onRangeDragCommit={(sourceRange, targetOrigin, mode) => session.rangeDrag(sourceRange, targetOrigin, mode)}
                drawingSelectionMode={state.drawingSelectionMode}
                onExitDrawingSelectionMode={() => session.setDrawingSelectionMode(false)}
                onFloatingSelect={(hit, mode) => {
                  const drawing = hit ? state.selectedSheet.drawings.find((entry) => entry.id === hit.id) : undefined;
                  const payload = drawing ? state.selectedSheet.drawingPayloads.get(drawing.payloadId) : undefined;
                  if (hit && mode === 'replace' && !state.drawingSelectionMode) {
                    if (payload?.kind === 'form-control') {
                      dispatchCommand({ commandId: 'formControl.activate', params: { sheetId: state.activeSheetId, drawingId: hit.id } });
                      return;
                    }
                  }
                  if (payload?.kind === 'slicer' && mode === 'replace') {
                    session.setPanelOpen(true);
                    dispatchSessionIntent({ type: 'panel.open', panel: 'slicer' });
                  }
                  session.setDrawingSelection(hit ? [hit.id] : [], mode);
                }}
                onChartElementAction={(drawingId, data) => {
                  const drawing = state.selectedSheet.drawings.find((entry) => entry.id === drawingId && entry.kind === 'chart');
                  const payload = drawing ? state.selectedSheet.drawingPayloads.get(drawing.payloadId) : undefined;
                  if (!drawing || payload?.kind !== 'chart' || !data || typeof data !== 'object' || !('kind' in data)) return;
                  const kind = String((data as { kind: string }).kind);
                  if ((kind === 'point' || kind === 'series' || kind === 'data-label') && 'seriesId' in data) {
                    session.selectChartElement({ kind: kind as 'point' | 'series' | 'data-label', chartId: payload.chartId, seriesId: String((data as { seriesId: string }).seriesId), ...('pointIndex' in data ? { pointIndex: Number((data as { pointIndex: number }).pointIndex) } : {}) });
                  } else if (['chart-area', 'plot-area', 'title', 'legend', 'axis', 'axis-title', 'gridline', 'data-table', 'trendline', 'error-bar'].includes(kind)) {
                    session.selectChartElement({ kind: kind as Exclude<ChartElementSelection, { kind: 'series' | 'point' | 'data-label' }>['kind'], chartId: payload.chartId });
                  }
                }}
                onFloatingMove={(drawingId, bounds, rotation) => dispatchCommand({ commandId: "drawing.move", params: { sheetId: state.activeSheetId, drawingId, transform: { ...bounds, rotation } } })}
                onFloatingRemove={(drawingId) => dispatchCommand({ commandId: "drawing.remove", params: { sheetId: state.activeSheetId, drawingId } })}
                onTextBoxPlacementCommit={(bounds) => session.placeTextBox(bounds)}
                onCancelTextBoxPlacement={() => session.cancelTextBoxPlacement()}
                onBeginTextBoxEdit={(drawingId, initialText) => session.beginTextBoxEdit(drawingId, initialText)}
                onTextBoxDraftChange={(value) => session.setTextBoxDraft(value)}
                onCommitTextBoxEdit={() => session.commitTextBoxEdit()}
                onCancelTextBoxEdit={() => session.cancelTextBoxEdit()}
                onCommand={dispatchCommand}
                onClearSelection={(mode) => session.clearSelection(mode)}
                formatPainterActive={state.formatPainter !== null}
                onCancelFormatPainter={() => session.cancelFormatPainter()}
                onCopy={() => session.copy()}
                onCut={() => session.cut()}
                onPaste={() => session.paste()}
                onPasteSpecial={() => dispatchSessionIntent({ type: "dialog.open", dialog: "paste-special" })}
                onUndo={() => session.undo()}
                onRedo={() => session.redo()}
                onShortcut={controller.executeShortcut}
                canRepeat={session.canRepeatLastCommand()}
                onOpenInspector={() => dispatchSessionIntent({ type: "panel.open", panel: "inspector", notice: "Select a cell and use Review tools for comments." })}
                onOpenHyperlink={() => dispatchSessionIntent({ type: "dialog.open", dialog: "hyperlink" })}
                onRemoveHyperlink={session.removeHyperlink.bind(session)}
                hasActiveHyperlink={Boolean(session.getActiveHyperlink())}
                onApplyFilter={(column, patch) => session.applyFilter(column, patch)}
                onSortFilterColumn={(column, ascending) => session.sortFilterColumn(column, ascending)}
                onToggleOutline={(groupId) => session.toggleOutlineGroup(groupId)}
                onRetry={session.retry.bind(session)}
                onCreateSheet={session.addSheet.bind(session)}
                resolveAssetUrl={session.resolveAssetUrl.bind(session)}
              />
            </Suspense>
          </Box>
          <FeaturePanelHost
            state={state}
            session={session}
            locale={locale}
            sidebarOpen={state.panels.open}
            onSidebarOpenChange={session.setPanelOpen.bind(session)}
            selectedRange={controller.selectedRange}
            dispatchCommand={dispatchCommand}
            dispatchSessionIntent={dispatchSessionIntent}
            commands={controller}
            title={locale === "zh-CN" ? zhCN.sidebar.title : enUS.sidebar.title}
          />
        </Inline>
      </DesignerShell>
      <EditorDialogHost
        state={state}
        session={session}
        locale={locale}
        sortColumns={controller.sortColumns}
        formatCellsInitial={formatCellsInitial}
        pivotSourceOptions={controller.pivotSourceOptions}
        createPivotFromDialog={controller.createPivotFromDialog}
        hyperlinkInitial={session.getActiveHyperlink()}
        hyperlinkSheets={session.getSheetOptions()}
      />
      <ColumnWidthDialog
        open={state.dialogs.active === 'column-width'}
        columnCount={state.dialogs.columnWidth?.columns.length ?? 0}
        defaultMode={state.dialogs.columnWidth?.defaultMode}
        maximumDigitWidthPx={state.selectedSheet.maximumDigitWidthPx}
        initialWidthPx={state.dialogs.columnWidth?.defaultMode ? state.selectedSheet.defaultColumnWidthPx : state.selectedSheet.columnWidthsPx[state.dialogs.columnWidth?.columns[0] ?? -1] ?? state.selectedSheet.defaultColumnWidthPx}
        onClose={() => session.closeActiveDialog()}
        onApply={(excelWidth) => {
          if (state.dialogs.columnWidth?.defaultMode) columnDimensions.setDefaultExcelWidth(excelWidth);
          else columnDimensions.setExcelWidth(state.dialogs.columnWidth?.columns ?? [], excelWidth);
          session.closeActiveDialog();
        }}
      />
      <RowHeightDialog
        open={state.dialogs.active === 'row-height'}
        rowCount={state.dialogs.rowHeight?.rows.length ?? 0}
        initialHeightPx={state.selectedSheet.rowHeightsPx[state.dialogs.rowHeight?.rows[0] ?? -1] ?? state.selectedSheet.defaultRowHeightPx}
        onClose={() => session.closeActiveDialog()}
        onApply={(points) => {
          columnDimensions.setRowHeightPoints(state.dialogs.rowHeight?.rows ?? [], points);
          session.closeActiveDialog();
        }}
      />
    </>
  );
}
