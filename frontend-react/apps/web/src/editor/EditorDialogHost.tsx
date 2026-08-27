import { lazy, Suspense, type ReactNode } from "react";
import type { UiSnapshot, WorkbookSession } from "@react-sheets/spreadsheet-app";
import type { Locale } from "../i18n";
import type { EditorCommandController } from "./command-controller";

const FunctionWizardDialog = lazy(() => import("../components/dialogs/FunctionWizardDialog").then((module) => ({ default: module.FunctionWizardDialog })));
const SortDialog = lazy(() => import("../components/dialogs/SortDialog").then((module) => ({ default: module.SortDialog })));
const FindReplaceDialog = lazy(() => import("../components/dialogs/FindReplaceDialog").then((module) => ({ default: module.FindReplaceDialog })));
const GoToDialog = lazy(() => import("../components/dialogs/GoToDialog").then((module) => ({ default: module.GoToDialog })));
const PasteSpecialDialog = lazy(() => import("../components/dialogs/PasteSpecialDialog").then((module) => ({ default: module.PasteSpecialDialog })));
const FormatCellsDialog = lazy(() => import("../components/dialogs/FormatCellsDialog").then((module) => ({ default: module.FormatCellsDialog })));
const ShiftCellsDialog = lazy(() => import("../components/dialogs/ShiftCellsDialog").then((module) => ({ default: module.ShiftCellsDialog })));
const MergeConfirmDialog = lazy(() => import("../components/dialogs/MergeConfirmDialog").then((module) => ({ default: module.MergeConfirmDialog })));
const CreatePivotTableDialog = lazy(() => import("../components/dialogs/CreatePivotTableDialog").then((module) => ({ default: module.CreatePivotTableDialog })));
const CreateTableDialog = lazy(() => import("../components/dialogs/CreateTableDialog").then((module) => ({ default: module.CreateTableDialog })));
const PrintPreviewDialog = lazy(() => import("../components/dialogs/PrintPreviewDialog").then((module) => ({ default: module.PrintPreviewDialog })));
const CellTemplateDialog = lazy(() => import('../components/dialogs/CellTemplateDialog').then((module) => ({ default: module.CellTemplateDialog })));
const CellEditorDialog = lazy(() => import('../components/dialogs/CellEditorDialog').then((module) => ({ default: module.CellEditorDialog })));
const InsertPictureDialog = lazy(() => import('../components/dialogs/InsertPictureDialog').then((module) => ({ default: module.InsertPictureDialog })));
const HyperlinkDialog = lazy(() => import('../components/dialogs/HyperlinkDialog').then((module) => ({ default: module.HyperlinkDialog })));

export interface EditorDialogHostProps {
  state: UiSnapshot;
  session: WorkbookSession;
  locale: Locale;
  sortColumns: UiSnapshot["selectedSheet"]["columns"];
  formatCellsInitial: { numberFormat: string; style: UiSnapshot["homeRibbon"]["style"]; mixedFontFamily?: boolean };
  pivotSourceOptions: EditorCommandController["pivotSourceOptions"];
  createPivotFromDialog: EditorCommandController["createPivotFromDialog"];
  hyperlinkInitial?: import('@react-sheets/core-model').CellHyperlink;
  hyperlinkSheets: readonly import('../components/dialogs/HyperlinkDialog').HyperlinkSheetOption[];
}

/** Lazy dialog boundary. Dialog state remains session-owned; this host only maps typed callbacks. */
export function EditorDialogHost({
  state,
  session,
  locale,
  sortColumns,
  formatCellsInitial,
  pivotSourceOptions,
  createPivotFromDialog,
  hyperlinkInitial,
  hyperlinkSheets,
}: EditorDialogHostProps): ReactNode {
  return (
    <Suspense fallback={null}>
      <FunctionWizardDialog
        open={state.dialogs.active === 'function-wizard'}
        onClose={session.closeFunctionWizard.bind(session)}
        onInsertFormula={(formula) => {
          if (!state.editSession) session.beginEdit(undefined, 'functionInsert');
          session.setFormulaDraft(formula);
          session.commitEdit('none');
        }}
      />
      <SortDialog
        open={state.dialogs.active === 'sort-dialog'}
        columns={sortColumns}
        locale={locale}
        onClose={session.closeSortDialog.bind(session)}
        onSort={(criteria, hasHeader) => session.sortRange(criteria, hasHeader)}
      />
      <FindReplaceDialog
        open={state.dialogs.active === 'find-replace'}
        initialFind={state.dialogs.findQuery}
        mode={state.dialogs.findMode}
        locale={locale}
        onClose={session.closeFindReplace.bind(session)}
        onFindNext={(params) => session.findNext(params)}
        onFindPrevious={(params) => session.findPrevious(params)}
        onFindAll={(params) => session.findAll(params)}
        onReplace={(params) => session.replaceOne(params)}
        onReplaceAll={(params) => session.replaceAll(params)}
      />
      <GoToDialog
        open={state.dialogs.active === 'goto'}
        locale={locale}
        onClose={session.closeGoTo.bind(session)}
        onGoTo={(reference) => session.selectAddress(reference)}
        onGoToSpecial={(kind) => session.goToSpecial(kind)}
      />
      <PasteSpecialDialog
        open={state.dialogs.active === 'paste-special'}
        locale={locale}
        onClose={session.closePasteSpecial.bind(session)}
        onPaste={(spec) => session.pasteSpecial(spec)}
      />
      <FormatCellsDialog
        open={state.dialogs.active === 'format-cells'}
        initial={formatCellsInitial}
        locale={locale}
        onClose={session.closeFormatCells.bind(session)}
        onApply={(draft) => session.formatCells({ numberFormat: draft.numberFormat, style: draft.style, border: draft.border })}
      />
      <ShiftCellsDialog
        open={state.dialogs.active === 'shift-cells'}
        locale={locale}
        operation={state.dialogs.cellShiftOperation}
        onClose={session.closeShiftCells.bind(session)}
        onShift={(axis) => session.applyCellShift(state.dialogs.cellShiftOperation, axis)}
      />
      <MergeConfirmDialog
        open={state.dialogs.active === 'merge-confirm'}
        discardedCellCount={state.dialogs.mergeDiscardCount}
        operation={state.dialogs.mergeOperation}
        locale={locale}
        onCancel={() => session.cancelMergeAction()}
        onConfirm={() => session.confirmMergeAction()}
      />
      <CreatePivotTableDialog
        open={state.dialogs.active === 'create-pivot'}
        sourceRegion={session.getCurrentRegion()}
        sourceOptions={pivotSourceOptions.map(({ id, label }) => ({ id, label }))}
        activeSheetName={state.selectedSheet.name}
        locale={locale}
        onClose={session.closeCreatePivotDialog.bind(session)}
        onCreate={createPivotFromDialog}
      />
      <CreateTableDialog
        open={state.dialogs.active === 'create-table'}
        locale={locale}
        sourceRange={session.getPrimaryRange()}
        onClose={session.closeCreateTableDialog.bind(session)}
        onCreate={(request) => session.createSheetTableFromDialog(request)}
      />
      <PrintPreviewDialog
        open={state.dialogs.active === 'print-preview'}
        onClose={() => session.setShowPrintPreview(false)}
        sheetId={state.activeSheetId}
        rowCount={state.selectedSheet.rowCount}
        columnCount={state.selectedSheet.columnCount}
        columns={state.selectedSheet.columns}
        rows={[]}
        layout={state.printLayout}
        pages={state.printPages}
        getRow={(row) => state.selectedSheet.hiddenRows.includes(row) ? undefined : ({
          rowNumber: row + 1,
          cells: Array.from({ length: state.selectedSheet.columnCount }, (_, column) => ({ value: state.selectedSheet.getCell(row, column)?.value ?? "" })),
        })}
      />
      <CellTemplateDialog
        open={state.dialogs.active === 'cell-template'}
        templates={state.cellStyleTemplates}
        onClose={session.closeActiveDialog.bind(session)}
        onApply={(templateId) => { session.applyCellStyleTemplate(templateId); session.closeActiveDialog(); }}
        onRemove={(templateId) => session.removeCellStyleTemplate(templateId)}
        onSave={(template) => session.setCellStyleTemplate(template)}
      />
      <CellEditorDialog
        open={state.dialogs.active === 'cell-editor'}
        onClose={session.closeActiveDialog.bind(session)}
        onApply={(editor) => { session.setCellEditor(editor); session.closeActiveDialog(); }}
      />
      <InsertPictureDialog
        open={state.dialogs.active === 'insert-picture'}
        onClose={session.closeActiveDialog.bind(session)}
        onInsert={(file, placement) => session.insertImageFile(file, placement)}
      />
      <HyperlinkDialog
        open={state.dialogs.active === 'hyperlink'}
        initial={hyperlinkInitial}
        sheets={hyperlinkSheets}
        definedNames={state.definedNameModels}
        onClose={session.closeActiveDialog.bind(session)}
        onApply={(target, tooltip) => { session.setActiveHyperlink(target, tooltip); session.closeActiveDialog(); }}
        onRemove={session.removeHyperlink.bind(session)}
      />
    </Suspense>
  );
}
