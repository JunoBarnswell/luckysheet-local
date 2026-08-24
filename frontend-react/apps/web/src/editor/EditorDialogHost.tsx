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
const PrintPreviewDialog = lazy(() => import("../components/dialogs/PrintPreviewDialog").then((module) => ({ default: module.PrintPreviewDialog })));

export interface EditorDialogHostProps {
  state: UiSnapshot;
  session: WorkbookSession;
  locale: Locale;
  sortColumns: UiSnapshot["selectedSheet"]["columns"];
  formatCellsInitial: { numberFormat: string; style: UiSnapshot["homeRibbon"]["style"] };
  pivotSourceOptions: EditorCommandController["pivotSourceOptions"];
  createPivotFromDialog: EditorCommandController["createPivotFromDialog"];
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
}: EditorDialogHostProps): ReactNode {
  return (
    <Suspense fallback={null}>
      <FunctionWizardDialog
        open={state.dialogs.active === 'function-wizard'}
        onClose={session.closeFunctionWizard.bind(session)}
        onInsertFormula={(formula) => { session.setFormulaDraft(formula); session.commitFormula(formula); }}
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
        locale={locale}
        onClose={session.closeFindReplace.bind(session)}
        onReplaceAll={(params) => session.findReplace(params)}
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
        onPaste={(mode) => session.pasteSpecial(mode)}
      />
      <FormatCellsDialog
        open={state.dialogs.active === 'format-cells'}
        initial={formatCellsInitial}
        locale={locale}
        onClose={session.closeFormatCells.bind(session)}
        onApply={(draft) => session.formatCells({ numberFormat: draft.numberFormat, style: draft.style })}
      />
      <ShiftCellsDialog
        open={state.dialogs.active === 'shift-cells'}
        locale={locale}
        onClose={session.closeShiftCells.bind(session)}
        onShift={(direction) => session.shiftCells(direction)}
      />
      <MergeConfirmDialog
        open={state.dialogs.active === 'merge-confirm'}
        discardedCellCount={state.dialogs.mergeDiscardCount}
        locale={locale}
        onCancel={() => session.cancelMergeCells()}
        onConfirm={() => session.confirmMergeCells()}
      />
      <CreatePivotTableDialog
        open={state.dialogs.active === 'create-pivot'}
        sourceRegion={session.getCurrentRegion()}
        sourceOptions={pivotSourceOptions.map(({ id, label }) => ({ id, label }))}
        activeSheetName={state.selectedSheet.name}
        onClose={session.closeCreatePivotDialog.bind(session)}
        onCreate={createPivotFromDialog}
      />
      <PrintPreviewDialog
        open={state.dialogs.active === 'print-preview'}
        onClose={() => session.setShowPrintPreview(false)}
        sheetId={state.activeSheetId}
        rowCount={state.selectedSheet.rowCount}
        columnCount={state.selectedSheet.columnCount}
        columns={state.selectedSheet.columns}
        rows={state.selectedSheet.previewRows}
        layout={state.printLayout}
        pages={state.printPages}
        getRow={(row) => state.selectedSheet.previewRows[row] ?? {
          rowNumber: row + 1,
          cells: Array.from({ length: state.selectedSheet.columnCount }, (_, column) => ({ value: state.selectedSheet.getCell(row, column)?.value ?? "" })),
        }}
      />
    </Suspense>
  );
}
