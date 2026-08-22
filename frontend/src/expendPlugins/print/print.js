import Store from "../../store";
import {
    collectPrintRange,
    resolvePrintRange,
    buildPages,
    preparePrint,
    updatePrintConfig,
    updatePrintRenderConfig,
    setPrintTitles,
    currentFile,
    layoutOf,
    renderOf,
    persist,
    buildAllPageCanvases,
    createPreparedPrintSession,
} from "./printManager";
import { paginateByPaper } from "./printLayout";
import {
    mountDialog,
    closeDialogDom,
    readDialogLayout,
    renderPreview,
    mountPreviewPages,
    DIALOG_ID,
} from "./printDialog";
import { ensurePrintStyleTag, createPrintStyle } from "./printBrowser";
import { runInPrintSession } from "./printSession";
import {
    emitBeforeSheetPrintOpen,
    emitAfterSheetPrintOpen,
    emitBeforeSheetPrintConfirm,
    emitAfterSheetPrintConfirm,
    emitBeforeSheetPrintCanceled,
} from "./printEvents";

function instanceAttr() {
    return Store.instanceId || "default";
}

function createDialog() {
    if (!emitBeforeSheetPrintOpen({ workbook: Store.luckysheetfile })) {
        return;
    }
    mountDialog(instanceAttr());
    renderPreview();
    emitAfterSheetPrintOpen({ workbook: Store.luckysheetfile });
}

function closeDialog() {
    emitBeforeSheetPrintCanceled({});
    closeDialogDom();
}

function init(license) {
    ensurePrintStyleTag();
    const $dlg = $("#" + DIALOG_ID);
    if (!$dlg.length) {
        return;
    }
    $dlg.off(".lsPrint");
    $dlg.on("change.lsPrint", "select,input", function () {
        renderPreview();
    });
    $dlg.on("click.lsPrint", "#luckysheet-print-preview-btn", function () {
        renderPreview();
    });
    $dlg.on("click.lsPrint", "#luckysheet-print-do-btn", function () {
        print();
    });
    $dlg.on("click.lsPrint", ".luckysheet-modal-dialog-title-close,.luckysheet-model-close-btn", function () {
        closeDialog();
    });
    if (license) {
        $dlg.attr("data-license", "1");
    }
}

function print(options) {
    const file = currentFile();
    if (!file) {
        return Promise.resolve({ pageCount: 0 });
    }
    if (!emitBeforeSheetPrintConfirm({ file: file, options: options })) {
        return Promise.resolve({ pageCount: 0, canceled: true });
    }
    ensurePrintStyleTag();
    const pair = options
        ? {
              layout: layoutOf(file, options.layout || options),
              render: renderOf(file, options.render || options),
          }
        : readDialogLayout(file);
    persist(file, pair.layout, pair.render);
    return createPreparedPrintSession(pair.layout, pair.render).then(function (prepared) {
        return runInPrintSession(prepared.session, function () {
            const plan = prepared.plan;
            mountPreviewPages(plan, prepared.session.instanceId);
            const style = createPrintStyle(plan.paper.pageW, plan.paper.pageH, prepared.session.layout.direction, prepared.session.id);
            document.head.appendChild(style);
            const finish = function () {
                style.remove();
                window.removeEventListener("afterprint", finish);
                runInPrintSession(prepared.session, function () {
                    closeDialogDom();
                    emitAfterSheetPrintConfirm({
                        pageCount: plan.pageCount,
                        resource: prepared.resource,
                        instanceId: prepared.session.instanceId,
                    });
                });
            };
            window.addEventListener("afterprint", finish);
            if (typeof window !== "undefined" && typeof window.print === "function") {
                window.print();
            } else {
                finish();
            }
            return {
                pageCount: plan.pageCount,
                instanceId: prepared.session.instanceId,
                diagnostics: prepared.resource.diagnostics,
                resourceTimedOut: prepared.resource.timedOut,
            };
        });
    });
}

export const luckysheetPrint = {
    createDialog: createDialog,
    closeDialog: closeDialog,
    init: init,
    print: print,
    updatePrintConfig: updatePrintConfig,
    updatePrintRenderConfig: updatePrintRenderConfig,
    setPrintTitles: setPrintTitles,
    resolvePrintRange: resolvePrintRange,
    collectPrintRange: collectPrintRange,
    paginateByPaper: paginateByPaper,
    buildPages: buildPages,
    buildAllPageCanvases: buildAllPageCanvases,
};

export { resolvePrintRange, collectPrintRange };
export default luckysheetPrint;
