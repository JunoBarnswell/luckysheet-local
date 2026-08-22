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
import { exportPdf } from "./printPdf";
import { saveScreenshotToClipboard, getScreenshot } from "./printScreenshot";
import { drawPageCanvas } from "./printRenderer";
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
    $dlg.on("click.lsPrint", "#luckysheet-print-pdf-btn", function () {
        exportPrintPdf();
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
    const pair =
        options && (options.layout || options.render)
            ? {
                  layout: layoutOf(file, options.layout),
                  render: renderOf(file, options.render),
              }
            : readDialogLayout(file);
    persist(file, pair.layout, pair.render);
    return preparePrint(file, pair.layout, pair.render).then(function () {
        const pack = buildPages(file, pair.layout, pair.render);
        mountPreviewPages(pack, file, pair.layout, pair.render, instanceAttr());
        const style = createPrintStyle(pack.paper.pageW, pack.paper.pageH, pair.layout.direction);
        document.head.appendChild(style);
        if (typeof window !== "undefined" && typeof window.print === "function") {
            window.print();
        }
        emitAfterSheetPrintConfirm({ pageCount: pack.pages.length, range: pack.range });
        return { pageCount: pack.pages.length, range: pack.range };
    });
}

function exportPrintPdf(filename) {
    const file = currentFile();
    if (!file) {
        return Promise.resolve({ pageCount: 0 });
    }
    const pair = readDialogLayout(file);
    persist(file, pair.layout, pair.render);
    return preparePrint(file, pair.layout, pair.render).then(function () {
        const pack = buildPages(file, pair.layout, pair.render);
        return exportPdf(pack, pair.layout, pair.render, drawPageCanvas, filename);
    });
}

export const luckysheetPrint = {
    createDialog: createDialog,
    closeDialog: closeDialog,
    init: init,
    print: print,
    exportPrintPdf: exportPrintPdf,
    updatePrintConfig: updatePrintConfig,
    updatePrintRenderConfig: updatePrintRenderConfig,
    setPrintTitles: setPrintTitles,
    resolvePrintRange: resolvePrintRange,
    collectPrintRange: collectPrintRange,
    paginateByPaper: paginateByPaper,
    saveScreenshotToClipboard: saveScreenshotToClipboard,
    getScreenshot: getScreenshot,
    buildPages: buildPages,
    buildAllPageCanvases: buildAllPageCanvases,
};

export { resolvePrintRange, collectPrintRange };
export default luckysheetPrint;
