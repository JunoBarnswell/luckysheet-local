import Store from "../../store";
import { drawPageCanvas } from "./printRenderer";
import { buildPages, currentFile, layoutOf, renderOf, resolvePrintRange } from "./printManager";

export function canvasToDataUrl(canvas, type, quality) {
    if (!canvas) {
        return "";
    }
    return canvas.toDataURL(type || "image/png", quality == null ? 0.92 : quality);
}

export function getScreenshot(range) {
    const file = currentFile();
    if (!file) {
        return Promise.resolve(null);
    }
    const layout = layoutOf(file, { area: range ? "CurrentSelection" : undefined });
    const render = renderOf(file);
    if (range && range.row && range.column) {
        layout.area = "CurrentSelection";
    }
    const pack = buildPages(file, layout, render);
    const page = pack.pages[0] || { row: resolvePrintRange(layout.area, file).row, column: resolvePrintRange(layout.area, file).column };
    const meta = { pageIndex: 0, pageTotal: pack.pages.length, sheetPage: 1, sheetPageTotal: pack.pages.length };
    const canvas = drawPageCanvas(page, file, layout, render, pack, meta);
    return Promise.resolve({
        dataUrl: canvasToDataUrl(canvas),
        width: canvas.width,
        height: canvas.height,
        pageCount: pack.pages.length,
    });
}

export function saveScreenshotToClipboard(range) {
    return getScreenshot(range).then(function (shot) {
        if (!shot || !shot.dataUrl) {
            return { ok: false, reason: "empty" };
        }
        if (typeof navigator === "undefined" || !navigator.clipboard || typeof window.ClipboardItem !== "function") {
            return { ok: false, dataUrl: shot.dataUrl, reason: "clipboard_unsupported" };
        }
        return fetch(shot.dataUrl)
            .then(function (res) {
                return res.blob();
            })
            .then(function (blob) {
                return navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);
            })
            .then(function () {
                return { ok: true, dataUrl: shot.dataUrl };
            })
            .catch(function () {
                return { ok: false, dataUrl: shot.dataUrl, reason: "clipboard_denied" };
            });
    });
}

export function getRangeScreenshot(range) {
    return getScreenshot(range);
}
