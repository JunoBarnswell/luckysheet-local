import Store from "../../store";
import locale from "../../locale/locale";
import formula from "../../global/formula";
import { getSheetIndex } from "../../methods/get";
import { modelHTML } from "../../controllers/constant";
import { replaceHtml, chatatABC } from "../../utils/util";
import {
    PrintArea,
    PrintPaperSize,
    PrintDirection,
    PrintScale,
    PrintPaperMargin,
    cellDisplayValue,
    usedRange,
    getPrintOptions,
    normalizeLayoutFromPrintoptions,
    normalizeRenderFromPrintoptions,
    writePrintoptions,
    resolvePaperPx,
    paginateByPaper,
    computeFitScale,
    rowStartPx,
    colStartPx,
    rowHeight,
    colWidth,
    clampScale,
} from "./printLayout";

const DIALOG_ID = "luckysheet-print-dialog";
const PREVIEW_ID = "luckysheet-print-preview";
const RESOURCE_WAIT_MS = 10000;

function currentFile() {
    const index = getSheetIndex(Store.currentSheetIndex);
    return Store.luckysheetfile && Store.luckysheetfile[index];
}

function ensureFileConfig(file) {
    if (!file.config) {
        file.config = {};
    }
    if (!file.config.printoptions) {
        file.config.printoptions = {};
    }
    return file.config;
}

function parseRangeText(text) {
    if (!text) {
        return null;
    }
    let raw = String(text).replace(/\$/g, "");
    if (raw.indexOf("!") > -1) {
        raw = raw.split("!").pop();
    }
    try {
        return formula.getcellrange(raw);
    } catch (e) {
        return null;
    }
}

function selectionRange() {
    const last = Store.luckysheet_select_save && Store.luckysheet_select_save[Store.luckysheet_select_save.length - 1];
    if (!last || last.row == null || last.column == null) {
        return usedRange(Store.flowdata);
    }
    return { row: last.row.slice(), column: last.column.slice() };
}

export function resolvePrintRange(area, file) {
    const sheet = file || currentFile();
    const po = getPrintOptions(sheet);
    const data = (sheet && sheet.data) || Store.flowdata || [];
    const mode = area || (po.PrintArea ? PrintArea.CurrentSheet : PrintArea.CurrentSheet);

    if (mode === PrintArea.CurrentSelection || mode === "selection") {
        return selectionRange();
    }
    if (mode === PrintArea.AllSelection) {
        const ranges = Store.luckysheet_select_save || [];
        if (!ranges.length) {
            return selectionRange();
        }
        let r0 = ranges[0].row[0], r1 = ranges[0].row[1], c0 = ranges[0].column[0], c1 = ranges[0].column[1];
        for (let i = 1; i < ranges.length; i++) {
            r0 = Math.min(r0, ranges[i].row[0]);
            r1 = Math.max(r1, ranges[i].row[1]);
            c0 = Math.min(c0, ranges[i].column[0]);
            c1 = Math.max(c1, ranges[i].column[1]);
        }
        return { row: [r0, r1], column: [c0, c1] };
    }
    if (po.PrintArea) {
        const parsed = parseRangeText(po.PrintArea);
        if (parsed && parsed.row && parsed.column) {
            return { row: parsed.row.slice(), column: parsed.column.slice() };
        }
    }
    return usedRange(data);
}

export function collectPrintRange(area) {
    return resolvePrintRange(area);
}

function layoutOf(file, overrides) {
    const po = getPrintOptions(file);
    const saved = (file && file.config && file.config.printLayout) || {};
    return normalizeLayoutFromPrintoptions(po, Object.assign({}, saved, overrides || {}));
}

function renderOf(file, overrides) {
    const po = getPrintOptions(file);
    const saved = (file && file.config && file.config.printRender) || {};
    return normalizeRenderFromPrintoptions(po, Object.assign({}, saved, overrides || {}));
}

function persist(file, layout, render) {
    const config = ensureFileConfig(file);
    config.printoptions = writePrintoptions(config.printoptions, layout, render);
    config.printLayout = layout;
    config.printRender = render;
}

function waitPrintResources(draft) {
    if (draft || typeof document === "undefined") {
        return Promise.resolve();
    }
    const root = Store.container ? document.getElementById(Store.container) : document;
    if (!root) {
        return Promise.resolve();
    }
    const canvases = root.querySelectorAll(".luckysheet-data-visualization-chart canvas, .luckysheet-modal-dialog-slider canvas");
    const pending = [];
    canvases.forEach(function (cv) {
        if (cv.width > 0 && cv.height > 0) {
            return;
        }
        pending.push(new Promise(function (resolve) {
            const start = Date.now();
            const tick = function () {
                if ((cv.width > 0 && cv.height > 0) || Date.now() - start > RESOURCE_WAIT_MS) {
                    resolve();
                    return;
                }
                setTimeout(tick, 50);
            };
            tick();
        }));
    });
    if (!pending.length) {
        return Promise.resolve();
    }
    return Promise.race([
        Promise.all(pending),
        new Promise(function (resolve) { setTimeout(resolve, RESOURCE_WAIT_MS); }),
    ]);
}

function buildPages(file, layout, render) {
    const range = resolvePrintRange(layout.area, file);
    const visibledatarow = Store.visibledatarow || [];
    const visibledatacolumn = Store.visibledatacolumn || [];
    const po = getPrintOptions(file);
    const paper = resolvePaperPx(layout, po);
    let scale = 1;
    if (layout.scale === PrintScale.Custom) {
        scale = clampScale(layout.customScale) / 100;
    } else if (layout.scale !== PrintScale.Origin) {
        scale = computeFitScale(range, visibledatarow, visibledatacolumn, paper, layout.scale);
    }
    const scaledPaper = {
        pageW: paper.pageW,
        pageH: paper.pageH,
        innerW: paper.innerW / scale,
        innerH: paper.innerH / scale,
        pad: paper.pad,
    };
    const pages = paginateByPaper(range, visibledatarow, visibledatacolumn, scaledPaper, layout);
    return { range: range, paper: paper, pages: pages, scale: scale, visibledatarow: visibledatarow, visibledatacolumn: visibledatacolumn };
}

function drawPageCanvas(page, file, layout, render, pack) {
    const dpr = Math.max(1, Store.devicePixelRatio || 1);
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(pack.paper.pageW * dpr);
    canvas.height = Math.ceil(pack.paper.pageH * dpr);
    canvas.style.width = pack.paper.pageW + "px";
    canvas.style.height = pack.paper.pageH + "px";
    const ctx = canvas.getContext("2d");
    ctx.scale(dpr, dpr);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, pack.paper.pageW, pack.paper.pageH);

    const ox = pack.paper.pad.left;
    const oy = pack.paper.pad.top;
    ctx.save();
    ctx.translate(ox, oy);
    ctx.scale(pack.scale, pack.scale);

    const data = (file && file.data) || Store.flowdata || [];
    const merge = ((file && file.config) || Store.config || {}).merge || {};
    const r0 = page.row[0];
    const r1 = page.row[1];
    const c0 = page.column[0];
    const c1 = page.column[1];
    const originX = colStartPx(pack.visibledatacolumn, c0);
    const originY = rowStartPx(pack.visibledatarow, r0);
    const headingW = render.headings ? 36 : 0;
    const headingH = render.headings ? 20 : 0;

    if (render.headings) {
        ctx.fillStyle = "#f3f4f6";
        ctx.fillRect(0, 0, headingW + colStartPx(pack.visibledatacolumn, c1 + 1) - originX, headingH);
        ctx.fillRect(0, 0, headingW, headingH + rowStartPx(pack.visibledatarow, r1 + 1) - originY);
        ctx.fillStyle = "#111827";
        ctx.font = "11px sans-serif";
        ctx.textBaseline = "middle";
        for (let c = c0; c <= c1; c++) {
            const x = headingW + colStartPx(pack.visibledatacolumn, c) - originX;
            const w = colWidth(pack.visibledatacolumn, c);
            ctx.fillText(chatatABC(c), x + 4, headingH / 2);
            if (w) { /* keep */ }
        }
        for (let r = r0; r <= r1; r++) {
            const y = headingH + rowStartPx(pack.visibledatarow, r) - originY;
            const h = rowHeight(pack.visibledatarow, r);
            ctx.fillText(String(r + 1), 4, y + h / 2);
        }
    }

    ctx.beginPath();
    ctx.rect(headingW, headingH, pack.paper.innerW / pack.scale, pack.paper.innerH / pack.scale);
    ctx.clip();

    for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
            const x = headingW + colStartPx(pack.visibledatacolumn, c) - originX;
            const y = headingH + rowStartPx(pack.visibledatarow, r) - originY;
            const w = colWidth(pack.visibledatacolumn, c);
            const h = rowHeight(pack.visibledatarow, r);
            const cell = data[r] && data[r][c];
            if (cell && cell.bg) {
                ctx.fillStyle = cell.bg;
                ctx.fillRect(x, y, w, h);
            }
            if (render.gridlines) {
                ctx.strokeStyle = "#d1d5db";
                ctx.lineWidth = 1;
                ctx.strokeRect(x + 0.5, y + 0.5, Math.max(w - 1, 0), Math.max(h - 1, 0));
            }
            const text = cellDisplayValue(cell);
            if (text) {
                ctx.save();
                ctx.beginPath();
                ctx.rect(x + 2, y + 1, Math.max(w - 4, 0), Math.max(h - 2, 0));
                ctx.clip();
                ctx.fillStyle = (cell && cell.fc) || "#111827";
                const fs = (cell && cell.fs) || 11;
                const italic = cell && cell.it ? "italic " : "";
                const bold = cell && cell.bl ? "bold " : "";
                ctx.font = italic + bold + fs + "px sans-serif";
                ctx.textBaseline = "middle";
                ctx.fillText(text, x + 3, y + h / 2);
                ctx.restore();
            }
        }
    }

    if (!render.draft && file && file.images) {
        Object.keys(file.images).forEach(function (id) {
            const img = file.images[id];
            if (!img || !img.default) {
                return;
            }
            const left = (img.default.left || 0) - originX + headingW;
            const top = (img.default.top || 0) - originY + headingH;
            const iw = img.default.width || 0;
            const ih = img.default.height || 0;
            if (iw <= 0 || ih <= 0) {
                return;
            }
            ctx.strokeStyle = "#9ca3af";
            ctx.strokeRect(left, top, iw, ih);
        });
    }

    ctx.restore();
    drawHeaderFooter(ctx, pack, render, file, page);
    return canvas;
}

function drawHeaderFooter(ctx, pack, render, file, page) {
    const setting = render.headerFooterSetting || {};
    const placeholders = render.headerFooter || [];
    const extras = placeholders.map(function (token) {
        if (token === "PageSize" || token === "PageNumber") {
            return String((pack.pages && pack.pages.indexOf(page) + 1) || 1);
        }
        if (token === "WorksheetTitle") {
            return file && file.name ? file.name : "";
        }
        if (token === "WorkbookTitle") {
            return (Store.toJsonOptions && Store.toJsonOptions.title) || "";
        }
        if (token === "Date") {
            return new Date().toLocaleDateString();
        }
        if (token === "Time") {
            return new Date().toLocaleTimeString();
        }
        return "";
    }).filter(Boolean).join("  ");
    ctx.fillStyle = "#4b5563";
    ctx.font = "10px sans-serif";
    ctx.textBaseline = "top";
    const top = [setting.topLeft, setting.topCenter || extras, setting.topRight].filter(Boolean).join("   ");
    const bottom = [setting.bottomLeft, setting.bottomCenter, setting.bottomRight].filter(Boolean).join("   ");
    if (top) {
        ctx.fillText(top, pack.paper.pad.left, 8);
    }
    if (bottom) {
        ctx.fillText(bottom, pack.paper.pad.left, pack.paper.pageH - 18);
    }
}

function dialogHtml() {
    const _locale = locale();
    const p = _locale.print || {};
    const button = _locale.button || {};
    const papers = ["A4", "Letter", "Legal", "A3", "A5", "Tabloid", "B5"];
    const paperOpts = papers.map(function (name) {
        return '<option value="' + name + '">' + name + "</option>";
    }).join("");
    return (
        '<div class="luckysheet-print" id="print-layout-options">' +
            '<div class="luckysheet-print-title">' + (p.title || "打印设置") + "</div>" +
            '<div class="luckysheet-print-row"><label>' + (p.range || "打印范围") + "</label>" +
                '<select id="luckysheet-print-area">' +
                    '<option value="' + PrintArea.CurrentSheet + '">' + (p.current || "当前工作表") + "</option>" +
                    '<option value="' + PrintArea.CurrentSelection + '">' + (p.area || "选中区域") + "</option>" +
                "</select></div>" +
            '<div class="luckysheet-print-row"><label>' + (p.size || "纸张大小") + "</label>" +
                '<select id="luckysheet-print-paper">' + paperOpts + "</select></div>" +
            '<div class="luckysheet-print-row"><label>' + (p.direction || "打印方向") + "</label>" +
                '<div class="luckysheet-print-radio">' +
                    '<div><label><input type="radio" name="ls-print-dir" value="' + PrintDirection.Portrait + '"/> ' + (p.vertical || "纵向") + "</label></div>" +
                    '<div><label><input type="radio" name="ls-print-dir" value="' + PrintDirection.Landscape + '"/> ' + (p.horizontal || "横向") + "</label></div>" +
                "</div></div>" +
            '<div class="luckysheet-print-row"><label>' + (p.showLine || "显示网格线") + "</label>" +
                '<label><input type="checkbox" id="luckysheet-print-grid" checked/> ' + (p.show || "显示") + "</label></div>" +
            '<div class="luckysheet-print-suggest">' + (p.suggest || "") + "</div>" +
        "</div>" +
        '<div class="luckysheet-print-box" id="luckysheet-print-box"></div>'
    );
}

function instanceAttr() {
    return Store.instanceId || "default";
}

function scopedDialogId() {
    return DIALOG_ID;
}

function readDialogLayout(file) {
    const layout = layoutOf(file);
    const render = renderOf(file);
    const $dlg = $("#" + scopedDialogId());
    if ($dlg.length) {
        layout.area = $dlg.find("#luckysheet-print-area").val() || layout.area;
        layout.paperSize = $dlg.find("#luckysheet-print-paper").val() || layout.paperSize;
        const dir = $dlg.find('input[name="ls-print-dir"]:checked').val();
        if (dir) {
            layout.direction = dir;
        }
        render.gridlines = $dlg.find("#luckysheet-print-grid").prop("checked");
    }
    return { layout: layout, render: render };
}

function fillDialog(file) {
    const layout = layoutOf(file);
    const render = renderOf(file);
    const $dlg = $("#" + scopedDialogId());
    $dlg.find("#luckysheet-print-area").val(layout.area);
    $dlg.find("#luckysheet-print-paper").val(layout.paperSize || PrintPaperSize.A4);
    $dlg.find('input[name="ls-print-dir"][value="' + (layout.direction || PrintDirection.Portrait) + '"]').prop("checked", true);
    $dlg.find("#luckysheet-print-grid").prop("checked", render.gridlines !== false);
}

function renderPreview() {
    const file = currentFile();
    if (!file) {
        return { pages: [] };
    }
    const pair = readDialogLayout(file);
    persist(file, pair.layout, pair.render);
    const pack = buildPages(file, pair.layout, pair.render);
    const $box = $("#luckysheet-print-box");
    $box.empty();
    pack.pages.forEach(function (page, i) {
        pack.pages.indexOf = pack.pages.indexOf.bind(pack.pages);
        const canvas = drawPageCanvas(page, file, pair.layout, pair.render, pack);
        canvas.setAttribute("data-print-page", String(i));
        $box.append(canvas);
    });
    return pack;
}

function ensurePrintStyle() {
    if (typeof document === "undefined") {
        return;
    }
    if (document.getElementById("luckysheet-print-inline-style")) {
        return;
    }
    const style = document.createElement("style");
    style.id = "luckysheet-print-inline-style";
    style.textContent = [
        ".luckysheet-print-box canvas{display:block;margin:8px auto;box-shadow:0 1px 4px rgba(0,0,0,.12);background:#fff;}",
        ".luckysheet-print-row{display:flex;align-items:center;margin:8px 0;gap:8px;}",
        ".luckysheet-print-row label{min-width:88px;}",
        ".luckysheet-print-preview{position:fixed;inset:0;background:#fff;z-index:100010;overflow:auto;}",
        "@media print{:not(html,head,body,.luckysheet-print-preview,.luckysheet-print-preview *){display:none!important;}.luckysheet-print-break{page-break-after:always;}#print-layout-options{display:none!important;}#" + DIALOG_ID + "{display:none!important;}}",
    ].join("");
    document.head.appendChild(style);
}

function createDialog() {
    ensurePrintStyle();
    $("#luckysheet-modal-dialog-mask").hide();
    $("#" + DIALOG_ID).remove();
    $("#" + PREVIEW_ID).remove();

    const _locale = locale();
    const p = _locale.print || {};
    const button = _locale.button || {};
    $("body").append(replaceHtml(modelHTML, {
        id: DIALOG_ID,
        addclass: "luckysheet-print-dialog",
        title: p.title || "打印设置",
        content: dialogHtml(),
        botton:
            '<button class="btn btn-default" id="luckysheet-print-preview-btn">' + (p.preview || "预览") + "</button>" +
            '<button class="btn btn-default" id="luckysheet-print-do-btn">' + (p.menuItemPrint || "打印") + "</button>" +
            '<button class="btn btn-default luckysheet-model-close-btn">' + (button.close || "关闭") + "</button>",
        style: "z-index:100004;min-width:640px;",
        close: button.close || "关闭",
    }));
    const $dlg = $("#" + DIALOG_ID);
    $dlg.attr("data-ls-instance", instanceAttr());
    const myw = $dlg.outerWidth();
    const myh = $dlg.outerHeight();
    $dlg.css({
        left: Math.max(20, ($(window).width() - myw) / 2),
        top: Math.max(20, ($(window).height() - myh) / 4),
    }).show();
    fillDialog(currentFile());
    renderPreview();
}

function closeDialog() {
    $("#" + DIALOG_ID).remove();
    $("#" + PREVIEW_ID).remove();
}

function init(license) {
    ensurePrintStyle();
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

function mountPreviewPages(pack, file, layout, render) {
    $("#" + PREVIEW_ID).remove();
    const $preview = $('<div class="luckysheet-print-preview" id="' + PREVIEW_ID + '" data-ls-instance="' + instanceAttr() + '"></div>');
    pack.pages.forEach(function (page, i) {
        const canvas = drawPageCanvas(page, file, layout, render, pack);
        const $page = $('<div class="luckysheet-print-break"></div>');
        $page.append(canvas);
        $preview.append($page);
        if (i === pack.pages.length - 1) {
            $page.removeClass("luckysheet-print-break");
        }
    });
    $("body").append($preview);
    return $preview;
}

function print(options) {
    const file = currentFile();
    if (!file) {
        return { pageCount: 0 };
    }
    ensurePrintStyle();
    const pair = options && (options.layout || options.render)
        ? {
            layout: layoutOf(file, options.layout),
            render: renderOf(file, options.render),
        }
        : readDialogLayout(file);
    persist(file, pair.layout, pair.render);
    return waitPrintResources(pair.render.draft).then(function () {
        const pack = buildPages(file, pair.layout, pair.render);
        mountPreviewPages(pack, file, pair.layout, pair.render);
        if (typeof window !== "undefined" && typeof window.print === "function") {
            window.print();
        }
        return { pageCount: pack.pages.length, range: pack.range };
    });
}

function updatePrintConfig(config) {
    const file = currentFile();
    if (!file) {
        return null;
    }
    const layout = layoutOf(file, config || {});
    if (config) {
        if (config.area) {
            layout.area = config.area;
        }
        if (config.printArea || config.PrintArea) {
            layout.rangeText = config.printArea || config.PrintArea;
        }
        if (config.paperSize) {
            layout.paperSize = config.paperSize;
        }
        if (config.direction) {
            layout.direction = config.direction;
        }
        if (config.scale) {
            layout.scale = config.scale;
        }
        if (config.customScale != null) {
            layout.customScale = clampScale(config.customScale);
        }
        if (config.margin) {
            layout.margin = config.margin;
        }
        if (config.area === "selection" || config.area === PrintArea.CurrentSelection) {
            const sel = selectionRange();
            layout.rangeText = chatatABC(sel.column[0]) + (sel.row[0] + 1) + ":" + chatatABC(sel.column[1]) + (sel.row[1] + 1);
            layout.area = PrintArea.CurrentSelection;
        }
    }
    persist(file, layout, renderOf(file));
    return layout;
}

function updatePrintRenderConfig(config) {
    const file = currentFile();
    if (!file) {
        return null;
    }
    const render = renderOf(file, config || {});
    persist(file, layoutOf(file), render);
    return render;
}

function setPrintTitles(which) {
    const file = currentFile();
    if (!file) {
        return;
    }
    const config = ensureFileConfig(file);
    const po = config.printoptions;
    const sel = selectionRange();
    po.PrintTitles = po.PrintTitles || {};
    if (which === "rows" || (which && which.row)) {
        po.PrintTitles.row = (file.name || "Sheet1") + "!$" + (sel.row[0] + 1) + ":$" + (sel.row[1] + 1);
    }
    if (which === "columns" || (which && which.column)) {
        po.PrintTitles.column = (file.name || "Sheet1") + "!$" + chatatABC(sel.column[0]) + ":$" + chatatABC(sel.column[1]);
    }
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
};

export default luckysheetPrint;
