import locale from "../../locale/locale";
import Store from "../../store";
import { modelHTML } from "../../controllers/constant";
import { replaceHtml } from "../../utils/util";
import {
    PrintArea,
    PrintPaperSize,
    PrintDirection,
    PrintScale,
    PrintPaperMargin,
    PrintAlign,
} from "./printLayout";
import { ensurePrintStyleTag, PRINT_CONTAINER_CLASS, PRINT_CANVAS_CLASS } from "./printBrowser";
import { runInPrintSession } from "./printSession";
import { drawPageCanvas } from "./printRenderer";
import {
    currentFile,
    layoutOf,
    renderOf,
    persist,
    buildPages,
    preparePrint,
    createPreparedPrintSession,
} from "./printManager";

export const DIALOG_ID = "luckysheet-print-dialog";
export const PREVIEW_ID = "luckysheet-print-preview";

function printLocale() {
    return (locale().print || {});
}

function dialogName(instanceId) {
    return "ls-print-dir-" + instanceId;
}

export function dialogHtml(instanceId) {
    const p = printLocale();
    const papers = ["Letter", "Tabloid", "Legal", "Statement", "Executive", "Folio", "A3", "A4", "A5", "B4", "B5"];
    const paperOpts = papers.map(function (name) {
        return '<option value="' + name + '">' + name + "</option>";
    }).join("");
    const margins = [
        PrintPaperMargin.Normal,
        PrintPaperMargin.Narrow,
        PrintPaperMargin.Wide,
        PrintPaperMargin.None,
        PrintPaperMargin.Custom,
    ];
    const marginOpts = margins
        .map(function (name) {
            return '<option value="' + name + '">' + (p["margin" + name] || name) + "</option>";
        })
        .join("");
    const scales = [
        PrintScale.Origin,
        PrintScale.FitWidth,
        PrintScale.FitHeight,
        PrintScale.FitPage,
        PrintScale.Custom,
    ];
    const scaleOpts = scales
        .map(function (name) {
            return '<option value="' + name + '">' + (p["scale" + name] || name) + "</option>";
        })
        .join("");

    return (
        '<div class="luckysheet-print" id="print-layout-options">' +
            '<div class="luckysheet-print-panel">' +
                '<div class="luckysheet-print-settings">' +
                    '<div class="luckysheet-print-title">' + (p.title || "打印设置") + "</div>" +
                    '<div class="luckysheet-print-section">' +
                        '<div class="luckysheet-print-section-title">' + (p.sectionPage || "页面") + "</div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.range || "打印范围") + "</label>" +
                            '<select id="luckysheet-print-area">' +
                                '<option value="' + PrintArea.CurrentSheet + '">' + (p.current || "当前工作表") + "</option>" +
                                '<option value="' + PrintArea.CurrentSelection + '">' + (p.area || "选中区域") + "</option>" +
                                '<option value="' + PrintArea.AllSelection + '">' + (p.allSelection || "全部选区") + "</option>" +
                                '<option value="' + PrintArea.Workbook + '">' + (p.workbook || "整个工作簿") + "</option>" +
                            "</select></div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.size || "纸张大小") + "</label>" +
                            '<select id="luckysheet-print-paper">' + paperOpts + "</select></div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.direction || "打印方向") + "</label>" +
                            '<div class="luckysheet-print-radio">' +
                                '<label><input type="radio" name="' + dialogName(instanceId) + '" value="' + PrintDirection.Portrait + '"/> ' + (p.vertical || "纵向") + "</label>" +
                                '<label><input type="radio" name="' + dialogName(instanceId) + '" value="' + PrintDirection.Landscape + '"/> ' + (p.horizontal || "横向") + "</label>" +
                            "</div></div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.margin || "页边距") + "</label>" +
                            '<select id="luckysheet-print-margin">' + marginOpts + "</select></div>" +
                    "</div>" +
                    '<div class="luckysheet-print-section">' +
                        '<div class="luckysheet-print-section-title">' + (p.sectionScale || "缩放") + "</div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.scale || "缩放") + "</label>" +
                            '<select id="luckysheet-print-scale">' + scaleOpts + "</select></div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.customScale || "自定义%") + "</label>" +
                            '<input type="number" id="luckysheet-print-custom-scale" min="10" max="400" value="100"/></div>' +
                    "</div>" +
                    '<div class="luckysheet-print-section">' +
                        '<div class="luckysheet-print-section-title">' + (p.sectionRender || "显示") + "</div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.showLine || "显示网格线") + "</label>" +
                            '<input type="checkbox" id="luckysheet-print-grid" checked/></div>' +
                        '<div class="luckysheet-print-row"><label>' + (p.headings || "行列标题") + "</label>" +
                            '<input type="checkbox" id="luckysheet-print-headings"/></div>' +
                        '<div class="luckysheet-print-row"><label>' + (p.draft || "草稿模式") + "</label>" +
                            '<input type="checkbox" id="luckysheet-print-draft"/></div>' +
                    "</div>" +
                    '<div class="luckysheet-print-section">' +
                        '<div class="luckysheet-print-section-title">' + (p.sectionHeader || "页眉页脚") + "</div>" +
                        '<div class="luckysheet-print-row"><label>' + (p.headerCenter || "页眉") + "</label>" +
                            '<input type="text" id="luckysheet-print-header-center" placeholder="@WorksheetTitle @Page/@TotalPage"/></div>' +
                        '<div class="luckysheet-print-row"><label>' + (p.footerCenter || "页脚") + "</label>" +
                            '<input type="text" id="luckysheet-print-footer-center" placeholder="@DateA @TimeA"/></div>' +
                    "</div>" +
                    '<div class="luckysheet-print-suggest">' + (p.suggest || "") + "</div>" +
                "</div>" +
                '<div class="luckysheet-print-preview-pane">' +
                    '<div class="luckysheet-print-box" id="luckysheet-print-box"></div>' +
                "</div>" +
            "</div>" +
        "</div>"
    );
}

export function readDialogLayout(file) {
    const layout = layoutOf(file);
    const render = renderOf(file);
    const $dlg = $("#" + DIALOG_ID);
    if (!$dlg.length) {
        return { layout: layout, render: render };
    }
    layout.area = $dlg.find("#luckysheet-print-area").val() || layout.area;
    layout.paperSize = $dlg.find("#luckysheet-print-paper").val() || layout.paperSize;
    layout.margin = $dlg.find("#luckysheet-print-margin").val() || layout.margin;
    layout.scale = $dlg.find("#luckysheet-print-scale").val() || layout.scale;
    layout.customScale = Number($dlg.find("#luckysheet-print-custom-scale").val()) || layout.customScale;
    const dir = $dlg.find('input[name="' + dialogName(Store.instanceId) + '"]:checked').val();
    if (dir) {
        layout.direction = dir;
    }
    render.gridlines = $dlg.find("#luckysheet-print-grid").prop("checked");
    render.headings = $dlg.find("#luckysheet-print-headings").prop("checked");
    render.draft = $dlg.find("#luckysheet-print-draft").prop("checked");
    render.headerFooterSetting = render.headerFooterSetting || {};
    render.headerFooterSetting.topCenter = $dlg.find("#luckysheet-print-header-center").val() || "";
    render.headerFooterSetting.bottomCenter = $dlg.find("#luckysheet-print-footer-center").val() || "";
    render.isCustomHeaderFooter = true;
    return { layout: layout, render: render };
}

export function fillDialog(file) {
    const layout = layoutOf(file);
    const render = renderOf(file);
    const $dlg = $("#" + DIALOG_ID);
    $dlg.find("#luckysheet-print-area").val(layout.area || PrintArea.CurrentSheet);
    $dlg.find("#luckysheet-print-paper").val(layout.paperSize || PrintPaperSize.A4);
    $dlg.find("#luckysheet-print-margin").val(layout.margin || PrintPaperMargin.Normal);
    $dlg.find("#luckysheet-print-scale").val(layout.scale || PrintScale.Origin);
    $dlg.find("#luckysheet-print-custom-scale").val(layout.customScale || 100);
    $dlg.find('input[name="' + dialogName(Store.instanceId) + '"][value="' + (layout.direction || PrintDirection.Portrait) + '"]').prop("checked", true);
    $dlg.find("#luckysheet-print-grid").prop("checked", render.gridlines !== false);
    $dlg.find("#luckysheet-print-headings").prop("checked", !!render.headings);
    $dlg.find("#luckysheet-print-draft").prop("checked", !!render.draft);
    const hf = render.headerFooterSetting || {};
    $dlg.find("#luckysheet-print-header-center").val(hf.topCenter || "");
    $dlg.find("#luckysheet-print-footer-center").val(hf.bottomCenter || "");
}

export function renderPreview() {
    const file = currentFile();
    if (!file) {
        return { pages: [] };
    }
    const pair = readDialogLayout(file);
    persist(file, pair.layout, pair.render);
    return createPreparedPrintSession(pair.layout, pair.render).then(function (prepared) {
        return runInPrintSession(prepared.session, function () {
            const plan = prepared.plan;
            const $box = $("#luckysheet-print-box");
            $box.empty();
            plan.entries.forEach(function (entry, i) {
                const meta = {
                    pageIndex: i,
                    pageTotal: plan.entries.length,
                    sheetPage: entry.sheetPage,
                    sheetPageTotal: entry.pack.pages.length,
                };
                const canvas = drawPageCanvas(entry.page, entry.file, prepared.session.layout, prepared.session.render, entry.pack, meta, prepared.session);
                canvas.setAttribute("data-print-page", String(i));
                $box.append(canvas);
            });
            return plan;
        });
    }).catch(function (error) {
        if (error && /Print instance was destroyed/.test(error.message || "")) {
            return { canceled: true, pages: [] };
        }
        throw error;
    });
}

export function createDialogMarkup(instanceAttr) {
    const p = printLocale();
    const button = locale().button || {};
    return {
        id: DIALOG_ID,
        addclass: "luckysheet-print-dialog",
        title: p.title || "打印设置",
        content: dialogHtml(instanceAttr),
        botton:
            '<button class="btn btn-default" id="luckysheet-print-preview-btn">' + (p.preview || "预览") + "</button>" +
            '<button class="btn btn-default" id="luckysheet-print-do-btn">' + (p.menuItemPrint || "打印") + "</button>" +
            '<button class="btn btn-default luckysheet-model-close-btn">' + (button.close || "关闭") + "</button>",
        style: "z-index:100004;min-width:860px;",
        close: button.close || "关闭",
        instanceAttr: instanceAttr,
    };
}

export function mountDialog(instanceAttr) {
    ensurePrintStyleTag();
    $("#luckysheet-modal-dialog-mask").hide();
    $("#" + DIALOG_ID).remove();
    $("#" + PREVIEW_ID).remove();
    const markup = createDialogMarkup(instanceAttr);
    $("body").append(
        replaceHtml(modelHTML, {
            id: markup.id,
            addclass: markup.addclass,
            title: markup.title,
            content: markup.content,
            botton: markup.botton,
            style: markup.style,
            close: markup.close,
        })
    );
    const $dlg = $("#" + DIALOG_ID);
    $dlg.attr("data-ls-instance", instanceAttr);
    const myw = $dlg.outerWidth();
    const myh = $dlg.outerHeight();
    $dlg.css({
        left: Math.max(20, ($(window).width() - myw) / 2),
        top: Math.max(20, ($(window).height() - myh) / 6),
    }).show();
    fillDialog(currentFile());
    return $dlg;
}

export function closeDialogDom() {
    $("#" + DIALOG_ID).remove();
    $("#" + PREVIEW_ID).remove();
}

export function mountPreviewPages(plan, instanceAttr) {
    $("#" + PREVIEW_ID).remove();
    const $preview = $("<div></div>")
        .addClass("luckysheet-print-preview " + PRINT_CONTAINER_CLASS)
        .attr("id", PREVIEW_ID)
        .attr("data-ls-instance", instanceAttr)
        .attr("data-print-session", plan.session.id);
    plan.entries.forEach(function (entry, i) {
        const meta = {
            pageIndex: i,
            pageTotal: plan.entries.length,
            sheetPage: entry.sheetPage,
            sheetPageTotal: entry.pack.pages.length,
        };
        const canvas = drawPageCanvas(entry.page, entry.file, plan.session.layout, plan.session.render, entry.pack, meta, plan.session);
        const $page = $('<div class="luckysheet-print-break ' + PRINT_CANVAS_CLASS + '"></div>');
        $page.append(canvas);
        $preview.append($page);
        if (i === plan.entries.length - 1) {
            $page.removeClass("luckysheet-print-break");
        }
    });
    $("body").append($preview);
    return $preview;
}
