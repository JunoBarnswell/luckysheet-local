/**
 * LuckySheet print layout mapping.
 * Public Univer Print enums only — no @univerjs-pro source.
 */

export const PrintArea = {
    CurrentSheet: "CurrentSheet",
    Workbook: "Workbook",
    CurrentSelection: "CurrentSelection",
    AllSelection: "AllSelection",
};

export const PrintPaperSize = {
    Letter: "Letter",
    Tabloid: "Tabloid",
    Legal: "Legal",
    Statement: "Statement",
    Executive: "Executive",
    Folio: "Folio",
    A3: "A3",
    A4: "A4",
    A5: "A5",
    B4: "B4",
    B5: "B5",
};

export const PrintDirection = {
    Portrait: "Portrait",
    Landscape: "Landscape",
};

export const PrintScale = {
    Origin: "Origin",
    FitWidth: "FitWidth",
    FitHeight: "FitHeight",
    FitPage: "FitPage",
    Custom: "Custom",
};

export const PrintFreeze = {
    Row: "Row",
    Column: "Column",
};

export const PrintPaperMargin = {
    Normal: "Normal",
    Narrow: "Narrow",
    Wide: "Wide",
    None: "None",
    Custom: "Custom",
};

export const PrintAlign = {
    Start: "Start",
    Middle: "Middle",
    End: "End",
};

export const PrintHeaderFooter = {
    PageSize: "PageSize",
    WorkbookTitle: "WorkbookTitle",
    WorksheetTitle: "WorksheetTitle",
    Date: "Date",
    Time: "Time",
};

export const PrintHeaderFooterSymbol = {
    WorkbookTitle: "@WorkbookTitle",
    WorksheetTitle: "@WorksheetTitle",
    DateA: "@DateA",
    DateB: "@DateB",
    DateC: "@DateC",
    DateD: "@DateD",
    DateE: "@DateE",
    TimeA: "@TimeA",
    TimeB: "@TimeB",
    TimeC: "@TimeC",
    TimeD: "@TimeD",
    Page: "@Page",
    SheetPage: "@SheetPage",
    PageTotal: "@TotalPage",
    SheetPageTotal: "@TotalSheetPage",
};

export const EXCEL_PAPER = {
    1: { name: PrintPaperSize.Letter, wmm: 215.9, hmm: 279.4 },
    3: { name: PrintPaperSize.Tabloid, wmm: 279.4, hmm: 431.8 },
    5: { name: PrintPaperSize.Legal, wmm: 215.9, hmm: 355.6 },
    6: { name: PrintPaperSize.Statement, wmm: 139.7, hmm: 215.9 },
    7: { name: PrintPaperSize.Executive, wmm: 184.15, hmm: 266.7 },
    8: { name: PrintPaperSize.A3, wmm: 297, hmm: 420 },
    9: { name: PrintPaperSize.A4, wmm: 210, hmm: 297 },
    11: { name: PrintPaperSize.A5, wmm: 148, hmm: 210 },
    12: { name: PrintPaperSize.B4, wmm: 250, hmm: 353 },
    13: { name: PrintPaperSize.B5, wmm: 176, hmm: 250 },
    14: { name: PrintPaperSize.Folio, wmm: 215.9, hmm: 330.2 },
};

export const PAPER_TO_EXCEL = {
    Letter: 1,
    Tabloid: 3,
    Legal: 5,
    Statement: 6,
    Executive: 7,
    Folio: 14,
    A3: 8,
    A4: 9,
    A5: 11,
    B4: 12,
    B5: 13,
};

export const MARGIN_PRESET_IN = {
    Normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    Narrow: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
    Wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
    None: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
};

export const PX_PER_MM = 96 / 25.4;
export const PX_PER_IN = 96;

export function excelPaperToName(code) {
    const item = EXCEL_PAPER[code];
    return item ? item.name : PrintPaperSize.A4;
}

export function paperSizeToMm(paperSize, pageSizeCustom, excelCode) {
    if (pageSizeCustom && pageSizeCustom.w && pageSizeCustom.h) {
        return { wmm: pageSizeCustom.w, hmm: pageSizeCustom.h };
    }
    if (paperSize && EXCEL_PAPER[PAPER_TO_EXCEL[paperSize]]) {
        const item = EXCEL_PAPER[PAPER_TO_EXCEL[paperSize]];
        return { wmm: item.wmm, hmm: item.hmm };
    }
    const item = EXCEL_PAPER[excelCode] || EXCEL_PAPER[9];
    return { wmm: item.wmm, hmm: item.hmm };
}

export function mmToPx(mm) {
    return (Number(mm) || 0) * PX_PER_MM;
}

export function inchToPx(inch) {
    return (Number(inch) || 0) * PX_PER_IN;
}

export function getPrintOptions(file) {
    const config = (file && file.config) || {};
    return config.printoptions || {};
}

export function defaultLayout() {
    return {
        area: PrintArea.CurrentSheet,
        paperSize: PrintPaperSize.A4,
        direction: PrintDirection.Portrait,
        scale: PrintScale.Origin,
        customScale: 100,
        freeze: [],
        margin: PrintPaperMargin.Normal,
        marginCustom: null,
        pageSizeCustom: null,
        maxRowsEachPage: 0,
        maxColumnsEachPage: 0,
        pageOrder: 0,
        subUnitIds: [],
        sheetIndex: null,
        range: null,
        rangeText: null,
    };
}

export function defaultRender() {
    return {
        gridlines: true,
        headings: false,
        hAlign: PrintAlign.Start,
        vAlign: PrintAlign.Start,
        headerFooter: [],
        headerFooterSetting: {
            topLeft: "",
            topCenter: "",
            topRight: "",
            bottomLeft: "",
            bottomCenter: "",
            bottomRight: "",
        },
        isCustomHeaderFooter: false,
        draft: false,
    };
}

export function normalizeLayoutFromPrintoptions(printoptions, overrides) {
    const po = printoptions || {};
    const setup = po.pageSetup || {};
    const layout = defaultLayout();
    if (po.PrintArea) {
        layout.area = PrintArea.CurrentSheet;
        layout.rangeText = po.PrintArea;
    }
    layout.paperSize = excelPaperToName(setup.paperSize);
    if (setup.paperWidth && setup.paperHeight) {
        layout.pageSizeCustom = {
            w: parseLengthToMm(setup.paperWidth),
            h: parseLengthToMm(setup.paperHeight),
        };
    }
    if (setup.orientation === 1) {
        layout.direction = PrintDirection.Landscape;
    } else if (setup.orientation === 2) {
        layout.direction = PrintDirection.Portrait;
    }
    if (setup.fitToWidth) {
        layout.scale = PrintScale.FitWidth;
    } else if (setup.fitToHeight) {
        layout.scale = PrintScale.FitHeight;
    } else if (setup.scale && setup.scale !== 100) {
        layout.scale = PrintScale.Custom;
        layout.customScale = clampScale(setup.scale);
    }
    if (po.PrintTitles) {
        if (po.PrintTitles.row) {
            layout.freeze.push(PrintFreeze.Row);
        }
        if (po.PrintTitles.column) {
            layout.freeze.push(PrintFreeze.Column);
        }
    }
    layout.margin = inferMarginPreset(po.pageMargins);
    if (layout.margin === PrintPaperMargin.Custom || (overrides && overrides.marginCustom)) {
        layout.marginCustom = marginCustomFromExcel(po.pageMargins);
    }
    return Object.assign(layout, overrides || {});
}

export function normalizeRenderFromPrintoptions(printoptions, overrides) {
    const po = printoptions || {};
    const opts = po.printOptions || {};
    const setup = po.pageSetup || {};
    const render = defaultRender();
    render.gridlines = opts.gridLines == null ? true : !!opts.gridLines;
    render.headings = !!opts.headings;
    render.hAlign = opts.horizontalCentered ? PrintAlign.Middle : PrintAlign.Start;
    render.vAlign = opts.verticalCentered ? PrintAlign.Middle : PrintAlign.Start;
    render.draft = !!setup.draft;
    if (po.headerFooter) {
        render.headerFooterSetting = headerFooterFromExcel(po.headerFooter);
    }
    return Object.assign(render, overrides || {});
}

export function writePrintoptions(printoptions, layout, render) {
    const next = Object.assign({}, printoptions || {});
    const setup = Object.assign({}, next.pageSetup || {});
    const opts = Object.assign({}, next.printOptions || {});
    const margins = resolveMarginInches(layout, next.pageMargins);

    if (layout.rangeText) {
        next.PrintArea = layout.rangeText;
    } else if (layout.rangeText === null) {
        delete next.PrintArea;
    }
    if (layout.paperSize && PAPER_TO_EXCEL[layout.paperSize] != null) {
        setup.paperSize = PAPER_TO_EXCEL[layout.paperSize];
    }
    if (layout.pageSizeCustom && layout.pageSizeCustom.w && layout.pageSizeCustom.h) {
        setup.paperWidth = layout.pageSizeCustom.w + "mm";
        setup.paperHeight = layout.pageSizeCustom.h + "mm";
    } else {
        delete setup.paperWidth;
        delete setup.paperHeight;
    }
    if (layout.direction === PrintDirection.Landscape) {
        setup.orientation = 1;
    } else if (layout.direction === PrintDirection.Portrait) {
        setup.orientation = 2;
    }
    setup.scale = clampScale(layout.customScale || setup.scale || 100);
    setup.fitToWidth = layout.scale === PrintScale.FitWidth || layout.scale === PrintScale.FitPage ? 1 : 0;
    setup.fitToHeight = layout.scale === PrintScale.FitHeight || layout.scale === PrintScale.FitPage ? 1 : 0;
    setup.draft = render && render.draft ? 1 : setup.draft || 0;

    if (render) {
        opts.gridLines = render.gridlines ? 1 : 0;
        opts.headings = render.headings ? 1 : 0;
        opts.horizontalCentered = render.hAlign === PrintAlign.Middle || render.hAlign === PrintAlign.End ? 1 : 0;
        opts.verticalCentered = render.vAlign === PrintAlign.Middle || render.vAlign === PrintAlign.End ? 1 : 0;
    }

    next.pageSetup = setup;
    next.printOptions = opts;
    next.pageMargins = margins;
    return next;
}

export function inferMarginPreset(pageMargins) {
    if (!pageMargins) {
        return PrintPaperMargin.Normal;
    }
    const left = Number(pageMargins.left);
    if (left === 0 && Number(pageMargins.right) === 0) {
        return PrintPaperMargin.None;
    }
    if (left <= 0.3) {
        return PrintPaperMargin.Narrow;
    }
    if (left >= 0.95) {
        return PrintPaperMargin.Wide;
    }
    return PrintPaperMargin.Normal;
}

export function clampScale(value) {
    const n = Number(value);
    if (!isFinite(n)) {
        return 100;
    }
    return Math.min(400, Math.max(10, n));
}

export function parseLengthToMm(value) {
    if (value == null) {
        return null;
    }
    if (typeof value === "number") {
        return value;
    }
    const text = String(value).trim();
    const n = parseFloat(text);
    if (!isFinite(n)) {
        return null;
    }
    if (/in$/i.test(text)) {
        return n * 25.4;
    }
    if (/cm$/i.test(text)) {
        return n * 10;
    }
    if (/pt$/i.test(text)) {
        return n * 25.4 / 72;
    }
    return n;
}

export function headerFooterFromExcel(headerFooter) {
    const setting = defaultRender().headerFooterSetting;
    const first = headerFooter.oddHeader || headerFooter.firstHeader || headerFooter;
    if (first && first.left && first.left[0] && first.left[0].v) {
        setting.topLeft = String(first.left[0].v);
    }
    if (first && first.center && first.center[0] && first.center[0].v) {
        setting.topCenter = String(first.center[0].v);
    }
    if (first && first.right && first.right[0] && first.right[0].v) {
        setting.topRight = String(first.right[0].v);
    }
    return setting;
}

export function cellDisplayValue(cell) {
    if (cell == null) {
        return "";
    }
    if (typeof cell !== "object") {
        return String(cell);
    }
    if (cell.m != null && cell.m !== "") {
        return String(cell.m);
    }
    if (cell.v != null && cell.v !== "") {
        return String(cell.v);
    }
    return "";
}

export function usedRange(data) {
    if (!data || !data.length) {
        return { row: [0, 0], column: [0, 0] };
    }
    let maxR = 0;
    let maxC = 0;
    let found = false;
    const rows = data.length;
    for (let r = 0; r < rows; r++) {
        const row = data[r];
        if (!row) {
            continue;
        }
        const cols = row.length || 0;
        for (let c = 0; c < cols; c++) {
            const cell = row[c];
            if (cell == null || cell === "") {
                continue;
            }
            if (typeof cell === "object") {
                if (cell.v != null || cell.m != null || cell.f) {
                    found = true;
                    if (r > maxR) {
                        maxR = r;
                    }
                    if (c > maxC) {
                        maxC = c;
                    }
                }
                continue;
            }
        }
    }
    if (!found) {
        return { row: [0, 0], column: [0, 0] };
    }
    return { row: [0, maxR], column: [0, maxC] };
}

export function rowStartPx(visibledatarow, r) {
    if (r <= 0) {
        return 0;
    }
    return visibledatarow[r - 1] || 0;
}

export function colStartPx(visibledatacolumn, c) {
    if (c <= 0) {
        return 0;
    }
    return visibledatacolumn[c - 1] || 0;
}

export function rowHeight(visibledatarow, r) {
    const end = visibledatarow[r] || 0;
    const start = rowStartPx(visibledatarow, r);
    return Math.max(end - start, 0);
}

export function colWidth(visibledatacolumn, c) {
    const end = visibledatacolumn[c] || 0;
    const start = colStartPx(visibledatacolumn, c);
    return Math.max(end - start, 0);
}

export function resolvePaperPx(layout, printoptions) {
    const setup = (printoptions && printoptions.pageSetup) || {};
    const size = paperSizeToMm(layout.paperSize, layout.pageSizeCustom, setup.paperSize);
    let w = mmToPx(size.wmm);
    let h = mmToPx(size.hmm);
    if (layout.direction === PrintDirection.Landscape) {
        const t = w;
        w = h;
        h = t;
    }
    const margins = resolveMarginInches(layout, printoptions && printoptions.pageMargins);
    const pad = {
        left: inchToPx(margins.left),
        right: inchToPx(margins.right),
        top: inchToPx(margins.top),
        bottom: inchToPx(margins.bottom),
    };
    return {
        pageW: w,
        pageH: h,
        innerW: Math.max(40, w - pad.left - pad.right),
        innerH: Math.max(40, h - pad.top - pad.bottom),
        pad: pad,
    };
}

export function resolveMarginInches(layout, pageMargins) {
    if (layout.margin === PrintPaperMargin.Custom && layout.marginCustom) {
        return Object.assign({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 }, layout.marginCustom);
    }
    if (layout.margin && MARGIN_PRESET_IN[layout.margin]) {
        return Object.assign({}, MARGIN_PRESET_IN[layout.margin]);
    }
    return Object.assign({}, MARGIN_PRESET_IN.Normal, pageMargins || {});
}

export function marginCustomFromExcel(pageMargins) {
    if (!pageMargins) {
        return null;
    }
    return {
        left: Number(pageMargins.left) || 0,
        right: Number(pageMargins.right) || 0,
        top: Number(pageMargins.top) || 0,
        bottom: Number(pageMargins.bottom) || 0,
    };
}

export function parsePrintTitles(file, layout) {
    const po = getPrintOptions(file);
    const titles = po.PrintTitles || {};
    const result = { row: null, column: null };
    if (titles.row) {
        result.row = parseTitleRange(titles.row);
    }
    if (titles.column) {
        result.column = parseTitleRange(titles.column);
    }
    if (layout && layout.freeze) {
        if (layout.freeze.indexOf(PrintFreeze.Row) > -1 && !result.row) {
            result.row = [0, 0];
        }
        if (layout.freeze.indexOf(PrintFreeze.Column) > -1 && !result.column) {
            result.column = [0, 0];
        }
    }
    return result;
}

function parseTitleRange(text) {
    if (!text) {
        return null;
    }
    let raw = String(text).replace(/\$/g, "");
    if (raw.indexOf("!") > -1) {
        raw = raw.split("!").pop();
    }
    const parts = raw.split(":");
    if (parts.length !== 2) {
        return null;
    }
    const a = parts[0];
    const b = parts[1];
    const aNum = /^\d+$/.test(a);
    const bNum = /^\d+$/.test(b);
    if (aNum && bNum) {
        const r0 = Number(a) - 1;
        const r1 = Number(b) - 1;
        return [Math.min(r0, r1), Math.max(r0, r1)];
    }
    if (!aNum && !bNum) {
        const c0 = colLabelToIndex(a);
        const c1 = colLabelToIndex(b);
        return [Math.min(c0, c1), Math.max(c0, c1)];
    }
    return null;
}

function colLabelToIndex(label) {
    let n = 0;
    const s = String(label || "A").toUpperCase();
    for (let i = 0; i < s.length; i++) {
        n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return Math.max(0, n - 1);
}

export function resolveHeaderFooterText(text, ctx) {
    if (!text) {
        return "";
    }
    const now = new Date();
    const map = {
        "@WorkbookTitle": ctx.workbookTitle || "",
        "@WorksheetTitle": ctx.worksheetTitle || "",
        "@DateA": now.toLocaleDateString(),
        "@DateB": now.toISOString().slice(0, 10),
        "@DateC": now.getFullYear() + "/" + (now.getMonth() + 1) + "/" + now.getDate(),
        "@DateD": now.toLocaleDateString("en-US"),
        "@DateE": now.toLocaleDateString("zh-CN"),
        "@TimeA": now.toLocaleTimeString(),
        "@TimeB": now.toTimeString().slice(0, 8),
        "@TimeC": now.toLocaleTimeString("en-US", { hour12: false }),
        "@TimeD": now.getHours() + ":" + String(now.getMinutes()).padStart(2, "0"),
        "@Page": String(ctx.page || 1),
        "@SheetPage": String(ctx.sheetPage || ctx.page || 1),
        "@TotalPage": String(ctx.pageTotal || 1),
        "@TotalSheetPage": String(ctx.sheetPageTotal || ctx.pageTotal || 1),
        PageSize: String(ctx.page || 1),
        PageNumber: String(ctx.page || 1),
        WorkbookTitle: ctx.workbookTitle || "",
        WorksheetTitle: ctx.worksheetTitle || "",
        Date: now.toLocaleDateString(),
        Time: now.toLocaleTimeString(),
    };
    return String(text).replace(/@[A-Za-z]+|[A-Za-z]+/g, function (token) {
        return map[token] != null ? map[token] : token;
    });
}

export function adjustPaginationForMerge(pages, merge, visibledatarow, visibledatacolumn) {
    if (!merge || !pages || !pages.length) {
        return pages;
    }
    const keys = Object.keys(merge);
    if (!keys.length) {
        return pages;
    }
    return pages.map(function (page) {
        let r0 = page.row[0];
        let r1 = page.row[1];
        let c0 = page.column[0];
        let c1 = page.column[1];
        keys.forEach(function (key) {
            const parts = key.split("_");
            const mr = Number(parts[0]);
            const mc = Number(parts[1]);
            const span = merge[key];
            const mre = mr + span.rs - 1;
            const mce = mc + span.cs - 1;
            const overlapRow = mr <= r1 && mre >= r0;
            const overlapCol = mc <= c1 && mce >= c0;
            const splitRow = mr < r0 && mre >= r0 && mre <= r1;
            const splitCol = mc < c0 && mce >= c0 && mce <= c1;
            if (overlapRow && splitRow) {
                r0 = mre + 1;
            }
            if (overlapCol && splitCol) {
                c0 = mce + 1;
            }
            if (overlapRow && mr > r1) {
                r1 = Math.max(r0, mr - 1);
            }
            if (overlapCol && mc > c1) {
                c1 = Math.max(c0, mc - 1);
            }
        });
        if (r0 > r1) {
            r0 = page.row[0];
            r1 = page.row[0];
        }
        if (c0 > c1) {
            c0 = page.column[0];
            c1 = page.column[0];
        }
        return { row: [r0, r1], column: [c0, c1] };
    }).filter(function (page) {
        return page.row[0] <= page.row[1] && page.column[0] <= page.column[1];
    });
}

function advanceAxis(start, end, sizes, startPxFn, limitPx, maxCount) {
    let cursor = start;
    let used = 0;
    let count = 0;
    let lastVisible = null;
    while (cursor <= end) {
        const size = Math.max(0, (sizes[cursor] || 0) - startPxFn(sizes, cursor));
        if (size === 0) {
            cursor += 1;
            continue;
        }
        if (count > 0 && used + size > limitPx) {
            break;
        }
        used += size;
        count += 1;
        lastVisible = cursor;
        cursor += 1;
        if (maxCount && count >= maxCount) {
            break;
        }
    }
    if (count === 0) {
        while (cursor <= end && ((sizes[cursor] || 0) - startPxFn(sizes, cursor)) <= 0) {
            cursor += 1;
        }
        return cursor > end ? end : cursor;
    }
    return lastVisible;
}

export function paginateByPaper(range, visibledatarow, visibledatacolumn, paper, layout) {
    const pages = [];
    if (!range || !visibledatarow || !visibledatacolumn) {
        return pages;
    }
    const maxR = layout && layout.maxRowsEachPage;
    const maxC = layout && layout.maxColumnsEachPage;
    const pageOrder = layout && layout.pageOrder;
    const innerW = paper.innerW;
    const innerH = paper.innerH;

    const rowBands = [];
    let r = range.row[0];
    while (r <= range.row[1]) {
        const r2 = advanceAxis(r, range.row[1], visibledatarow, rowStartPx, innerH, maxR);
        rowBands.push([r, r2]);
        r = Math.max(r2 + 1, r + 1);
    }
    const colBands = [];
    let c = range.column[0];
    while (c <= range.column[1]) {
        const c2 = advanceAxis(c, range.column[1], visibledatacolumn, colStartPx, innerW, maxC);
        colBands.push([c, c2]);
        c = Math.max(c2 + 1, c + 1);
    }

    if (pageOrder === 1) {
        for (let i = 0; i < rowBands.length; i++) {
            for (let j = 0; j < colBands.length; j++) {
                pages.push({
                    row: rowBands[i],
                    column: colBands[j],
                });
            }
        }
    } else {
        for (let j = 0; j < colBands.length; j++) {
            for (let i = 0; i < rowBands.length; i++) {
                pages.push({
                    row: rowBands[i],
                    column: colBands[j],
                });
            }
        }
    }
    if (layout && layout.merge) {
        return adjustPaginationForMerge(pages, layout.merge, visibledatarow, visibledatacolumn);
    }
    return pages;
}

export function computeFitScale(range, visibledatarow, visibledatacolumn, paper, scaleType) {
    const contentW = colStartPx(visibledatacolumn, range.column[1] + 1) - colStartPx(visibledatacolumn, range.column[0]);
    const contentH = rowStartPx(visibledatarow, range.row[1] + 1) - rowStartPx(visibledatarow, range.row[0]);
    if (contentW <= 0 || contentH <= 0) {
        return 1;
    }
    const sx = paper.innerW / contentW;
    const sy = paper.innerH / contentH;
    if (scaleType === PrintScale.FitWidth) {
        return Math.min(1, sx);
    }
    if (scaleType === PrintScale.FitHeight) {
        return Math.min(1, sy);
    }
    if (scaleType === PrintScale.FitPage) {
        return Math.min(1, sx, sy);
    }
    return 1;
}
