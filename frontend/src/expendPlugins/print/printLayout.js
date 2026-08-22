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
};

export const PAPER_TO_EXCEL = {
    Letter: 1,
    Tabloid: 3,
    Legal: 5,
    Statement: 6,
    Executive: 7,
    Folio: 5,
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
        pageSizeCustom: null,
        maxRowsEachPage: 0,
        maxColumnsEachPage: 0,
        sheetIndex: null,
        range: null,
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
    const margins = Object.assign({}, next.pageMargins || {}, MARGIN_PRESET_IN[layout.margin] || {});

    if (layout.rangeText) {
        next.PrintArea = layout.rangeText;
    }
    if (layout.paperSize && PAPER_TO_EXCEL[layout.paperSize] != null) {
        setup.paperSize = PAPER_TO_EXCEL[layout.paperSize];
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
            if (typeof cell === "object" && cell.v == null && cell.m == null && !cell.f) {
                continue;
            }
            found = true;
            if (r > maxR) {
                maxR = r;
            }
            if (c > maxC) {
                maxC = c;
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
    const margins = (printoptions && printoptions.pageMargins) || MARGIN_PRESET_IN[layout.margin] || MARGIN_PRESET_IN.Normal;
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

function advanceAxis(start, end, sizes, startPxFn, limitPx, maxCount) {
    let cursor = start;
    let used = 0;
    let count = 0;
    while (cursor <= end) {
        const size = Math.max(1, (sizes[cursor] || 0) - startPxFn(sizes, cursor));
        if (count > 0 && used + size > limitPx) {
            break;
        }
        used += size;
        count += 1;
        cursor += 1;
        if (maxCount && count >= maxCount) {
            break;
        }
    }
    if (count === 0) {
        return start;
    }
    return cursor - 1;
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
        r = r2 + 1;
    }
    const colBands = [];
    let c = range.column[0];
    while (c <= range.column[1]) {
        const c2 = advanceAxis(c, range.column[1], visibledatacolumn, colStartPx, innerW, maxC);
        colBands.push([c, c2]);
        c = c2 + 1;
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
