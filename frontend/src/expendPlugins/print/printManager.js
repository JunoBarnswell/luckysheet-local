import Store from "../../store";
import formula from "../../global/formula";
import { getSheetIndex } from "../../methods/get";
import { chatatABC } from "../../utils/util";
import {
    PrintArea,
    PrintScale,
    getPrintOptions,
    normalizeLayoutFromPrintoptions,
    normalizeRenderFromPrintoptions,
    writePrintoptions,
    resolvePaperPx,
    paginateByPaper,
    computeFitScale,
    clampScale,
    usedRange,
    parsePrintTitles,
    rowStartPx,
    colStartPx,
} from "./printLayout";
import { drawPageCanvas, preloadPrintImages } from "./printRenderer";
import { waitPrintResources } from "./printResources";
import { createPrintSession, preparePrintSession, runInPrintSession } from "./printSession";

export function currentFile() {
    const index = getSheetIndex(Store.currentSheetIndex);
    return Store.luckysheetfile && Store.luckysheetfile[index];
}

export function fileByIndex(sheetIndex) {
    const index = getSheetIndex(sheetIndex);
    return Store.luckysheetfile && Store.luckysheetfile[index];
}

export function ensureFileConfig(file) {
    if (!file.config) {
        file.config = {};
    }
    if (!file.config.printoptions) {
        file.config.printoptions = {};
    }
    return file.config;
}

export function parseRangeText(text) {
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

export function selectionRange() {
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
    const mode = area || PrintArea.CurrentSheet;

    if (mode === PrintArea.CurrentSelection || mode === "selection") {
        return selectionRange();
    }
    if (mode === PrintArea.AllSelection) {
        const ranges = Store.luckysheet_select_save || [];
        if (!ranges.length) {
            return selectionRange();
        }
        let r0 = ranges[0].row[0],
            r1 = ranges[0].row[1],
            c0 = ranges[0].column[0],
            c1 = ranges[0].column[1];
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

export function layoutOf(file, overrides) {
    const po = getPrintOptions(file);
    const saved = (file && file.config && file.config.printLayout) || {};
    return normalizeLayoutFromPrintoptions(po, Object.assign({}, saved, overrides || {}));
}

export function renderOf(file, overrides) {
    const po = getPrintOptions(file);
    const saved = (file && file.config && file.config.printRender) || {};
    return normalizeRenderFromPrintoptions(po, Object.assign({}, saved, overrides || {}));
}

export function persist(file, layout, render) {
    const config = ensureFileConfig(file);
    config.printoptions = writePrintoptions(config.printoptions, layout, render);
    config.printLayout = layout;
    config.printRender = render;
}

export function resolveTargetFiles(layout) {
    if (layout.area === PrintArea.Workbook || layout.area === "Workbook") {
        return (Store.luckysheetfile || []).slice();
    }
    if (layout.subUnitIds && layout.subUnitIds.length) {
        const files = [];
        layout.subUnitIds.forEach(function (item) {
            if (typeof item === "string") {
                const f = fileByIndex(item);
                if (f) {
                    files.push(f);
                }
            } else if (item && item.id) {
                const f = fileByIndex(item.id);
                if (f) {
                    files.push(f);
                }
            }
        });
        if (files.length) {
            return files;
        }
    }
    const file = currentFile();
    return file ? [file] : [];
}

export function buildPagesForFile(file, layout, render) {
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
    const layoutWithMerge = Object.assign({}, layout, {
        merge: ((file && file.config) || {}).merge || {},
    });
    const pages = paginateByPaper(range, visibledatarow, visibledatacolumn, scaledPaper, layoutWithMerge);
    return {
        file: file,
        range: range,
        paper: paper,
        pages: pages,
        scale: scale,
        visibledatarow: visibledatarow,
        visibledatacolumn: visibledatacolumn,
    };
}

function buildPagesForTarget(target, layout, render) {
    const file = target.file;
    const range = target.range;
    const visibledatarow = target.visibledatarow;
    const visibledatacolumn = target.visibledatacolumn;
    const po = getPrintOptions(file);
    const paper = resolvePaperPx(layout, po);
    let scale = 1;
    if (layout.scale === PrintScale.Custom) {
        scale = clampScale(layout.customScale) / 100;
    } else if (layout.scale !== PrintScale.Origin) {
        scale = computeFitScale(range, visibledatarow, visibledatacolumn, paper, layout.scale);
    }
    const titles = parsePrintTitles(file, layout);
    const titleH = titles.row
        ? rowStartPx(visibledatarow, titles.row[1] + 1) - rowStartPx(visibledatarow, titles.row[0])
        : 0;
    const titleW = titles.column
        ? colStartPx(visibledatacolumn, titles.column[1] + 1) - colStartPx(visibledatacolumn, titles.column[0])
        : 0;
    const scaledPaper = {
        pageW: paper.pageW,
        pageH: paper.pageH,
        innerW: Math.max(20, paper.innerW / scale - titleW),
        innerH: Math.max(20, paper.innerH / scale - titleH),
        pad: paper.pad,
    };
    const pages = paginateByPaper(range, visibledatarow, visibledatacolumn, scaledPaper, Object.assign({}, layout, {
        merge: ((file && file.config) || {}).merge || {},
    }));
    return {
        file: file,
        range: range,
        paper: paper,
        pages: pages,
        scale: scale,
        visibledatarow: visibledatarow,
        visibledatacolumn: visibledatacolumn,
        titles: titles,
    };
}

export function buildSessionPages(session) {
    const entries = [];
    session.targets.forEach(function (target) {
        target.ranges.forEach(function (range) {
            const pack = buildPagesForTarget({
                file: target.file,
                range: range,
                visibledatarow: target.visibledatarow,
                visibledatacolumn: target.visibledatacolumn,
            }, session.layout, session.render);
            pack.pages.forEach(function (page, sheetPage) {
                entries.push({ file: pack.file, page: page, pack: pack, sheetPage: sheetPage + 1 });
            });
        });
    });
    return {
        session: session,
        entries: entries,
        pageCount: entries.length,
        paper: entries.length ? entries[0].pack.paper : resolvePaperPx(session.layout, {}),
    };
}

export function createPreparedPrintSession(layout, render) {
    const session = createPrintSession(layout, render, function (area, snapshotFile, sourceFile) {
        if (area === PrintArea.CurrentSelection || area === PrintArea.AllSelection) {
            return null;
        }
        const po = getPrintOptions(snapshotFile);
        if (po.PrintArea) {
            return parseRangeText(po.PrintArea);
        }
        return usedRange(snapshotFile.data || []);
    });
    return preparePrintSession(session).then(function (prepared) {
        return runInPrintSession(session, function () {
            return { session: session, resource: prepared.resource, plan: buildSessionPages(session) };
        });
    });
}

export function buildPages(file, layout, render) {
    const files = file ? [file] : resolveTargetFiles(layout);
    if (!files.length) {
        return { pages: [], range: { row: [0, 0], column: [0, 0] }, paper: resolvePaperPx(layout, {}), scale: 1, file: null };
    }
    const primary = files[0];
    const pack = buildPagesForFile(primary, layout, render);
    pack.file = primary;
    if (files.length > 1) {
        pack.workbookPages = [];
        files.forEach(function (f) {
            const sub = buildPagesForFile(f, layout, render);
            sub.pages.forEach(function (page) {
                pack.workbookPages.push({ file: f, page: page, pack: sub });
            });
        });
    }
    return pack;
}

export function buildAllPageCanvases(pack, layout, render) {
    const canvases = [];
    const pages = pack.workbookPages || pack.pages.map(function (page) {
        return { file: pack.file, page: page, pack: pack };
    });
    pages.forEach(function (entry, index) {
        const meta = {
            pageIndex: index,
            pageTotal: pages.length,
            sheetPage: pack.pages.indexOf(entry.page) + 1,
            sheetPageTotal: entry.pack.pages.length,
        };
        canvases.push(drawPageCanvas(entry.page, entry.file, layout, render, entry.pack, meta));
    });
    return canvases;
}

export function preparePrint(file, layout, render) {
    return waitPrintResources(render.draft).then(function () {
        return preloadPrintImages(file);
    });
}

export function setPrintTitles(which) {
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

export function updatePrintConfig(config) {
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
        if (config.marginCustom) {
            layout.marginCustom = config.marginCustom;
        }
        if (config.pageSizeCustom) {
            layout.pageSizeCustom = config.pageSizeCustom;
        }
        if (config.freeze) {
            layout.freeze = config.freeze.slice ? config.freeze.slice() : config.freeze;
        }
        if (config.subUnitIds) {
            layout.subUnitIds = config.subUnitIds;
        }
        if (config.maxRowsEachPage != null) {
            layout.maxRowsEachPage = config.maxRowsEachPage;
        }
        if (config.maxColumnsEachPage != null) {
            layout.maxColumnsEachPage = config.maxColumnsEachPage;
        }
        if (config.pageOrder != null) {
            layout.pageOrder = config.pageOrder;
        }
        if (config.rangeText === null || config.printArea === null || config.PrintArea === null) {
            layout.rangeText = null;
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

export function updatePrintRenderConfig(config) {
    const file = currentFile();
    if (!file) {
        return null;
    }
    const render = renderOf(file, config || {});
    persist(file, layoutOf(file), render);
    return render;
}
