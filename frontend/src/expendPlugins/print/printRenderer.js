import Store from "../../store";
import { getBorderInfoCompute } from "../../global/border";
import { chatatABC } from "../../utils/util";
import { collectChartCanvases, loadImage } from "./printResources";
import {
    cellDisplayValue,
    rowStartPx,
    colStartPx,
    rowHeight,
    colWidth,
    resolveHeaderFooterText,
    parsePrintTitles,
} from "./printLayout";

function mergeMap(config) {
    return (config && config.merge) || {};
}

function mergeKey(r, c) {
    return r + "_" + c;
}

function findMergeMaster(merge, r, c) {
    const key = mergeKey(r, c);
    if (merge[key]) {
        return { r: r, c: c, mc: merge[key] };
    }
    const keys = Object.keys(merge);
    for (let i = 0; i < keys.length; i++) {
        const parts = keys[i].split("_");
        const mr = Number(parts[0]);
        const mc = Number(parts[1]);
        const span = merge[keys[i]];
        if (r >= mr && r <= mr + span.rs - 1 && c >= mc && c <= mc + span.cs - 1) {
            return { r: mr, c: mc, mc: span };
        }
    }
    return null;
}

function isMergeSlave(merge, r, c) {
    const master = findMergeMaster(merge, r, c);
    if (!master) {
        return false;
    }
    return master.r !== r || master.c !== c;
}

function mergeSpanPx(visibledatarow, visibledatacolumn, master, span) {
    const r0 = master.r;
    const c0 = master.c;
    const r1 = r0 + span.rs - 1;
    const c1 = c0 + span.cs - 1;
    const w = colStartPx(visibledatacolumn, c1 + 1) - colStartPx(visibledatacolumn, c0);
    const h = rowStartPx(visibledatarow, r1 + 1) - rowStartPx(visibledatarow, r0);
    return { w: Math.max(w, 0), h: Math.max(h, 0) };
}

function drawBorderEdge(ctx, edge, x, y, w, h) {
    if (!edge || edge.style == null || edge.style === 0) {
        return;
    }
    ctx.strokeStyle = edge.color || "#374151";
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (edge.side === "t") {
        ctx.moveTo(x, y + 0.5);
        ctx.lineTo(x + w, y + 0.5);
    } else if (edge.side === "b") {
        ctx.moveTo(x, y + h - 0.5);
        ctx.lineTo(x + w, y + h - 0.5);
    } else if (edge.side === "l") {
        ctx.moveTo(x + 0.5, y);
        ctx.lineTo(x + 0.5, y + h);
    } else if (edge.side === "r") {
        ctx.moveTo(x + w - 0.5, y);
        ctx.lineTo(x + w - 0.5, y + h);
    }
    ctx.stroke();
}

function drawCellBorders(ctx, borders, x, y, w, h) {
    if (!borders) {
        return;
    }
    ["t", "b", "l", "r"].forEach(function (side) {
        if (borders[side]) {
            drawBorderEdge(ctx, Object.assign({ side: side }, borders[side]), x, y, w, h);
        }
    });
}

function textAlignFromCell(cell, render) {
    const ht = cell && cell.ht;
    if (ht === 0) {
        return "center";
    }
    if (ht === 2) {
        return "right";
    }
    if (render.hAlign === "Middle" || render.hAlign === "End") {
        return render.hAlign === "End" ? "right" : "center";
    }
    return "left";
}

function textBaselineFromCell(cell, render, h) {
    const vt = cell && cell.vt;
    if (vt === 0) {
        return { baseline: "middle", y: h / 2 };
    }
    if (vt === 2) {
        return { baseline: "bottom", y: h - 2 };
    }
    if (render.vAlign === "Middle") {
        return { baseline: "middle", y: h / 2 };
    }
    if (render.vAlign === "End") {
        return { baseline: "bottom", y: h - 2 };
    }
    return { baseline: "top", y: 2 };
}

function drawCellText(ctx, cell, x, y, w, h) {
    const text = cellDisplayValue(cell);
    if (!text) {
        return;
    }
    ctx.save();
    ctx.beginPath();
    ctx.rect(x + 2, y + 1, Math.max(w - 4, 0), Math.max(h - 2, 0));
    ctx.clip();
    ctx.fillStyle = (cell && cell.fc) || "#111827";
    const fs = (cell && cell.fs) || 11;
    const italic = cell && cell.it ? "italic " : "";
    const bold = cell && cell.bl ? "bold " : "";
    ctx.font = italic + bold + fs + "px sans-serif";
    const align = textAlignFromCell(cell, {});
    ctx.textAlign = align;
    const base = textBaselineFromCell(cell, {}, h);
    ctx.textBaseline = base.baseline;
    let tx = x + 3;
    if (align === "center") {
        tx = x + w / 2;
    } else if (align === "right") {
        tx = x + w - 3;
    }
    const wrap = cell && cell.tb === 2;
    if (wrap && text.length > 0) {
        const words = text.split("");
        let line = "";
        let ly = y + 2;
        const lineH = fs + 2;
        for (let i = 0; i < words.length; i++) {
            const test = line + words[i];
            if (ctx.measureText(test).width > w - 6 && line) {
                ctx.fillText(line, tx, ly + fs);
                line = words[i];
                ly += lineH;
                if (ly + lineH > y + h) {
                    break;
                }
            } else {
                line = test;
            }
        }
        if (line && ly + fs <= y + h) {
            ctx.fillText(line, tx, ly + fs);
        }
    } else {
        ctx.fillText(text, tx, y + base.y);
    }
    ctx.restore();
}

function drawSheetRegion(ctx, file, page, pack, render, options) {
    const data = (file && file.data) || Store.flowdata || [];
    const merge = mergeMap((file && file.config) || Store.config);
    const sheetIndex = file && file.index != null ? file.index : Store.currentSheetIndex;
    let borderInfoCompute = {};
    try {
        borderInfoCompute = getBorderInfoCompute(sheetIndex) || {};
    } catch (e) {
        borderInfoCompute = {};
    }

    const r0 = page.row[0];
    const r1 = page.row[1];
    const c0 = page.column[0];
    const c1 = page.column[1];
    const originX = options.originX != null ? options.originX : colStartPx(pack.visibledatacolumn, c0);
    const originY = options.originY != null ? options.originY : rowStartPx(pack.visibledatarow, r0);
    const offsetX = options.offsetX || 0;
    const offsetY = options.offsetY || 0;
    const headingW = options.headingW || 0;
    const headingH = options.headingH || 0;

    for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
            if (isMergeSlave(merge, r, c)) {
                continue;
            }
            const master = findMergeMaster(merge, r, c);
            let w = colWidth(pack.visibledatacolumn, c);
            let h = rowHeight(pack.visibledatarow, r);
            if (master && master.mc) {
                const span = mergeSpanPx(pack.visibledatarow, pack.visibledatacolumn, master, master.mc);
                w = span.w;
                h = span.h;
            }
            const x = headingW + offsetX + colStartPx(pack.visibledatacolumn, c) - originX;
            const y = headingH + offsetY + rowStartPx(pack.visibledatarow, r) - originY;
            const cell = data[r] && data[r][c];
            if (cell && cell.bg) {
                ctx.fillStyle = cell.bg;
                ctx.fillRect(x, y, w, h);
            }
            const bdKey = r + "_" + c;
            const borders = (cell && cell.bd) || borderInfoCompute[bdKey];
            if (render.gridlines && !borders) {
                ctx.strokeStyle = "#d1d5db";
                ctx.lineWidth = 1;
                ctx.strokeRect(x + 0.5, y + 0.5, Math.max(w - 1, 0), Math.max(h - 1, 0));
            }
            drawCellBorders(ctx, borders, x, y, w, h);
            drawCellText(ctx, cell, x, y, w, h);
        }
    }
}

function drawHeadings(ctx, page, pack, headingW, headingH, originX, originY, offsetX, offsetY) {
    const r0 = page.row[0];
    const r1 = page.row[1];
    const c0 = page.column[0];
    const c1 = page.column[1];
    ctx.fillStyle = "#f3f4f6";
    ctx.fillRect(offsetX, offsetY, headingW + colStartPx(pack.visibledatacolumn, c1 + 1) - originX, headingH);
    ctx.fillRect(offsetX, offsetY, headingW, headingH + rowStartPx(pack.visibledatarow, r1 + 1) - originY);
    ctx.fillStyle = "#111827";
    ctx.font = "11px sans-serif";
    ctx.textBaseline = "middle";
    for (let c = c0; c <= c1; c++) {
        const x = headingW + offsetX + colStartPx(pack.visibledatacolumn, c) - originX;
        ctx.fillText(chatatABC(c), x + 4, offsetY + headingH / 2);
    }
    for (let r = r0; r <= r1; r++) {
        const y = headingH + offsetY + rowStartPx(pack.visibledatarow, r) - originY;
        const h = rowHeight(pack.visibledatarow, r);
        ctx.fillText(String(r + 1), offsetX + 4, y + h / 2);
    }
}

function drawImagesSync(ctx, file, pack, page, headingW, headingH, originX, originY) {
    if (!file || !file.images || page.renderDraft) {
        return;
    }
    const r0 = page.row[0];
    const r1 = page.row[1];
    const c0 = page.column[0];
    const c1 = page.column[1];
    const regionRight = colStartPx(pack.visibledatacolumn, c1 + 1);
    const regionBottom = rowStartPx(pack.visibledatarow, r1 + 1);
    Object.keys(file.images).forEach(function (id) {
        const imgMeta = file.images[id];
        if (!imgMeta || !imgMeta.src) {
            return;
        }
        const left = (imgMeta.default && imgMeta.default.left) || imgMeta.left || 0;
        const top = (imgMeta.default && imgMeta.default.top) || imgMeta.top || 0;
        const iw = (imgMeta.default && imgMeta.default.width) || imgMeta.width || 0;
        const ih = (imgMeta.default && imgMeta.default.height) || imgMeta.height || 0;
        if (left + iw < originX || top + ih < originY || left > regionRight || top > regionBottom) {
            return;
        }
        const cached = imgMeta._printImage;
        if (cached && cached.complete) {
            ctx.drawImage(cached, headingW + left - originX, headingH + top - originY, iw, ih);
        }
    });
}

function drawChartsSync(ctx, pack, page, headingW, headingH, originX, originY, containerRect) {
    const charts = collectChartCanvases();
    const r0 = page.row[0];
    const r1 = page.row[1];
    const c0 = page.column[0];
    const c1 = page.column[1];
    const regionRight = colStartPx(pack.visibledatacolumn, c1 + 1);
    const regionBottom = rowStartPx(pack.visibledatarow, r1 + 1);
    charts.forEach(function (item) {
        const nodeRect = item.node.getBoundingClientRect();
        const relLeft = nodeRect.left - containerRect.left;
        const relTop = nodeRect.top - containerRect.top;
        if (relLeft + item.width < originX || relTop + item.height < originY) {
            return;
        }
        if (relLeft > regionRight || relTop > regionBottom) {
            return;
        }
        ctx.drawImage(item.canvas, headingW + relLeft - originX, headingH + relTop - originY, item.width, item.height);
    });
}

function drawWatermark(ctx, watermark, pageW, pageH, enforce) {
    if (!watermark && !enforce) {
        return;
    }
    const text = typeof watermark === "string" ? watermark : watermark && watermark.text;
    const content = text || (enforce ? "LuckySheet" : "");
    if (!content) {
        return;
    }
    ctx.save();
    ctx.globalAlpha = (watermark && watermark.opacity) || 0.12;
    ctx.fillStyle = (watermark && watermark.color) || "#6b7280";
    ctx.font = "bold 48px sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.translate(pageW / 2, pageH / 2);
    ctx.rotate(-Math.PI / 6);
    ctx.fillText(content, 0, 0);
    ctx.restore();
}

function drawHeaderFooter(ctx, pack, render, file, pageMeta) {
    const setting = render.headerFooterSetting || {};
    const ctxInfo = {
        workbookTitle: (Store.toJsonOptions && Store.toJsonOptions.title) || "",
        worksheetTitle: (file && file.name) || "",
        page: (pageMeta && pageMeta.pageIndex != null ? pageMeta.pageIndex : 0) + 1,
        pageTotal: (pageMeta && pageMeta.pageTotal) || pack.pages.length,
        sheetPage: (pageMeta && pageMeta.sheetPage) || 1,
        sheetPageTotal: (pageMeta && pageMeta.sheetPageTotal) || pack.pages.length,
    };
    const topLeft = resolveHeaderFooterText(setting.topLeft, ctxInfo);
    const topCenter = resolveHeaderFooterText(setting.topCenter, ctxInfo);
    const topRight = resolveHeaderFooterText(setting.topRight, ctxInfo);
    const bottomLeft = resolveHeaderFooterText(setting.bottomLeft, ctxInfo);
    const bottomCenter = resolveHeaderFooterText(setting.bottomCenter, ctxInfo);
    const bottomRight = resolveHeaderFooterText(setting.bottomRight, ctxInfo);
    ctx.fillStyle = "#4b5563";
    ctx.font = "10px sans-serif";
    ctx.textBaseline = "top";
    const pad = pack.paper.pad;
    const pageW = pack.paper.pageW;
    const pageH = pack.paper.pageH;
    if (topLeft) {
        ctx.textAlign = "left";
        ctx.fillText(topLeft, pad.left, 8);
    }
    if (topCenter) {
        ctx.textAlign = "center";
        ctx.fillText(topCenter, pageW / 2, 8);
    }
    if (topRight) {
        ctx.textAlign = "right";
        ctx.fillText(topRight, pageW - pad.right, 8);
    }
    if (bottomLeft) {
        ctx.textAlign = "left";
        ctx.fillText(bottomLeft, pad.left, pageH - 18);
    }
    if (bottomCenter) {
        ctx.textAlign = "center";
        ctx.fillText(bottomCenter, pageW / 2, pageH - 18);
    }
    if (bottomRight) {
        ctx.textAlign = "right";
        ctx.fillText(bottomRight, pageW - pad.right, pageH - 18);
    }
}

export function preloadPrintImages(file) {
    if (!file || !file.images) {
        return Promise.resolve();
    }
    const tasks = Object.keys(file.images).map(function (id) {
        const imgMeta = file.images[id];
        const src = imgMeta && imgMeta.src;
        if (!src) {
            return Promise.resolve();
        }
        return loadImage(src).then(function (img) {
            if (img) {
                imgMeta._printImage = img;
            }
        });
    });
    return Promise.all(tasks);
}

export function drawPageCanvas(page, file, layout, render, pack, pageMeta) {
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

    const titles = parsePrintTitles(file, layout);
    const headingW = render.headings ? 36 : 0;
    const headingH = render.headings ? 20 : 0;
    const originX = colStartPx(pack.visibledatacolumn, page.column[0]);
    const originY = rowStartPx(pack.visibledatarow, page.row[0]);

    let titleRowH = 0;
    let titleColW = 0;
    if (titles.row) {
        titleRowH = rowStartPx(pack.visibledatarow, titles.row[1] + 1) - rowStartPx(pack.visibledatarow, titles.row[0]);
    }
    if (titles.column) {
        titleColW = colStartPx(pack.visibledatacolumn, titles.column[1] + 1) - colStartPx(pack.visibledatacolumn, titles.column[0]);
    }

    const contentOriginX = originX;
    const contentOriginY = originY;

    if (render.headings) {
        drawHeadings(ctx, page, pack, headingW, headingH, contentOriginX, contentOriginY, 0, 0);
    }

    if (titles.row) {
        drawSheetRegion(
            ctx,
            file,
            { row: titles.row, column: page.column },
            pack,
            render,
            {
                originX: colStartPx(pack.visibledatacolumn, titles.row ? page.column[0] : page.column[0]),
                originY: rowStartPx(pack.visibledatarow, titles.row[0]),
                offsetX: headingW + titleColW,
                offsetY: headingH,
                headingW: 0,
                headingH: 0,
            }
        );
    }

    if (titles.column) {
        drawSheetRegion(
            ctx,
            file,
            { row: page.row, column: titles.column },
            pack,
            render,
            {
                originX: colStartPx(pack.visibledatacolumn, titles.column[0]),
                originY: rowStartPx(pack.visibledatarow, page.row[0]),
                offsetX: headingW,
                offsetY: headingH + titleRowH,
                headingW: 0,
                headingH: 0,
            }
        );
    }

    ctx.beginPath();
    ctx.rect(
        headingW + titleColW,
        headingH + titleRowH,
        pack.paper.innerW / pack.scale - titleColW,
        pack.paper.innerH / pack.scale - titleRowH
    );
    ctx.clip();

    drawSheetRegion(ctx, file, page, pack, render, {
        originX: contentOriginX,
        originY: contentOriginY,
        offsetX: headingW + titleColW,
        offsetY: headingH + titleRowH,
        headingW: 0,
        headingH: 0,
    });

    if (!render.draft) {
        const root = Store.container ? document.getElementById(Store.container) : null;
        const containerRect = root ? root.getBoundingClientRect() : { left: 0, top: 0 };
        page.renderDraft = false;
        drawImagesSync(ctx, file, pack, page, headingW + titleColW, headingH + titleRowH, contentOriginX, contentOriginY);
        drawChartsSync(ctx, pack, page, headingW + titleColW, headingH + titleRowH, contentOriginX, contentOriginY, containerRect);
    }

    ctx.restore();
    drawHeaderFooter(ctx, pack, render, file, pageMeta);

    const enforce = Store.printPluginConfig && Store.printPluginConfig.enforceWatermark;
    drawWatermark(ctx, render.watermark, pack.paper.pageW, pack.paper.pageH, enforce);

    return canvas;
}

export { findMergeMaster, isMergeSlave };
