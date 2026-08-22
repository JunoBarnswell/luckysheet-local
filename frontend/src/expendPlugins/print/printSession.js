import { getFocusedContext, getUnit, withInstance } from "../../store/registry";
import { getPrintOptions, usedRange, rowHeight, colWidth } from "./printLayout";
import { PrintResourceCollector } from "./printResourceCollector";
import { loadImage, collectChartCanvases, waitForCanvasReady } from "./printResources";

function clone(value) {
    if (Array.isArray(value)) {
        return value.map(clone);
    }
    if (value && Object.prototype.toString.call(value) === "[object Object]") {
        const copy = {};
        Object.keys(value).forEach(function (key) {
            // Image elements are loaded separately and must not be serialised.
            if (key !== "_printImage") {
                copy[key] = clone(value[key]);
            }
        });
        return copy;
    }
    return value;
}

function visibleAxis(file, axis, fallback, defaultSize) {
    const count = Number(file && file[axis]) || (fallback && fallback.length) || 0;
    const config = (file && file.config) || {};
    const sizes = axis === "row" ? config.rowlen || {} : config.columnlen || {};
    const hidden = axis === "row" ? config.rowhidden || {} : config.colhidden || {};
    const result = [];
    let current = 0;
    for (let index = 0; index < count; index++) {
        const size = hidden[index] ? 0 : Number(sizes[index] == null ? defaultSize : sizes[index]);
        current += Math.max(0, size || 0);
        result.push(current);
    }
    return result.length ? result : (fallback || []).slice();
}

function findFile(context, id) {
    if (!context || !context.luckysheetfile) {
        return null;
    }
    return context.luckysheetfile.filter(function (file) {
        return String(file.index) === String(id) || String(file.order) === String(id);
    })[0] || null;
}

function selectionTargets(context, all) {
    const selections = context.luckysheet_select_save || [];
    if (!selections.length) {
        return [];
    }
    const input = all ? selections : [selections[selections.length - 1]];
    return input.map(function (selection) {
        return { row: selection.row.slice(), column: selection.column.slice() };
    });
}

function targetDescriptors(context, layout, selectedFile) {
    const files = context.luckysheetfile || [];
    if (layout.subUnitIds && layout.subUnitIds.length) {
        return layout.subUnitIds.map(function (subUnit) {
            const item = typeof subUnit === "string" ? { id: subUnit } : subUnit;
            const file = findFile(context, item.id);
            return file ? { file: file, ranges: item.range ? [clone(item.range)] : null } : null;
        }).filter(Boolean);
    }
    if (layout.area === "Workbook") {
        return files.map(function (file) { return { file: file, ranges: null }; });
    }
    if (layout.area === "AllSelection") {
        const ranges = selectionTargets(context, true);
        return selectedFile ? [{ file: selectedFile, ranges: ranges }] : [];
    }
    if (layout.area === "CurrentSelection") {
        const ranges = selectionTargets(context, false);
        return selectedFile ? [{ file: selectedFile, ranges: ranges }] : [];
    }
    return selectedFile ? [{ file: selectedFile, ranges: null }] : [];
}

export function createPrintSession(layout, render, rangeResolver) {
    const context = getFocusedContext();
    if (!context) {
        throw new Error("No focused LuckySheet instance for print");
    }
    const selectedFile = findFile(context, context.currentSheetIndex);
    const charts = render.draft ? [] : collectChartCanvases().map(function (entry) {
        return {
            canvas: entry.canvas,
            node: entry.node,
            width: entry.width,
            height: entry.height,
        };
    });
    const targets = targetDescriptors(context, layout, selectedFile).map(function (target) {
        const file = clone(target.file);
        const ranges = target.ranges && target.ranges.length
            ? target.ranges
            : [rangeResolver(layout.area, file, target.file) || usedRange(file.data || [])];
        return {
            file: file,
            sourceSheetId: target.file.index,
            ranges: ranges,
            visibledatarow: String(target.file.index) === String(context.currentSheetIndex)
                ? (context.visibledatarow || []).slice()
                : visibleAxis(target.file, "row", [], context.defaultrowlen || 19),
            visibledatacolumn: String(target.file.index) === String(context.currentSheetIndex)
                ? (context.visibledatacolumn || []).slice()
                : visibleAxis(target.file, "column", [], context.defaultcolwidth || 73),
        };
    });
    return Object.freeze({
        id: "print_" + context.instanceId + "_" + Date.now().toString(36),
        instanceId: context.instanceId,
        workbookTitle: (context.toJsonOptions && context.toJsonOptions.title) || "",
        container: context.container,
        layout: clone(layout),
        render: clone(render),
        targets: targets,
        devicePixelRatio: context.devicePixelRatio || 1,
        charts: charts,
        createdAt: Date.now(),
    });
}

export function preparePrintSession(session) {
    const collector = new PrintResourceCollector();
    if (session.render.draft) {
        return collector.wait().then(function (result) { return { session: session, resource: result }; });
    }
    session.targets.forEach(function (target) {
        Object.keys(target.file.images || {}).forEach(function (id) {
            const imageMeta = target.file.images[id];
            if (!imageMeta || !imageMeta.src) {
                return;
            }
            collector.add(loadImage(imageMeta.src).then(function (image) {
                imageMeta._printImage = image;
                return image;
            }), "image:" + id);
        });
    });
    withInstance(session.instanceId, function () {
        collectChartCanvases().forEach(function (entry, index) {
            collector.add(waitForCanvasReady(entry.canvas), "chart:" + index);
        });
    });
    return collector.wait().then(function (resource) {
        return { session: session, resource: resource };
    });
}

export function runInPrintSession(session, callback) {
    if (!getUnit(session.instanceId)) {
        throw new Error("Print instance was destroyed before completion");
    }
    return withInstance(session.instanceId, callback);
}

export function isVisibleRange(range, rows, columns) {
    for (let r = range.row[0]; r <= range.row[1]; r++) {
        if (rowHeight(rows, r) > 0) {
            for (let c = range.column[0]; c <= range.column[1]; c++) {
                if (colWidth(columns, c) > 0) {
                    return true;
                }
            }
        }
    }
    return false;
}

export { getPrintOptions };
