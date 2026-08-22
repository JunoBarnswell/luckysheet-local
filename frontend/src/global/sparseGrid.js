const SPARSE_GRID = "__luckysheetSparseGrid";
const SPARSE_ROW = "__luckysheetSparseRow";

function toIndex(prop) {
    if (typeof prop === "number") {
        return Number.isInteger(prop) && prop >= 0 ? prop : null;
    }
    if (typeof prop === "string" && prop.length > 0 && prop.indexOf(".") === -1 && /^\d+$/.test(prop)) {
        return parseInt(prop, 10);
    }
    return null;
}

function cloneCell(cell) {
    if (cell == null || typeof cell !== "object") {
        return cell;
    }
    return JSON.parse(JSON.stringify(cell));
}

function ingestRowMap(rowValue, colCount) {
    const rowMap = new Map();
    if (rowValue == null) {
        return { rowMap, colCount };
    }
    if (isSparseRow(rowValue)) {
        const src = rowValue[SPARSE_ROW].map;
        if (src != null) {
            src.forEach(function (v, c) {
                if (v != null) {
                    rowMap.set(c, v);
                    if (c + 1 > colCount) {
                        colCount = c + 1;
                    }
                }
            });
        }
        return { rowMap, colCount };
    }
    const len = rowValue.length != null ? rowValue.length : 0;
    if (len > colCount) {
        colCount = len;
    }
    for (let c = 0; c < len; c++) {
        const v = rowValue[c];
        if (v != null) {
            rowMap.set(c, v);
        }
    }
    return { rowMap, colCount };
}

function applyRowSplice(state, start, deleteCount, items) {
    if (start < 0) {
        start = 0;
    }
    if (start > state.rowCount) {
        start = state.rowCount;
    }
    if (deleteCount < 0) {
        deleteCount = 0;
    }
    if (start + deleteCount > state.rowCount) {
        deleteCount = state.rowCount - start;
    }

    const insertCount = items.length;
    const delta = insertCount - deleteCount;
    const moving = [];

    state.store.forEach(function (rowMap, r) {
        if (r >= start && r < start + deleteCount) {
            state.store.delete(r);
        }
        else if (r >= start + deleteCount) {
            moving.push([r, rowMap]);
            state.store.delete(r);
        }
    });

    for (let i = 0; i < moving.length; i++) {
        state.store.set(moving[i][0] + delta, moving[i][1]);
    }

    for (let i = 0; i < items.length; i++) {
        const ingested = ingestRowMap(items[i], state.colCount);
        state.colCount = ingested.colCount;
        if (ingested.rowMap.size > 0) {
            state.store.set(start + i, ingested.rowMap);
        }
        else {
            state.store.delete(start + i);
        }
    }

    state.rowCount += delta;
    state.rowCache.clear();
    return [];
}

function applyColSplice(rowMap, start, deleteCount, items) {
    if (rowMap == null) {
        rowMap = new Map();
    }
    if (start < 0) {
        start = 0;
    }
    if (deleteCount < 0) {
        deleteCount = 0;
    }

    const insertCount = items.length;
    const delta = insertCount - deleteCount;
    const moving = [];

    rowMap.forEach(function (value, c) {
        if (c >= start && c < start + deleteCount) {
            rowMap.delete(c);
        }
        else if (c >= start + deleteCount) {
            moving.push([c, value]);
            rowMap.delete(c);
        }
    });

    for (let i = 0; i < moving.length; i++) {
        rowMap.set(moving[i][0] + delta, moving[i][1]);
    }

    for (let i = 0; i < items.length; i++) {
        if (items[i] != null) {
            rowMap.set(start + i, items[i]);
        }
    }

    return rowMap;
}

function createSparseRow(state, r) {
    if (state.rowCache.has(r)) {
        return state.rowCache.get(r);
    }

    const target = [];
    const rowApi = {
        splice: function (start, deleteCount) {
            const items = Array.prototype.slice.call(arguments, 2);
            let rowMap = state.store.get(r);
            rowMap = applyColSplice(rowMap, start, deleteCount, items);
            if (rowMap.size > 0) {
                state.store.set(r, rowMap);
            }
            else {
                state.store.delete(r);
            }
            if (r === 0) {
                const next = state.colCount - Math.max(0, deleteCount) + items.length;
                state.colCount = next < 0 ? 0 : next;
            }
            return [];
        },
        concat: function () {
            return Array.prototype.concat.apply(createSparseRow(state, r), arguments);
        }
    };

    const row = new Proxy(target, {
        get: function (t, prop, receiver) {
            if (prop === SPARSE_ROW) {
                return { state: state, r: r, map: state.store.get(r) };
            }
            if (prop === "length") {
                return state.colCount;
            }
            if (prop === Symbol.isConcatSpreadable) {
                return true;
            }
            if (prop === Symbol.toStringTag) {
                return "Array";
            }
            if (prop === "splice") {
                return rowApi.splice;
            }
            if (typeof prop === "string" && typeof Array.prototype[prop] === "function") {
                return function () {
                    return Array.prototype[prop].apply(receiver, arguments);
                };
            }
            const c = toIndex(prop);
            if (c == null) {
                return Reflect.get(t, prop, receiver);
            }
            if (c < 0 || c >= state.colCount) {
                return undefined;
            }
            const rowMap = state.store.get(r);
            if (rowMap == null || !rowMap.has(c)) {
                return null;
            }
            return rowMap.get(c);
        },
        set: function (t, prop, value) {
            if (prop === "length") {
                const next = Number(value);
                if (Number.isFinite(next) && next >= 0) {
                    if (next < state.colCount) {
                        state.store.forEach(function (rowMap) {
                            rowMap.forEach(function (_v, c) {
                                if (c >= next) {
                                    rowMap.delete(c);
                                }
                            });
                        });
                    }
                    state.colCount = next;
                }
                return true;
            }
            const c = toIndex(prop);
            if (c == null) {
                t[prop] = value;
                return true;
            }
            if (c >= state.colCount) {
                state.colCount = c + 1;
            }
            if (value == null) {
                const rowMap = state.store.get(r);
                if (rowMap != null) {
                    rowMap.delete(c);
                    if (rowMap.size === 0) {
                        state.store.delete(r);
                    }
                }
                return true;
            }
            let rowMap = state.store.get(r);
            if (rowMap == null) {
                rowMap = new Map();
                state.store.set(r, rowMap);
            }
            rowMap.set(c, value);
            return true;
        },
        has: function (t, prop) {
            if (prop === "length" || prop === SPARSE_ROW) {
                return true;
            }
            const c = toIndex(prop);
            if (c != null) {
                return c >= 0 && c < state.colCount;
            }
            return Reflect.has(t, prop);
        },
        ownKeys: function () {
            return ["length"];
        },
        getOwnPropertyDescriptor: function (t, prop) {
            if (prop === "length") {
                return { configurable: true, enumerable: false, writable: true, value: state.colCount };
            }
            const c = toIndex(prop);
            if (c != null && c >= 0 && c < state.colCount) {
                return { configurable: true, enumerable: false, writable: true, value: null };
            }
            return undefined;
        }
    });

    state.rowCache.set(r, row);
    return row;
}

function makeGridProxy(state) {
    const target = {};

    const gridApi = {
        splice: function (start, deleteCount) {
            const items = Array.prototype.slice.call(arguments, 2);
            applyRowSplice(state, start, deleteCount, items);
            return [];
        },
        push: function () {
            const items = Array.prototype.slice.call(arguments);
            applyRowSplice(state, state.rowCount, 0, items);
            return state.rowCount;
        },
        unshift: function () {
            const items = Array.prototype.slice.call(arguments);
            applyRowSplice(state, 0, 0, items);
            return state.rowCount;
        },
        clone: function () {
            return cloneSheetData(grid);
        },
        toCelldata: function () {
            return sparseGridToCelldata(grid);
        },
        slice: function (start, end) {
            if (start == null) {
                start = 0;
            }
            if (end == null) {
                end = state.rowCount;
            }
            if (start < 0) {
                start = state.rowCount + start;
            }
            if (end < 0) {
                end = state.rowCount + end;
            }
            start = Math.max(0, start);
            end = Math.min(state.rowCount, end);
            const next = createSparseGrid(Math.max(0, end - start), state.colCount);
            const dst = getSparseState(next);
            state.store.forEach(function (rowMap, r) {
                if (r >= start && r < end) {
                    const copied = new Map();
                    rowMap.forEach(function (cell, c) {
                        copied.set(c, cloneCell(cell));
                    });
                    dst.store.set(r - start, copied);
                }
            });
            return next;
        }
    };

    const grid = new Proxy(target, {
        get: function (t, prop, receiver) {
            if (prop === SPARSE_GRID) {
                return state;
            }
            if (prop === "length") {
                return state.rowCount;
            }
            if (prop === Symbol.isConcatSpreadable) {
                return false;
            }
            if (prop === Symbol.toStringTag) {
                return "Array";
            }
            if (prop === "splice") {
                return gridApi.splice;
            }
            if (prop === "push") {
                return gridApi.push;
            }
            if (prop === "unshift") {
                return gridApi.unshift;
            }
            if (prop === "clone") {
                return gridApi.clone;
            }
            if (prop === "toCelldata") {
                return gridApi.toCelldata;
            }
            if (prop === "slice") {
                return gridApi.slice;
            }
            if (prop === "toJSON") {
                return function () {
                    return undefined;
                };
            }
            const r = toIndex(prop);
            if (r == null) {
                if (typeof prop === "string" && typeof Array.prototype[prop] === "function") {
                    return function () {
                        return Array.prototype[prop].apply(receiver, arguments);
                    };
                }
                return Reflect.get(t, prop, receiver);
            }
            if (r < 0 || r >= state.rowCount) {
                return undefined;
            }
            return createSparseRow(state, r);
        },
        set: function (t, prop, value) {
            if (prop === "length") {
                const next = Number(value);
                if (Number.isFinite(next) && next >= 0) {
                    if (next < state.rowCount) {
                        state.store.forEach(function (_rowMap, r) {
                            if (r >= next) {
                                state.store.delete(r);
                            }
                        });
                    }
                    state.rowCount = next;
                    state.rowCache.clear();
                }
                return true;
            }
            const r = toIndex(prop);
            if (r == null) {
                t[prop] = value;
                return true;
            }
            if (r >= state.rowCount) {
                state.rowCount = r + 1;
            }
            const ingested = ingestRowMap(value, state.colCount);
            state.colCount = ingested.colCount;
            if (ingested.rowMap.size > 0) {
                state.store.set(r, ingested.rowMap);
            }
            else {
                state.store.delete(r);
            }
            state.rowCache.delete(r);
            return true;
        },
        has: function (t, prop) {
            if (prop === "length" || prop === SPARSE_GRID) {
                return true;
            }
            const r = toIndex(prop);
            if (r != null) {
                return r >= 0 && r < state.rowCount;
            }
            return Reflect.has(t, prop);
        },
        ownKeys: function () {
            return ["length"];
        },
        getOwnPropertyDescriptor: function (t, prop) {
            if (prop === "length") {
                return { configurable: true, enumerable: false, writable: true, value: state.rowCount };
            }
            const r = toIndex(prop);
            if (r != null && r >= 0 && r < state.rowCount) {
                return { configurable: true, enumerable: false, writable: true, value: createSparseRow(state, r) };
            }
            return undefined;
        }
    });

    return grid;
}

export function isSparseGrid(data) {
    return data != null && data[SPARSE_GRID] != null;
}

export function isSparseRow(data) {
    return data != null && data[SPARSE_ROW] != null;
}

export function getSparseState(data) {
    return data != null ? data[SPARSE_GRID] : null;
}

export function createSparseGrid(rowCount, colCount) {
    return makeGridProxy({
        rowCount: Math.max(0, rowCount | 0),
        colCount: Math.max(0, colCount | 0),
        store: new Map(),
        rowCache: new Map()
    });
}

export function createSparseGridFromDense(data, rowCount, colCount) {
    const rows = data != null && data.length > 0 ? data.length : 0;
    let cols = 0;
    if (data != null) {
        for (let r = 0; r < data.length; r++) {
            if (data[r] != null && data[r].length > cols) {
                cols = data[r].length;
            }
        }
    }
    const grid = createSparseGrid(
        Math.max(rowCount || 0, rows),
        Math.max(colCount || 0, cols)
    );
    if (data == null) {
        return grid;
    }
    const state = getSparseState(grid);
    for (let r = 0; r < data.length; r++) {
        const row = data[r];
        if (row == null) {
            continue;
        }
        const ingested = ingestRowMap(row, state.colCount);
        state.colCount = ingested.colCount;
        if (ingested.rowMap.size > 0) {
            state.store.set(r, ingested.rowMap);
        }
    }
    return grid;
}

export function createSparseGridFromCelldata(celldata, rowCount, colCount) {
    let maxR = Math.max(0, (rowCount || 0) - 1);
    let maxC = Math.max(0, (colCount || 0) - 1);
    if (celldata != null) {
        for (let i = 0; i < celldata.length; i++) {
            const item = celldata[i];
            if (item == null) {
                continue;
            }
            if (item.r > maxR) {
                maxR = item.r;
            }
            if (item.c > maxC) {
                maxC = item.c;
            }
        }
    }
    const grid = createSparseGrid(maxR + 1, maxC + 1);
    const state = getSparseState(grid);
    if (celldata != null) {
        for (let i = 0; i < celldata.length; i++) {
            const item = celldata[i];
            if (item == null || item.v == null) {
                continue;
            }
            let rowMap = state.store.get(item.r);
            if (rowMap == null) {
                rowMap = new Map();
                state.store.set(item.r, rowMap);
            }
            rowMap.set(item.c, item.v);
        }
    }
    return grid;
}

export function asSparseGrid(data, rowCount, colCount) {
    if (isSparseGrid(data)) {
        if (rowCount != null || colCount != null) {
            ensureSparseSize(
                data,
                rowCount != null ? rowCount - 1 : -1,
                colCount != null ? colCount - 1 : -1
            );
        }
        return data;
    }
    if (data == null || data.length == 0) {
        return createSparseGrid(rowCount || 0, colCount || 0);
    }
    return createSparseGridFromDense(data, rowCount, colCount);
}

export function ensureSparseSize(data, r, c) {
    if (data == null) {
        return data;
    }
    if (isSparseGrid(data)) {
        const state = getSparseState(data);
        if (r != null && r >= state.rowCount) {
            state.rowCount = r + 1;
        }
        if (c != null && c >= state.colCount) {
            state.colCount = c + 1;
        }
        return data;
    }
    return data;
}

export function cloneSheetData(data) {
    if (data == null) {
        return createSparseGrid(0, 0);
    }
    if (isSparseGrid(data)) {
        const src = getSparseState(data);
        const grid = createSparseGrid(src.rowCount, src.colCount);
        const dst = getSparseState(grid);
        src.store.forEach(function (rowMap, r) {
            const next = new Map();
            rowMap.forEach(function (cell, c) {
                next.set(c, cloneCell(cell));
            });
            dst.store.set(r, next);
        });
        return grid;
    }
    return createSparseGridFromDense(data);
}

export function sparseGridToCelldata(data) {
    const ret = [];
    if (data == null) {
        return ret;
    }
    if (isSparseGrid(data)) {
        const state = getSparseState(data);
        const rows = Array.from(state.store.keys()).sort(function (a, b) { return a - b; });
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const rowMap = state.store.get(r);
            const cols = Array.from(rowMap.keys()).sort(function (a, b) { return a - b; });
            for (let j = 0; j < cols.length; j++) {
                const c = cols[j];
                const v = rowMap.get(c);
                if (v != null) {
                    ret.push({ r: r, c: c, v: v });
                }
            }
        }
        return ret;
    }
    for (let r = 0; r < data.length; r++) {
        const row = data[r];
        if (row == null) {
            continue;
        }
        if (isSparseRow(row)) {
            const rowMap = row[SPARSE_ROW].map;
            if (rowMap == null) {
                continue;
            }
            rowMap.forEach(function (v, c) {
                if (v != null) {
                    ret.push({ r: r, c: c, v: v });
                }
            });
            continue;
        }
        for (let c = 0; c < row.length; c++) {
            if (row[c] != null) {
                ret.push({ r: r, c: c, v: row[c] });
            }
        }
    }
    return ret;
}

export function occupiedCellCount(data) {
    if (isSparseGrid(data)) {
        let n = 0;
        getSparseState(data).store.forEach(function (rowMap) {
            n += rowMap.size;
        });
        return n;
    }
    if (data == null) {
        return 0;
    }
    let n = 0;
    for (let r = 0; r < data.length; r++) {
        const row = data[r];
        if (row == null) {
            continue;
        }
        for (let c = 0; c < row.length; c++) {
            if (row[c] != null) {
                n++;
            }
        }
    }
    return n;
}

export function estimateViewportRect(rowheight, colwidth, Store) {
    let visRows = 40;
    let visCols = 20;
    if (Store != null) {
        const hw = Store.luckysheetTableContentHW;
        const rowlen = Math.max(Store.defaultrowlen || 19, 1);
        const collen = Math.max(Store.defaultcollen || 73, 1);
        if (hw != null && hw[1] > 0) {
            visRows = Math.ceil(hw[1] / rowlen) + 4;
        }
        if (hw != null && hw[0] > 0) {
            visCols = Math.ceil(hw[0] / collen) + 4;
        }
    }
    return {
        row: [0, Math.max(0, Math.min(rowheight || visRows, visRows) - 1)],
        column: [0, Math.max(0, Math.min(colwidth || visCols, visCols) - 1)]
    };
}

export function growSparseToRect(data, rect) {
    if (!isSparseGrid(data) || rect == null) {
        return data;
    }
    const r2 = rect.row != null ? rect.row[1] : -1;
    const c2 = rect.column != null ? rect.column[1] : -1;
    return ensureSparseSize(data, r2, c2);
}
