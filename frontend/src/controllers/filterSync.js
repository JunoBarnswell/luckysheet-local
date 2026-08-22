/**
 * 行列插删/移动后同步 filter_select 与各列 filter 索引。
 * 语义对齐 Univer sheets-filter.controller / sheets-filter-sync (#5797)：
 * 插入列扩展或平移筛选范围，删除列丢弃被删列条件并重排 colId，移动列重映射索引。
 */

function cloneJson(value) {
    if (value == null) {
        return value;
    }
    return JSON.parse(JSON.stringify(value));
}

function isEmptyFilterSelect(filterSelect) {
    return filterSelect == null || JSON.stringify(filterSelect) === "{}";
}

function parseIndex(value) {
    const n = parseFloat(value);
    return Number.isNaN(n) ? 0 : n;
}

function shiftHiddenRows(rowhidden, index, count, direction) {
    const next = {};
    if (rowhidden == null) {
        return next;
    }
    for (const key in rowhidden) {
        const row = parseIndex(key);
        if (row < index) {
            next[row] = 0;
        } else if (row === index) {
            if (direction === "lefttop") {
                next[row + count] = 0;
            } else {
                next[row] = 0;
            }
        } else {
            next[row + count] = 0;
        }
    }
    return next;
}

function deleteHiddenRows(rowhidden, start, end) {
    const next = {};
    const removed = end - start + 1;
    if (rowhidden == null) {
        return next;
    }
    for (const key in rowhidden) {
        const row = parseIndex(key);
        if (row < start) {
            next[row] = 0;
        } else if (row > end) {
            next[row - removed] = 0;
        }
    }
    return next;
}

function resolveColumnIndex(item, key, originStart) {
    if (item != null && item.cindex != null && item.cindex !== "") {
        return parseIndex(item.cindex);
    }
    return originStart + parseIndex(key);
}

function finishColumnItem(item, cindex, startCol, endCol, startRow, endRow) {
    item.cindex = cindex;
    item.stc = startCol;
    item.edc = endCol;
    item.str = startRow;
    item.edr = endRow;
    return item;
}

/**
 * 插入行/列后同步筛选。
 * @param {"row"|"column"} type
 * @param {"lefttop"|"rightbottom"} direction
 */
export function syncFilterOnInsert(filterSelect, filter, type, index, value, direction) {
    if (isEmptyFilterSelect(filterSelect)) {
        return null;
    }

    const count = Math.floor(value);
    let startRow = filterSelect.row[0];
    let endRow = filterSelect.row[1];
    let startCol = filterSelect.column[0];
    let endCol = filterSelect.column[1];
    const originStartCol = startCol;
    let nextFilter = null;

    if (type === "row") {
        if (startRow < index) {
            if (endRow === index && direction === "lefttop") {
                endRow += count;
            } else if (endRow > index) {
                endRow += count;
            }
        } else if (startRow === index) {
            if (direction === "lefttop") {
                startRow += count;
                endRow += count;
            } else if (direction === "rightbottom" && endRow > index) {
                endRow += count;
            }
        } else {
            startRow += count;
            endRow += count;
        }

        if (filter != null) {
            nextFilter = {};
            for (const key in filter) {
                const item = cloneJson(filter[key]);
                item.rowhidden = shiftHiddenRows(item.rowhidden, index, count, direction);
                item.str = startRow;
                item.edr = endRow;
                nextFilter[key] = item;
            }
        }
    } else if (type === "column") {
        if (startCol < index) {
            if (endCol === index && direction === "lefttop") {
                endCol += count;
            } else if (endCol > index) {
                endCol += count;
            }
        } else if (startCol === index) {
            if (direction === "lefttop") {
                startCol += count;
                endCol += count;
            } else if (direction === "rightbottom" && endCol > index) {
                endCol += count;
            }
        } else {
            startCol += count;
            endCol += count;
        }

        if (filter != null) {
            nextFilter = {};
            for (const key in filter) {
                const item = cloneJson(filter[key]);
                let columnIndex = resolveColumnIndex(item, key, originStartCol);
                if (columnIndex === index && direction === "lefttop") {
                    columnIndex += count;
                } else if (columnIndex > index) {
                    columnIndex += count;
                }
                if (columnIndex < startCol || columnIndex > endCol) {
                    continue;
                }
                nextFilter[columnIndex - startCol] = finishColumnItem(
                    item,
                    columnIndex,
                    startCol,
                    endCol,
                    startRow,
                    endRow
                );
            }
        }
    }

    return {
        filter_select: {
            row: [startRow, endRow],
            column: [startCol, endCol],
        },
        filter: nextFilter,
    };
}

/**
 * 删除行/列后同步筛选。整段筛选被删则清空。
 */
export function syncFilterOnDelete(filterSelect, filter, type, start, end) {
    if (isEmptyFilterSelect(filterSelect)) {
        return null;
    }

    const removed = end - start + 1;
    let startRow = filterSelect.row[0];
    let endRow = filterSelect.row[1];
    let startCol = filterSelect.column[0];
    let endCol = filterSelect.column[1];
    const originStartCol = startCol;
    const result = {
        filter_select: null,
        filter: null,
    };

    if (type === "row") {
        if (startRow > end) {
            startRow -= removed;
            endRow -= removed;
            result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
        } else if (startRow < start) {
            if (endRow < start) {
                result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
            } else if (endRow <= end) {
                endRow = start - 1;
                result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
            } else {
                endRow -= removed;
                result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
            }
        } else if (endRow > end) {
            startRow = start;
            endRow -= removed;
            result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
        }

        if (result.filter_select != null && filter != null) {
            result.filter = {};
            for (const key in filter) {
                const item = cloneJson(filter[key]);
                item.rowhidden = deleteHiddenRows(item.rowhidden, start, end);
                item.str = result.filter_select.row[0];
                item.edr = result.filter_select.row[1];
                result.filter[key] = item;
            }
        }
    } else if (type === "column") {
        if (startCol > end) {
            startCol -= removed;
            endCol -= removed;
            result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
        } else if (startCol < start) {
            if (endCol < start) {
                result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
            } else if (endCol <= end) {
                endCol = start - 1;
                result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
            } else {
                endCol -= removed;
                result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
            }
        } else if (endCol > end) {
            startCol = start;
            endCol -= removed;
            result.filter_select = { row: [startRow, endRow], column: [startCol, endCol] };
        }

        if (result.filter_select != null && filter != null) {
            result.filter = {};
            const nextStart = result.filter_select.column[0];
            const nextEnd = result.filter_select.column[1];
            for (const key in filter) {
                const item = cloneJson(filter[key]);
                let columnIndex = resolveColumnIndex(item, key, originStartCol);
                if (columnIndex >= start && columnIndex <= end) {
                    continue;
                }
                if (columnIndex > end) {
                    columnIndex -= removed;
                }
                if (columnIndex < nextStart || columnIndex > nextEnd) {
                    continue;
                }
                result.filter[columnIndex - nextStart] = finishColumnItem(
                    item,
                    columnIndex,
                    nextStart,
                    nextEnd,
                    startRow,
                    endRow
                );
            }
        }
    }

    return result;
}

/**
 * 将 [fromStart, fromEnd] 整段移动到 toIndex（目标为移动前坐标系）。
 * 对齐 Univer moveMatrixArray / handleMoveColsCommand。
 */
export function remapMovedIndex(index, fromStart, fromEnd, toIndex) {
    const count = fromEnd - fromStart + 1;
    if (index >= fromStart && index <= fromEnd) {
        const offset = index - fromStart;
        if (toIndex <= fromStart) {
            return toIndex + offset;
        }
        return toIndex - count + offset;
    }
    if (toIndex <= fromStart) {
        if (index >= toIndex && index < fromStart) {
            return index + count;
        }
    } else if (toIndex > fromEnd) {
        if (index > fromEnd && index < toIndex) {
            return index - count;
        }
    }
    return index;
}

export function syncFilterOnMoveColumns(filterSelect, filter, fromStart, fromEnd, toIndex) {
    if (isEmptyFilterSelect(filterSelect)) {
        return null;
    }

    const startCol = filterSelect.column[0];
    const endCol = filterSelect.column[1];
    const outsideLeft = fromEnd < startCol && toIndex <= startCol;
    const outsideRight = fromStart > endCol && toIndex > endCol;
    if (outsideLeft || outsideRight) {
        return {
            filter_select: cloneJson(filterSelect),
            filter: cloneJson(filter),
        };
    }

    const columns = {};
    for (let col = startCol; col <= endCol; col++) {
        columns[col] = col;
    }

    const moving = [];
    const count = fromEnd - fromStart + 1;
    for (let i = 0; i < count; i++) {
        moving.push(columns[fromStart + i]);
        delete columns[fromStart + i];
    }

    if (toIndex < fromStart) {
        for (let col = fromStart - 1; col >= toIndex; col--) {
            if (col in columns) {
                columns[col + count] = columns[col];
                delete columns[col];
            }
        }
    } else if (toIndex > fromEnd) {
        for (let col = fromStart + count; col < toIndex; col++) {
            if (col in columns) {
                columns[col - count] = columns[col];
                delete columns[col];
            }
        }
    }

    const destStart = toIndex <= fromStart ? toIndex : toIndex - count;
    for (let i = 0; i < moving.length; i++) {
        columns[destStart + i] = moving[i];
    }

    let startBorder = startCol;
    let endBorder = endCol;
    if (
        startCol >= fromStart &&
        startCol <= fromEnd &&
        toIndex > fromStart &&
        fromEnd < endCol
    ) {
        startBorder = fromEnd + 1;
    }
    if (
        endCol >= fromStart &&
        endCol <= fromEnd &&
        toIndex < fromStart &&
        fromStart > startCol
    ) {
        endBorder = fromStart - 1;
    }

    const placed = Object.keys(columns).map(Number);
    const oldToNew = {};
    placed.forEach((nextCol) => {
        oldToNew[columns[nextCol]] = nextCol;
    });

    const nextStart = oldToNew[startBorder];
    const nextEnd = oldToNew[endBorder];
    if (nextStart == null || nextEnd == null) {
        return {
            filter_select: null,
            filter: null,
        };
    }

    let nextFilter = null;
    if (filter != null) {
        nextFilter = {};
        for (const key in filter) {
            const item = cloneJson(filter[key]);
            const oldCol = resolveColumnIndex(item, key, startCol);
            const newCol = oldToNew[oldCol];
            if (newCol == null || newCol < nextStart || newCol > nextEnd) {
                continue;
            }
            nextFilter[newCol - nextStart] = finishColumnItem(
                item,
                newCol,
                nextStart,
                nextEnd,
                filterSelect.row[0],
                filterSelect.row[1]
            );
        }
    }

    return {
        filter_select: {
            row: [filterSelect.row[0], filterSelect.row[1]],
            column: [nextStart, nextEnd],
        },
        filter: nextFilter,
    };
}
