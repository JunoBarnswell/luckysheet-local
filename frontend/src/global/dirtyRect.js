function asRangePair(pair) {
    if (pair == null || pair.length < 2) {
        return null;
    }
    const a = pair[0];
    const b = pair[1];
    if (a == null || b == null) {
        return null;
    }
    return a <= b ? [a, b] : [b, a];
}

export function normalizeDirtyRect(dirtyRect) {
    if (dirtyRect == null) {
        return null;
    }
    const row = asRangePair(dirtyRect.row);
    const column = asRangePair(dirtyRect.column);
    if (row == null || column == null) {
        return null;
    }
    const next = {
        row: row,
        column: column
    };
    if (dirtyRect.fromScroll) {
        next.fromScroll = true;
    }
    if (dirtyRect.axisPatch === false) {
        next.axisPatch = false;
    }
    return next;
}

export function unionDirtyRects(a, b) {
    const left = normalizeDirtyRect(a);
    const right = normalizeDirtyRect(b);
    if (left == null) {
        return right;
    }
    if (right == null) {
        return left;
    }
    return {
        row: [Math.min(left.row[0], right.row[0]), Math.max(left.row[1], right.row[1])],
        column: [Math.min(left.column[0], right.column[0]), Math.max(left.column[1], right.column[1])]
    };
}

export function unionRangesToDirtyRect(range) {
    if (range == null) {
        return null;
    }
    const list = Array.isArray(range) ? range : [range];
    let dirty = null;
    for (let i = 0; i < list.length; i++) {
        const item = list[i];
        if (item == null || item.row == null || item.column == null) {
            continue;
        }
        dirty = unionDirtyRects(dirty, { row: item.row, column: item.column });
    }
    return dirty;
}

export function intersectDirtyRect(a, b) {
    const left = normalizeDirtyRect(a);
    const right = normalizeDirtyRect(b);
    if (left == null || right == null) {
        return null;
    }
    const r1 = Math.max(left.row[0], right.row[0]);
    const r2 = Math.min(left.row[1], right.row[1]);
    const c1 = Math.max(left.column[0], right.column[0]);
    const c2 = Math.min(left.column[1], right.column[1]);
    if (r1 > r2 || c1 > c2) {
        return null;
    }
    return { row: [r1, r2], column: [c1, c2] };
}

function rangesIntersect(a1, a2, b1, b2) {
    return a1 <= b2 && b1 <= a2;
}

export function expandDirtyRectForMerges(dirtyRect, mergeConfig) {
    const dirty = normalizeDirtyRect(dirtyRect);
    if (dirty == null || mergeConfig == null) {
        return dirty;
    }

    let r1 = dirty.row[0];
    let r2 = dirty.row[1];
    let c1 = dirty.column[0];
    let c2 = dirty.column[1];
    let expanded = true;

    while (expanded) {
        expanded = false;
        for (const key in mergeConfig) {
            if (!Object.prototype.hasOwnProperty.call(mergeConfig, key)) {
                continue;
            }
            const mc = mergeConfig[key];
            if (mc == null || mc.rs == null || mc.cs == null) {
                continue;
            }
            const mr1 = mc.r;
            const mr2 = mc.r + mc.rs - 1;
            const mc1 = mc.c;
            const mc2 = mc.c + mc.cs - 1;
            if (!rangesIntersect(r1, r2, mr1, mr2) || !rangesIntersect(c1, c2, mc1, mc2)) {
                continue;
            }
            if (mr1 < r1) {
                r1 = mr1;
                expanded = true;
            }
            if (mr2 > r2) {
                r2 = mr2;
                expanded = true;
            }
            if (mc1 < c1) {
                c1 = mc1;
                expanded = true;
            }
            if (mc2 > c2) {
                c2 = mc2;
                expanded = true;
            }
        }
    }

    const next = { row: [r1, r2], column: [c1, c2] };
    if (dirty.fromScroll) {
        next.fromScroll = true;
        next.axisPatch = false;
    }
    if (dirty.axisPatch === false) {
        next.axisPatch = false;
    }
    return next;
}

export function visibleRectFromSearch(visibledatarow, visibledatacolumn, scrollWidth, scrollHeight, drawWidth, drawHeight, searcharray) {
    if (visibledatarow == null || visibledatacolumn == null || searcharray == null) {
        return null;
    }
    let row_st = searcharray(visibledatarow, scrollHeight);
    let row_ed = searcharray(visibledatarow, scrollHeight + drawHeight);
    let col_st = searcharray(visibledatacolumn, scrollWidth);
    let col_ed = searcharray(visibledatacolumn, scrollWidth + drawWidth);
    if (row_st < 0) {
        row_st = 0;
    }
    if (row_ed < 0) {
        row_ed = visibledatarow.length - 1;
    }
    if (col_st < 0) {
        col_st = 0;
    }
    if (col_ed < 0) {
        col_ed = visibledatacolumn.length - 1;
    }
    if (row_ed < row_st || col_ed < col_st) {
        return null;
    }
    return { row: [row_st, row_ed], column: [col_st, col_ed] };
}

export function inferScrollDirtyRect(prev, next) {
    const left = normalizeDirtyRect(prev);
    const right = normalizeDirtyRect(next);
    if (left == null || right == null) {
        return null;
    }

    const viewRows = right.row[1] - right.row[0] + 1;
    const viewCols = right.column[1] - right.column[0] + 1;
    if (Math.abs(right.row[0] - left.row[0]) > viewRows || Math.abs(right.column[0] - left.column[0]) > viewCols) {
        return null;
    }

    let dirty = null;
    if (right.row[0] > left.row[0]) {
        dirty = unionDirtyRects(dirty, {
            row: [left.row[1] + 1, right.row[1]],
            column: right.column
        });
    }
    else if (right.row[0] < left.row[0]) {
        dirty = unionDirtyRects(dirty, {
            row: [right.row[0], left.row[0] - 1],
            column: right.column
        });
    }

    if (right.column[0] > left.column[0]) {
        dirty = unionDirtyRects(dirty, {
            row: right.row,
            column: [left.column[1] + 1, right.column[1]]
        });
    }
    else if (right.column[0] < left.column[0]) {
        dirty = unionDirtyRects(dirty, {
            row: right.row,
            column: [right.column[0], left.column[0] - 1]
        });
    }

    if (dirty == null) {
        return null;
    }
    dirty.fromScroll = true;
    dirty.axisPatch = false;
    return dirty;
}
