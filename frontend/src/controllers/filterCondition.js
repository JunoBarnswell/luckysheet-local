/**
 * 筛选条件求值：数值按数值比较（Univer #5862），CUSTOM 双条件支持 AND/OR（ICustomFilters.and）。
 * 日期三条件（dateequal / datelessthan / datemorethan）仍由调用方传入序列化函数。
 */

function isBlankCell(cell) {
    if (cell == null) {
        return true;
    }
    const value = cell.v;
    if (value == null) {
        return true;
    }
    if (typeof value === "string" && value.replace(/\s/g, "") === "") {
        return true;
    }
    return false;
}

function cellText(cell) {
    if (cell == null) {
        return "";
    }
    if (cell.m != null && cell.m !== "") {
        return String(cell.m);
    }
    if (cell.v == null) {
        return "";
    }
    return String(cell.v);
}

export function isFilterNumericCell(cell) {
    if (isBlankCell(cell)) {
        return false;
    }
    if (typeof cell.v === "number" && !Number.isNaN(cell.v)) {
        return true;
    }
    if (cell.ct != null && cell.ct.t === "n") {
        const n = Number(cell.v);
        return !Number.isNaN(n) && Number.isFinite(n);
    }
    return false;
}

export function getFilterNumericValue(cell) {
    if (!isFilterNumericCell(cell)) {
        return null;
    }
    return typeof cell.v === "number" ? cell.v : Number(cell.v);
}

export function parseFilterCompareNumber(raw) {
    if (typeof raw === "number" && !Number.isNaN(raw)) {
        return raw;
    }
    if (raw == null || raw === "") {
        return NaN;
    }
    return parseFloat(raw);
}

/**
 * 按值筛选的稳定键：数值用 n#$$$#<number>，避免 10 与 "10" 被当成同一字符串。
 */
export function getFilterValueInfo(cell) {
    if (isBlankCell(cell)) {
        return {
            key: "null#$$$#null",
            text: null,
            isNumber: false,
            isDate: false,
            sortValue: null,
        };
    }
    if (cell.ct != null && cell.ct.t === "d") {
        return {
            key: "d#$$$#" + cell.v,
            text: cell.m != null ? String(cell.m) : String(cell.v),
            isNumber: false,
            isDate: true,
            sortValue: cell.v,
        };
    }
    if (isFilterNumericCell(cell)) {
        const n = getFilterNumericValue(cell);
        return {
            key: "n#$$$#" + n,
            text: cell.m != null && cell.m !== "" ? String(cell.m) : String(n),
            isNumber: true,
            isDate: false,
            sortValue: n,
        };
    }
    const display = cell.m != null ? cell.m : cell.v;
    return {
        key: String(cell.v) + "#$$$#" + display,
        text: display,
        isNumber: false,
        isDate: false,
        sortValue: display,
    };
}

export function sortFilterValueKeys(infos) {
    return infos.slice().sort(function (a, b) {
        if (a.isNumber && b.isNumber) {
            return a.sortValue - b.sortValue;
        }
        if (a.isNumber) {
            return -1;
        }
        if (b.isNumber) {
            return 1;
        }
        const left = a.text == null ? "" : String(a.text);
        const right = b.text == null ? "" : String(b.text);
        return left.localeCompare(right, "zh");
    });
}

export function readCaljsAnd(caljs) {
    if (caljs == null) {
        return true;
    }
    if (caljs.and === false || caljs.and === "false" || caljs.and === 0) {
        return false;
    }
    if (caljs.and === true || caljs.and === "true" || caljs.and === 1) {
        return true;
    }
    return caljs.value === "noinclude" ? false : true;
}

function matchNumericOp(cell, op, rawVal) {
    const actual = getFilterNumericValue(cell);
    if (actual == null) {
        return false;
    }
    const expected = parseFilterCompareNumber(rawVal);
    if (Number.isNaN(expected)) {
        return false;
    }
    if (op === "morethan") {
        return actual > expected;
    }
    if (op === "moreequalthan") {
        return actual >= expected;
    }
    if (op === "lessthan") {
        return actual < expected;
    }
    if (op === "lessequalthan") {
        return actual <= expected;
    }
    if (op === "equal") {
        return actual === expected;
    }
    if (op === "noequal") {
        return actual !== expected;
    }
    return false;
}

export function cellMatchesFilterOp(cell, op, rawVal, dateSerialFn) {
    if (op == null || op === "" || op === "null" || op === "none") {
        return true;
    }

    if (op === "cellnull") {
        return isBlankCell(cell);
    }
    if (op === "cellnonull") {
        return !isBlankCell(cell);
    }

    if (op === "textinclude") {
        if (isBlankCell(cell)) {
            return false;
        }
        return cellText(cell).indexOf(rawVal == null ? "" : String(rawVal)) > -1;
    }
    if (op === "textnotinclude") {
        if (isBlankCell(cell)) {
            return true;
        }
        return cellText(cell).indexOf(rawVal == null ? "" : String(rawVal)) === -1;
    }
    if (op === "textstart") {
        if (isBlankCell(cell)) {
            return false;
        }
        const needle = rawVal == null ? "" : String(rawVal);
        return cellText(cell).substr(0, needle.length) === needle;
    }
    if (op === "textend") {
        if (isBlankCell(cell)) {
            return false;
        }
        const text = cellText(cell);
        const needle = rawVal == null ? "" : String(rawVal);
        if (needle.length > text.length) {
            return false;
        }
        return text.substr(text.length - needle.length, needle.length) === needle;
    }
    if (op === "textequal") {
        if (isBlankCell(cell)) {
            return false;
        }
        return cellText(cell) === (rawVal == null ? "" : String(rawVal));
    }

    if (op === "dateequal" || op === "datelessthan" || op === "datemorethan") {
        if (isBlankCell(cell) || cell.ct == null || cell.ct.t !== "d") {
            return false;
        }
        const expected = dateSerialFn ? dateSerialFn(rawVal) : parseFilterCompareNumber(rawVal);
        const actual = parseInt(cell.v, 10);
        if (op === "dateequal") {
            return actual === expected;
        }
        if (op === "datelessthan") {
            return actual < expected;
        }
        return actual > expected;
    }

    if (
        op === "morethan" ||
        op === "moreequalthan" ||
        op === "lessthan" ||
        op === "lessequalthan" ||
        op === "equal" ||
        op === "noequal"
    ) {
        return matchNumericOp(cell, op, rawVal);
    }

    return false;
}

function hideByBetween(cell, caljs) {
    const actual = getFilterNumericValue(cell);
    if (actual == null) {
        return true;
    }
    const value1 = parseFilterCompareNumber(caljs.value1);
    const value2 = parseFilterCompareNumber(caljs.value2);
    const min = value1 < value2 ? value1 : value2;
    const max = value1 < value2 ? value2 : value1;
    const geMin = actual >= min;
    const leMax = actual <= max;

    if (caljs.value === "include") {
        const pass = readCaljsAnd(caljs) ? geMin && leMax : geMin || leMax;
        return !pass;
    }

    if (caljs.and == null) {
        return geMin && leMax;
    }
    const outsideLow = actual < min;
    const outsideHigh = actual > max;
    const pass = readCaljsAnd(caljs) ? outsideLow && outsideHigh : outsideLow || outsideHigh;
    return !pass;
}

export function shouldHideByCaljs(cell, caljs, dateSerialFn) {
    if (caljs == null || caljs.value == null || caljs.value === "null") {
        return false;
    }

    const value = caljs.value;

    if (value === "custom") {
        const first = cellMatchesFilterOp(cell, caljs.operator1, caljs.value1, dateSerialFn);
        const second = cellMatchesFilterOp(cell, caljs.operator2, caljs.value2, dateSerialFn);
        const pass = readCaljsAnd(caljs) ? first && second : first || second;
        return !pass;
    }

    if (value === "include" || value === "noinclude") {
        if (caljs.operator1 && caljs.operator2) {
            const first = cellMatchesFilterOp(cell, caljs.operator1, caljs.value1, dateSerialFn);
            const second = cellMatchesFilterOp(cell, caljs.operator2, caljs.value2, dateSerialFn);
            const pass = readCaljsAnd(caljs) ? first && second : first || second;
            return !pass;
        }
        return hideByBetween(cell, caljs);
    }

    return !cellMatchesFilterOp(cell, value, caljs.value1, dateSerialFn);
}

export const FILTER_CUSTOM_OPERATORS = [
    "textinclude",
    "textnotinclude",
    "textstart",
    "textend",
    "textequal",
    "morethan",
    "moreequalthan",
    "lessthan",
    "lessequalthan",
    "equal",
    "noequal",
];
