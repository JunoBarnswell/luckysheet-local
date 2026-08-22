import {
    remapMovedIndex,
    syncFilterOnDelete,
    syncFilterOnInsert,
    syncFilterOnMoveColumns,
} from "../../src/controllers/filterSync.js";
import {
    getFilterNumericValue,
    getFilterValueInfo,
    isFilterNumericCell,
    shouldHideByCaljs,
} from "../../src/controllers/filterCondition.js";

function assert(cond, name) {
    if (!cond) {
        throw new Error("FAIL " + name);
    }
    console.log("PASS " + name);
}

const filterSelect = { row: [0, 10], column: [2, 5] };
const filter = {
    0: { cindex: 2, stc: 2, edc: 5, str: 0, edr: 10, caljs: { value: "morethan", value1: "3" } },
    2: { cindex: 4, stc: 2, edc: 5, str: 0, edr: 10, caljs: { value: "equal", value1: "1" } },
};

const inserted = syncFilterOnInsert(filterSelect, filter, "column", 3, 1, "lefttop");
assert(inserted.filter_select.column[0] === 2 && inserted.filter_select.column[1] === 6, "insert-inside expands filter columns");
assert(inserted.filter[0].cindex === 2, "insert-inside keeps left filter column");
assert(inserted.filter[3].cindex === 5, "insert-inside shifts later filter column");
assert(inserted.filter[3].stc === 2 && inserted.filter[3].edc === 6, "insert-inside updates stc/edc");

const insertLeft = syncFilterOnInsert(filterSelect, filter, "column", 1, 2, "lefttop");
assert(insertLeft.filter_select.column[0] === 4 && insertLeft.filter_select.column[1] === 7, "insert-left shifts whole range");
assert(insertLeft.filter[0].cindex === 4, "insert-left remaps first filter key");
assert(insertLeft.filter[2].cindex === 6, "insert-left remaps later filter key relative to new start");

const deleteRight = syncFilterOnDelete(filterSelect, filter, "column", 8, 9);
assert(deleteRight.filter_select.column[0] === 2 && deleteRight.filter_select.column[1] === 5, "delete-right keeps range");
assert(deleteRight.filter[2].cindex === 4, "delete-right keeps filter indices");

const deleteInside = syncFilterOnDelete(filterSelect, filter, "column", 4, 4);
assert(deleteInside.filter_select.column[1] === 4, "delete-inside shrinks end");
assert(deleteInside.filter[0].cindex === 2, "delete-inside keeps earlier column");
assert(deleteInside.filter[2] == null, "delete-inside drops removed column criteria");

const deleteAll = syncFilterOnDelete(filterSelect, filter, "column", 2, 5);
assert(deleteAll.filter_select == null, "delete-all-range clears filter");

const moved = syncFilterOnMoveColumns(filterSelect, filter, 2, 2, 6);
assert(moved.filter_select.column[0] === 2 && moved.filter_select.column[1] === 4, "move-first-col shrinks range after header leaves");
assert(moved.filter[0] == null, "move-first-col drops criteria that left the range");
assert(moved.filter[1].cindex === 3, "move-first-col remaps remaining filter column");

assert(remapMovedIndex(4, 2, 2, 6) === 3, "remap later column after move-right");
assert(remapMovedIndex(2, 2, 2, 6) === 5, "remap moved column to destination");

const numberCell = { v: 10, m: "10.00", ct: { t: "n" } };
const numberNoCt = { v: 10, m: "10" };
const textTen = { v: "10", m: "10" };
assert(isFilterNumericCell(numberCell) && isFilterNumericCell(numberNoCt), "number cells are numeric without string compare");
assert(!isFilterNumericCell(textTen), "plain text 10 is not numeric");
assert(getFilterNumericValue(numberCell) === 10, "numeric value stays number");
assert(getFilterValueInfo(numberCell).key === "n#$$$#10", "numeric filter key is typed");
assert(getFilterValueInfo(textTen).key !== getFilterValueInfo(numberCell).key, "number 10 and text 10 do not share a key");

assert(
    !shouldHideByCaljs(numberNoCt, { value: "morethan", value1: "9" }),
    "greater-than uses numeric compare so 10 > 9"
);
assert(
    shouldHideByCaljs(textTen, { value: "morethan", value1: "9" }),
    "text 10 is hidden by numeric greater-than"
);
assert(
    !shouldHideByCaljs({ v: 5, ct: { t: "n" } }, { value: "include", value1: "1", value2: "9" }),
    "between keeps legacy AND"
);
assert(
    !shouldHideByCaljs({ v: 5, ct: { t: "n" } }, { value: "include", value1: "8", value2: "9", and: false }),
    "between OR keeps 5 because 5 <= 9"
);
assert(
    shouldHideByCaljs({ v: 5, ct: { t: "n" } }, { value: "noinclude", value1: "1", value2: "9" }),
    "not-between hides inside range when and omitted"
);
assert(
    !shouldHideByCaljs(
        { v: 12, ct: { t: "n" } },
        { value: "custom", operator1: "morethan", value1: "10", operator2: "lessthan", value2: "20", and: true }
    ),
    "custom AND keeps 12 when 12 > 10 and 12 < 20"
);
assert(
    shouldHideByCaljs(
        { v: 8, ct: { t: "n" } },
        { value: "custom", operator1: "morethan", value1: "10", operator2: "lessthan", value2: "5", and: true }
    ),
    "custom AND hides when only one side matches"
);
assert(
    !shouldHideByCaljs(
        { v: 8, ct: { t: "n" } },
        { value: "custom", operator1: "morethan", value1: "10", operator2: "lessthan", value2: "9", and: false }
    ),
    "custom OR keeps when second condition matches"
);
assert(
    shouldHideByCaljs({ v: "abc", m: "abc" }, { value: "dateequal", value1: "2020-01-01" }, function () { return 1; }),
    "date condition hides non-date text instead of string-comparing it"
);
assert(
    !shouldHideByCaljs({ v: 1, ct: { t: "d" } }, { value: "dateequal", value1: "2020-01-01" }, function () { return 1; }),
    "dateequal still matches date cells"
);

console.log("filter-phase3 all passed");
