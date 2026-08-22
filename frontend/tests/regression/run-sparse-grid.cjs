const fs = require("fs");
const path = require("path");
const Module = require("module");

function loadEsModule(absPath) {
    const code = fs.readFileSync(absPath, "utf8");
    const exported = [];
    const cjs = code.replace(/export function (\w+)/g, function (_all, name) {
        exported.push(name);
        return "function " + name;
    }) + "\nmodule.exports = { " + exported.join(", ") + " };\n";

    const m = new Module(absPath);
    m.filename = absPath;
    m.paths = Module._nodeModulePaths(path.dirname(absPath));
    m._compile(cjs, absPath);
    return m.exports;
}

const srcDir = path.resolve(__dirname, "../../src/global");
const sparse = loadEsModule(path.join(srcDir, "sparseGrid.js"));
const dirty = loadEsModule(path.join(srcDir, "dirtyRect.js"));

const {
    asSparseGrid,
    cloneSheetData,
    createSparseGridFromCelldata,
    ensureSparseSize,
    materializeGridData,
    occupiedCellCount,
    snapshotSheetFile,
    sparseGridToCelldata,
} = sparse;
const {
    expandDirtyRectForMerges,
    inferScrollDirtyRect,
    intersectDirtyRect,
    unionRangesToDirtyRect,
} = dirty;

function assert(cond, name, detail) {
    if (!cond) {
        throw new Error("FAIL " + name + (detail ? " — " + detail : ""));
    }
    console.log("PASS  " + name + (detail ? " — " + detail : ""));
}

function assertEqual(actual, expected, name) {
    assert(actual === expected, name, "actual=" + JSON.stringify(actual) + " expected=" + JSON.stringify(expected));
}

const grid = createSparseGridFromCelldata(
    [
        { r: 0, c: 0, v: { v: 1, m: "1" } },
        { r: 0, c: 1, v: { f: "=A1*2", v: 2, m: "2" } },
        { r: 99999, c: 3, v: { v: "tail" } },
    ],
    100000,
    20
);

assertEqual(grid.length, 100000, "logical row count keeps file.row");
assertEqual(grid[0].length, 20, "logical column count keeps file.column");
assertEqual(occupiedCellCount(grid), 3, "only occupied cells are stored");
assert(grid[0][0] != null && grid[0][0].v === 1, "data[r][c] read occupied");
assert(grid[1][1] == null, "empty cell is null without allocating a row array");
assert(grid[50000][0] == null, "mid-sheet empty row stays sparse");
assert(grid[99999][3] != null && grid[99999][3].v === "tail", "far cell is reachable");

const backing = grid.__luckysheetSparseGrid;
assert(backing != null && backing.store.size === 2, "store only keeps occupied rows");
assert(!(grid instanceof Array) || backing.store.size < 100, "grid target is not a preallocated Array(100000)");

const exported = sparseGridToCelldata(grid);
assertEqual(exported.length, 3, "toCelldata exports sparse {r,c,v}");
assert(exported[0].r === 0 && exported[0].c === 0, "first exported cell is A1");
assert(exported[2].r === 99999 && exported[2].c === 3, "last exported cell keeps coordinates");

const materialized = materializeGridData(grid);
assert(Array.isArray(materialized) && Array.isArray(materialized[0]), "public compatibility data is a real two-dimensional array");
assertEqual(materialized[0][0].v, 1, "materialized grid preserves cell values");
materialized[0][0].v = 88;
assertEqual(grid[0][0].v, 1, "materialized grid does not leak internal cells");
const fileSnapshot = snapshotSheetFile({ name: "S1", data: grid });
assert(Array.isArray(fileSnapshot.data) && fileSnapshot.data[99999][3].v === "tail", "sheet snapshot materializes sparse data");

const clone = cloneSheetData(grid);
clone[0][0] = { v: 99 };
assert(grid[0][0].v === 1, "cloneSheetData is isolated");
assertEqual(occupiedCellCount(clone), 3, "clone does not densify");

grid[2][5] = { v: "grow" };
assertEqual(grid.length, 100000, "in-range write does not grow logical rows");
assert(grid[2][5].v === "grow", "write uses [][] facade");
assertEqual(occupiedCellCount(grid), 4, "write adds one occupied cell");

ensureSparseSize(grid, 100010, 25);
assertEqual(grid.length, 100011, "ensureSparseSize grows logical rows only");
assertEqual(grid[0].length, 26, "ensureSparseSize grows logical columns only");
assertEqual(occupiedCellCount(grid), 4, "grow does not allocate empty cells");

const fromDense = asSparseGrid(
    [
        [{ v: "a" }, null, null],
        [null, null, { v: "b" }],
    ],
    80,
    10
);
assertEqual(fromDense.length, 80, "dense wrap keeps requested rows");
assertEqual(occupiedCellCount(fromDense), 2, "dense wrap drops nulls");

const merged = expandDirtyRectForMerges(
    { row: [1, 1], column: [1, 1] },
    { "0_0": { r: 0, c: 0, rs: 3, cs: 3 } }
);
assertEqual(merged.row[0], 0, "merge expands dirty row start");
assertEqual(merged.row[1], 2, "merge expands dirty row end");
assertEqual(merged.column[0], 0, "merge expands dirty column start");
assertEqual(merged.column[1], 2, "merge expands dirty column end");

const fromRange = unionRangesToDirtyRect([
    { row: [2, 4], column: [1, 1] },
    { row: [0, 1], column: [3, 5] },
]);
assertEqual(fromRange.row[0], 0, "union range row start");
assertEqual(fromRange.column[1], 5, "union range column end");

const hit = intersectDirtyRect({ row: [0, 10], column: [0, 10] }, { row: [8, 20], column: [9, 12] });
assertEqual(hit.row[0], 8, "dirty ∩ visible row start");
assertEqual(hit.column[1], 10, "dirty ∩ visible column end");

const scrollDirty = inferScrollDirtyRect(
    { row: [0, 20], column: [0, 10] },
    { row: [3, 23], column: [0, 10] }
);
assert(scrollDirty != null && scrollDirty.fromScroll === true, "small scroll yields incremental dirty");
assertEqual(scrollDirty.row[0], 21, "scroll dirty starts at newly exposed rows");
assert(inferScrollDirtyRect({ row: [0, 20], column: [0, 10] }, { row: [80, 100], column: [0, 10] }) == null, "large jump falls back to full redraw");

const sliced = grid.slice(0, 3);
assertEqual(sliced.length, 3, "slice keeps sparse row window");
assertEqual(occupiedCellCount(sliced), 3, "slice does not densify");

const snapshot = JSON.parse(JSON.stringify({
    row: grid.length,
    column: grid[0].length,
    celldata: sparseGridToCelldata(grid),
}));
assert(snapshot.celldata.every(function (item) {
    return item.r != null && item.c != null && item.v != null;
}), "JSON contract stays celldata sparse");

console.log("sparse-grid spec passed");
