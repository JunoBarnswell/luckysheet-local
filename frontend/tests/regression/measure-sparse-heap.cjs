const fs = require("fs");
const path = require("path");
const Module = require("module");

function loadEsModule(absPath) {
    const code = fs.readFileSync(absPath, "utf8");
    const exported = [];
    const cjs =
        code.replace(/export function (\w+)/g, function (_all, name) {
            exported.push(name);
            return "function " + name;
        }) +
        "\nmodule.exports = { " +
        exported.join(", ") +
        " };\n";

    const m = new Module(absPath);
    m.filename = absPath;
    m.paths = Module._nodeModulePaths(path.dirname(absPath));
    m._compile(cjs, absPath);
    return m.exports;
}

function mb(n) {
    return Number((n / 1048576).toFixed(2));
}

const sparse = loadEsModule(path.resolve(__dirname, "../../src/global/sparseGrid.js"));

const before = process.memoryUsage();
const t0 = process.hrtime.bigint();
const celldata = new Array(100000);
for (let i = 0; i < 100000; i++) {
    celldata[i] = { r: i, c: i % 20, v: { v: i, m: String(i) } };
}
const grid = sparse.createSparseGridFromCelldata(celldata, 100000, 20);
const t1 = process.hrtime.bigint();
const after = process.memoryUsage();
const exported = sparse.sparseGridToCelldata(grid);

const result = {
    env: "node " + process.version,
    platform: process.platform + " " + process.arch,
    logicalRC: [grid.length, grid[0] && grid[0].length],
    occupied: sparse.occupiedCellCount(grid),
    celldataN: exported.length,
    storeRows: grid.__luckysheetSparseGrid && grid.__luckysheetSparseGrid.store.size,
    create_ms: Number((Number(t1 - t0) / 1e6).toFixed(1)),
    heapUsed_delta_mb: mb(after.heapUsed - before.heapUsed),
    rss_delta_mb: mb(after.rss - before.rss),
    heapUsed_after_mb: mb(after.heapUsed),
    rss_after_mb: mb(after.rss),
    note: "Node process.memoryUsage only. Not browser performance.memory / FPS.",
};

console.log(JSON.stringify(result, null, 2));
if (result.occupied !== 100000 || result.celldataN !== 100000) {
    process.exit(1);
}
