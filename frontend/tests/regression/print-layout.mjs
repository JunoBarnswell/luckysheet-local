import {
    cellDisplayValue,
    usedRange,
    paginateByPaper,
    normalizeLayoutFromPrintoptions,
    writePrintoptions,
    PrintArea,
    PrintPaperSize,
    PrintDirection,
    PrintScale,
    clampScale,
} from "../../src/expendPlugins/print/printLayout.js";

function assert(cond, name) {
    if (!cond) {
        throw new Error("FAIL " + name);
    }
    console.log("PASS " + name);
}

assert(cellDisplayValue({ f: "=A1+1", v: 3, m: "3" }) === "3", "print uses m not f");
assert(cellDisplayValue({ f: "=A1+1", v: 8 }) === "8", "print falls back to v");
assert(cellDisplayValue({ f: "=A1" }) === "", "print never emits formula string");

const data = [
    [{ v: 1, m: "1" }, null],
    [null, { v: 2, m: "2" }],
    [null, null],
];
const used = usedRange(data);
assert(used.row[1] === 1 && used.column[1] === 1, "usedRange clips empty tail");

const visibledatarow = [20, 40, 60, 80, 100, 120];
const visibledatacolumn = [80, 160, 240, 320];
const pages = paginateByPaper(
    { row: [0, 5], column: [0, 3] },
    visibledatarow,
    visibledatacolumn,
    { innerW: 160, innerH: 40 },
    {}
);
assert(pages.length > 0, "paginate emits pages");
assert(pages[0].row[1] >= pages[0].row[0], "page rows are ordered");

const layout = normalizeLayoutFromPrintoptions({
    PrintArea: "$A$1:$C$4",
    pageSetup: { paperSize: 9, orientation: 1, scale: 120 },
    printOptions: { gridLines: 1 },
});
assert(layout.paperSize === PrintPaperSize.A4, "excel 9 → A4");
assert(layout.direction === PrintDirection.Landscape, "orientation 1 → landscape");
assert(layout.customScale === 120 || layout.scale === PrintScale.Custom, "scale maps");
assert(layout.rangeText === "$A$1:$C$4", "PrintArea kept");

const written = writePrintoptions({}, {
    paperSize: "Letter",
    direction: PrintDirection.Portrait,
    scale: PrintScale.Origin,
    customScale: 100,
    margin: "Normal",
}, { gridlines: false, draft: false, hAlign: "Start", vAlign: "Start" });
assert(written.pageSetup.paperSize === 1, "Letter writes excel code 1");
assert(written.printOptions.gridLines === 0, "gridlines false → 0");
assert(clampScale(5) === 10 && clampScale(900) === 400, "scale clamped 10-400");
assert(PrintArea.CurrentSheet === "CurrentSheet", "PrintArea enum");

console.log("print-layout.mjs all passed");
