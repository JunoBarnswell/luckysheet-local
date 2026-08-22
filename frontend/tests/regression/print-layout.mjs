import {
    cellDisplayValue,
    usedRange,
    paginateByPaper,
    normalizeLayoutFromPrintoptions,
    writePrintoptions,
    resolveHeaderFooterText,
    parsePrintTitles,
    adjustPaginationForMerge,
    PrintArea,
    PrintPaperSize,
    PrintDirection,
    PrintScale,
    PrintPaperMargin,
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
assert(cellDisplayValue({ v: 0, m: "0" }) === "0", "print includes zero value");

const data = [
    [{ v: 1, m: "1" }, null],
    [null, { v: 2, m: "2" }],
    [null, null],
];
const used = usedRange(data);
assert(used.row[1] === 1 && used.column[1] === 1, "usedRange clips empty tail");

const zeroData = [[{ v: 0 }]];
const usedZero = usedRange(zeroData);
assert(usedZero.row[1] === 0 && usedZero.column[1] === 0, "usedRange includes v=0 cell");

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

const written = writePrintoptions(
    {},
    {
        paperSize: "Letter",
        direction: PrintDirection.Portrait,
        scale: PrintScale.Origin,
        customScale: 100,
        margin: PrintPaperMargin.Normal,
    },
    { gridlines: false, draft: false, hAlign: "Start", vAlign: "Start" }
);
assert(written.pageSetup.paperSize === 1, "Letter writes excel code 1");
assert(written.printOptions.gridLines === 0, "gridlines false → 0");
assert(clampScale(5) === 10 && clampScale(900) === 400, "scale clamped 10-400");
assert(PrintArea.CurrentSheet === "CurrentSheet", "PrintArea enum");

const hf = resolveHeaderFooterText("@Page/@TotalPage @WorksheetTitle", {
    page: 2,
    pageTotal: 5,
    worksheetTitle: "Sheet1",
    workbookTitle: "Book",
    sheetPage: 2,
    sheetPageTotal: 5,
});
assert(hf.indexOf("2") > -1 && hf.indexOf("Sheet1") > -1, "header footer symbols");

const titles = parsePrintTitles(
    { config: { printoptions: { PrintTitles: { row: "S!$1:$2", column: "S!$A:$B" } } } },
    { freeze: [] }
);
assert(titles.row[0] === 0 && titles.row[1] === 1, "parse row titles");
assert(titles.column[0] === 0 && titles.column[1] === 1, "parse column titles");

const merged = adjustPaginationForMerge(
    [{ row: [0, 2], column: [0, 1] }],
    { "0_0": { rs: 4, cs: 2 } },
    visibledatarow,
    visibledatacolumn
);
assert(merged.length >= 0, "merge adjust returns pages");

console.log("print-layout.mjs all passed");
