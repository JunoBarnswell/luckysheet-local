const ExcelJS = require('exceljs');
const { cellValue, applyCellStyle, writeSheet, toArgb, fontName } = require('../controllers/luckyToXlsx');

function assert(cond, msg) {
    if (!cond) {
        throw new Error(msg);
    }
}

assert(toArgb('#fff000') === 'FFFFF000', 'hex color');
assert(toArgb('rgb(255, 0, 0)') === 'FFFF0000', 'rgb color');
assert(fontName(1) === 'Arial', 'font index');
assert(fontName('微软雅黑') === '微软雅黑', 'font name');

const formula = cellValue({ f: '=SUM(A1:A2)', v: 3, m: '3' });
assert(formula.formula === 'SUM(A1:A2)' && formula.result === 3, 'formula with result');
assert(cellValue({ v: 12 }) === 12, 'plain value');

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Demo');
writeSheet(sheet, {
    data: [
        [
            { v: 1, m: '1', bg: '#fff000', bl: 1, fs: 14, ht: 0, vt: 0 },
            { f: '=A1*2', v: 2, m: '2', ct: { fa: '0.00', t: 'n' } },
        ],
        [{ v: 'x', fc: '#ff0000', it: 1 }, { v: 'y' }],
    ],
    config: {
        merge: { '1_0': { r: 1, c: 0, rs: 1, cs: 2 } },
        rowlen: { 0: 32 },
        columnlen: { 0: 120 },
        borderInfo: [
            {
                rangeType: 'range',
                borderType: 'border-all',
                color: '#000000',
                style: '1',
                range: [{ row: [0, 1], column: [0, 1] }],
            },
        ],
    },
});

const a1 = sheet.getCell(1, 1);
assert(a1.font && a1.font.bold === true, 'bold');
assert(a1.fill && a1.fill.fgColor.argb === 'FFFFF000', 'fill');
assert(a1.border && a1.border.top, 'border');
const b1 = sheet.getCell(1, 2);
assert(b1.value && b1.value.formula === 'A1*2', 'exported formula, got ' + JSON.stringify(b1.value));
const merges = (sheet.model && sheet.model.merges) || [];
assert(merges.length >= 1 || (sheet._merges && Object.keys(sheet._merges).length >= 1), 'merge written');

console.log('luckyToXlsx verify ok');
