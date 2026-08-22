const ExcelJS = require('exceljs');

const FONT_INDEX = [
    'Times New Roman',
    'Arial',
    'Tahoma',
    'Verdana',
    '微软雅黑',
    '宋体',
    '黑体',
    '楷体',
    '仿宋',
    '新宋体',
    '华文新魏',
    '华文行楷',
    '华文隶书',
];

const BORDER_STYLE = {
    1: 'thin',
    2: 'hair',
    3: 'dotted',
    4: 'dashed',
    5: 'dashDot',
    6: 'dashDotDot',
    7: 'double',
    8: 'medium',
    9: 'mediumDashed',
    10: 'mediumDashDot',
    11: 'mediumDashDotDot',
    12: 'slantDashDot',
    13: 'thick',
};

const HT = { 0: 'center', 1: 'left', 2: 'right' };
const VT = { 0: 'middle', 1: 'top', 2: 'bottom' };
const TR = { 0: 0, 1: 45, 2: -45, 3: 255, 4: 90, 5: -90 };

function toArgb(color) {
    if (color == null || typeof color !== 'string') {
        return null;
    }
    const raw = color.trim();
    if (!raw) {
        return null;
    }
    const rgb = raw.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
    if (rgb) {
        return (
            'FF' +
            [rgb[1], rgb[2], rgb[3]]
                .map((n) => Number(n).toString(16).padStart(2, '0'))
                .join('')
                .toUpperCase()
        );
    }
    let hex = raw.startsWith('#') ? raw.slice(1) : raw;
    if (hex.length === 3) {
        hex = hex
            .split('')
            .map((ch) => ch + ch)
            .join('');
    }
    if (hex.length === 6) {
        return 'FF' + hex.toUpperCase();
    }
    if (hex.length === 8) {
        return hex.toUpperCase();
    }
    return null;
}

function fontName(ff) {
    if (ff == null) {
        return 'Calibri';
    }
    if (typeof ff === 'number' || /^\d+$/.test(String(ff))) {
        return FONT_INDEX[Number(ff)] || 'Calibri';
    }
    return String(ff).replace(/["']/g, '') || 'Calibri';
}

function borderEdge(style, color) {
    const mapped = BORDER_STYLE[Number(style)] || BORDER_STYLE[String(style)] || 'thin';
    const argb = toArgb(color) || 'FF000000';
    return { style: mapped, color: { argb } };
}

function cellValue(cell) {
    if (cell == null) {
        return null;
    }
    if (typeof cell !== 'object') {
        return cell;
    }
    if (cell.f) {
        const formula = String(cell.f).startsWith('=') ? String(cell.f).slice(1) : String(cell.f);
        const value = { formula };
        if (cell.v !== undefined && cell.v !== null) {
            value.result = cell.v;
        }
        return value;
    }
    if (cell.v !== undefined && cell.v !== null) {
        return cell.v;
    }
    if (cell.m !== undefined && cell.m !== null) {
        return cell.m;
    }
    return null;
}

function applyCellStyle(excelCell, source) {
    if (!source || typeof source !== 'object') {
        return;
    }

    const font = { name: fontName(source.ff) };
    if (source.fs != null && source.fs !== '') {
        font.size = Number(source.fs);
    }
    if (Number(source.bl) === 1) {
        font.bold = true;
    }
    if (Number(source.it) === 1) {
        font.italic = true;
    }
    if (Number(source.cl) === 1) {
        font.strike = true;
    }
    if (source.un != null && Number(source.un) !== 0) {
        font.underline = Number(source.un) === 2 ? 'double' : true;
    }
    const fontColor = toArgb(source.fc);
    if (fontColor) {
        font.color = { argb: fontColor };
    }
    excelCell.font = font;

    const fillColor = toArgb(source.bg);
    if (fillColor) {
        excelCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
        };
    }

    const alignment = {};
    if (source.ht != null && HT[source.ht] != null) {
        alignment.horizontal = HT[source.ht];
    }
    if (source.vt != null && VT[source.vt] != null) {
        alignment.vertical = VT[source.vt];
    }
    if (Number(source.tb) === 2) {
        alignment.wrapText = true;
    }
    if (source.tr != null && TR[source.tr] != null) {
        alignment.textRotation = TR[source.tr];
    }
    if (Object.keys(alignment).length) {
        excelCell.alignment = alignment;
    }

    if (source.ct && source.ct.fa && source.ct.fa !== 'General') {
        excelCell.numFmt = source.ct.fa;
    }
}

function rowsFromSheet(sheet) {
    if (Array.isArray(sheet.data) && sheet.data.length) {
        return sheet.data;
    }
    if (!Array.isArray(sheet.celldata) || sheet.celldata.length === 0) {
        return [];
    }
    let maxR = 0;
    let maxC = 0;
    sheet.celldata.forEach((item) => {
        if (!item) {
            return;
        }
        if (item.r != null) {
            maxR = Math.max(maxR, item.r);
        }
        if (item.c != null) {
            maxC = Math.max(maxC, item.c);
        }
    });
    const rows = Array.from({ length: maxR + 1 }, () => Array(maxC + 1).fill(null));
    sheet.celldata.forEach((item) => {
        if (item && item.r != null && item.c != null) {
            rows[item.r][item.c] = item.v;
        }
    });
    return rows;
}

function applyMerges(worksheet, sheet) {
    const merges = sheet.config && sheet.config.merge;
    if (!merges || typeof merges !== 'object') {
        return;
    }
    Object.values(merges).forEach((merge) => {
        if (!merge || merge.r == null || merge.c == null) {
            return;
        }
        const rs = merge.rs || 1;
        const cs = merge.cs || 1;
        if (rs <= 1 && cs <= 1) {
            return;
        }
        try {
            worksheet.mergeCells(merge.r + 1, merge.c + 1, merge.r + rs, merge.c + cs);
        } catch (err) {
            // overlapping / already-merged ranges from dirty config must not abort the workbook
        }
    });
}

function applyRowColSize(worksheet, sheet, rowCount, colCount) {
    const config = sheet.config || {};
    const rowlen = config.rowlen || {};
    const columnlen = config.columnlen || {};
    const rowhidden = config.rowhidden || {};
    const colhidden = config.colhidden || {};

    for (let r = 0; r < rowCount; r++) {
        const row = worksheet.getRow(r + 1);
        if (rowhidden[r] != null) {
            row.hidden = true;
        }
        if (rowlen[r] != null) {
            row.height = Number(rowlen[r]) * 0.75;
        }
    }

    for (let c = 0; c < colCount; c++) {
        const column = worksheet.getColumn(c + 1);
        if (colhidden[c] != null) {
            column.hidden = true;
        }
        if (columnlen[c] != null) {
            column.width = Math.max(Number(columnlen[c]) / 8, 1);
        }
    }
}

function applyRangeBorders(borderMap, range, style, color, sides) {
    if (!range || range.row == null || range.column == null) {
        return;
    }
    const r1 = range.row[0];
    const r2 = range.row[1];
    const c1 = range.column[0];
    const c2 = range.column[1];
    const edge = borderEdge(style, color);

    for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
            const key = r + '_' + c;
            if (!borderMap[key]) {
                borderMap[key] = {};
            }
            sides.forEach((side) => {
                borderMap[key][side] = edge;
            });
        }
    }
}

function sidesForBorderType(borderType, r, c, r1, r2, c1, c2) {
    switch (borderType) {
        case 'border-all':
            return ['top', 'bottom', 'left', 'right'];
        case 'border-inside':
            return ['top', 'bottom', 'left', 'right'];
        case 'border-horizontal':
            return ['top', 'bottom'];
        case 'border-vertical':
            return ['left', 'right'];
        case 'border-outside':
        case 'border-outer': {
            const sides = [];
            if (r === r1) {
                sides.push('top');
            }
            if (r === r2) {
                sides.push('bottom');
            }
            if (c === c1) {
                sides.push('left');
            }
            if (c === c2) {
                sides.push('right');
            }
            return sides;
        }
        case 'border-left':
            return c === c1 ? ['left'] : [];
        case 'border-right':
            return c === c2 ? ['right'] : [];
        case 'border-top':
            return r === r1 ? ['top'] : [];
        case 'border-bottom':
            return r === r2 ? ['bottom'] : [];
        case 'border-none':
            return [];
        default:
            return ['top', 'bottom', 'left', 'right'];
    }
}

function collectBorders(sheet) {
    const list = sheet.config && sheet.config.borderInfo;
    const borderMap = {};
    if (!Array.isArray(list)) {
        return borderMap;
    }

    list.forEach((info) => {
        if (!info) {
            return;
        }
        if (info.rangeType === 'cell' && info.value) {
            const r = info.value.row_index;
            const c = info.value.col_index;
            if (r == null || c == null) {
                return;
            }
            const key = r + '_' + c;
            if (!borderMap[key]) {
                borderMap[key] = {};
            }
            ['t', 'b', 'l', 'r'].forEach((side) => {
                const edge = info.value[side];
                if (edge) {
                    const excelSide = { t: 'top', b: 'bottom', l: 'left', r: 'right' }[side];
                    borderMap[key][excelSide] = borderEdge(edge.style, edge.color);
                }
            });
            return;
        }

        const ranges = Array.isArray(info.range) ? info.range : info.range ? [info.range] : [];
        ranges.forEach((range) => {
            if (!range || range.row == null || range.column == null) {
                return;
            }
            const r1 = range.row[0];
            const r2 = range.row[1];
            const c1 = range.column[0];
            const c2 = range.column[1];
            for (let r = r1; r <= r2; r++) {
                for (let c = c1; c <= c2; c++) {
                    const sides = sidesForBorderType(info.borderType, r, c, r1, r2, c1, c2);
                    if (sides.length) {
                        applyRangeBorders(borderMap, { row: [r, r], column: [c, c] }, info.style, info.color, sides);
                    }
                }
            }
        });
    });

    return borderMap;
}

function applyFreeze(worksheet, sheet) {
    const frozen = sheet.frozen || sheet.freezen;
    if (!frozen) {
        return;
    }
    if (frozen.type === 'row' || frozen.horizontal) {
        worksheet.views = [{ state: 'frozen', ySplit: (frozen.range && frozen.range.row_focus != null ? frozen.range.row_focus : 0) + 1 }];
        return;
    }
    if (frozen.type === 'column' || frozen.vertical) {
        worksheet.views = [{ state: 'frozen', xSplit: (frozen.range && frozen.range.column_focus != null ? frozen.range.column_focus : 0) + 1 }];
        return;
    }
    if (frozen.type === 'both' || (frozen.horizontal && frozen.vertical)) {
        worksheet.views = [
            {
                state: 'frozen',
                ySplit: (frozen.range && frozen.range.row_focus != null ? frozen.range.row_focus : 0) + 1,
                xSplit: (frozen.range && frozen.range.column_focus != null ? frozen.range.column_focus : 0) + 1,
            },
        ];
    }
}

function writeSheet(worksheet, sheet) {
    const rows = rowsFromSheet(sheet);
    const borderMap = collectBorders(sheet);
    let maxCol = 0;

    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row) {
            continue;
        }
        maxCol = Math.max(maxCol, row.length);
        for (let c = 0; c < row.length; c++) {
            const source = row[c];
            const excelCell = worksheet.getCell(r + 1, c + 1);
            const value = cellValue(source);
            if (value !== null) {
                excelCell.value = value;
            }
            applyCellStyle(excelCell, source && typeof source === 'object' ? source : null);
            const border = borderMap[r + '_' + c];
            if (border) {
                excelCell.border = border;
            }
        }
    }

    Object.keys(borderMap).forEach((key) => {
        const parts = key.split('_');
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        if (r >= rows.length || !rows[r] || c >= (rows[r] ? rows[r].length : 0)) {
            const excelCell = worksheet.getCell(r + 1, c + 1);
            excelCell.border = borderMap[key];
            maxCol = Math.max(maxCol, c + 1);
        }
    });

    applyMerges(worksheet, sheet);
    applyRowColSize(worksheet, sheet, Math.max(rows.length, 1), Math.max(maxCol, 1));
    applyFreeze(worksheet, sheet);
}

function selectExportSheets(payload) {
    const sheets = Array.isArray(payload.data) ? payload.data : [];
    const order = payload.exportXlsx && payload.exportXlsx.order;
    if (order === undefined || order === 'all') {
        return sheets;
    }
    const index = Number(order);
    if (!Number.isNaN(index) && sheets[index]) {
        return [sheets[index]];
    }
    return sheets;
}

const fn_luckyToXlsx = async (ctx) => {
    const payload = ctx.request.body || {};
    const exportSheets = selectExportSheets(payload);

    if (exportSheets.length === 0) {
        ctx.status = 400;
        ctx.body = { error: 'No sheet data to export' };
        return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'luckysheet-local';
    workbook.created = new Date();

    exportSheets.forEach((sheet, index) => {
        const name = (sheet && sheet.name) || `Sheet${index + 1}`;
        const worksheet = workbook.addWorksheet(String(name).substring(0, 31) || `Sheet${index + 1}`);
        writeSheet(worksheet, sheet || {});
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `${String(payload.title || 'luckysheet').replace(/[\\/:*?"<>|]/g, '_')}.xlsx`;

    ctx.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    ctx.set('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    ctx.body = buffer;
};

module.exports = {
    'POST /luckyToXlsx': fn_luckyToXlsx,
    cellValue,
    applyCellStyle,
    writeSheet,
    toArgb,
    fontName,
};
