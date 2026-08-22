/**
 * 结构化 Table 最小控制器。
 * sheet.table[] 契约：{ id, name, range, columns, tableStyleId }
 * 对照 UV @univerjs/sheets-table，不抄 Pro / UI 包。
 */

import { getSheetIndex, getRangetxt } from "../methods/get";
import { chatatABC } from "../utils/util";
import { createFilter } from "./filter";
import server from "./server";
import Store from "../store";

function createId(prefix) {
    return prefix + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function cloneRange(range) {
    return {
        row: [range.row[0], range.row[1]],
        column: [range.column[0], range.column[1]],
    };
}

function columnsFromRange(range, data, hasHeader) {
    const columns = [];
    const headerRow = hasHeader ? range.row[0] : null;
    for (let c = range.column[0]; c <= range.column[1]; c++) {
        let name = chatatABC(c);
        if (headerRow != null && data && data[headerRow] && data[headerRow][c]) {
            const cell = data[headerRow][c];
            const text = cell.m != null ? cell.m : cell.v;
            if (text != null && String(text).replace(/\s/g, "") !== "") {
                name = String(text);
            }
        }
        columns.push({
            id: createId("col"),
            name: name,
            index: c - range.column[0],
        });
    }
    return columns;
}

function normalizeTable(table, data) {
    if (!table || typeof table !== "object") {
        return null;
    }
    const range = table.range && table.range.row && table.range.column
        ? cloneRange(table.range)
        : { row: [0, 0], column: [0, 0] };
    const columns = Array.isArray(table.columns) && table.columns.length
        ? table.columns.map((col, index) => ({
            id: col.id || createId("col"),
            name: col.name || chatatABC(range.column[0] + index),
            index: col.index != null ? col.index : index,
        }))
        : columnsFromRange(range, data, table.hasHeader !== false);

    return {
        id: table.id || createId("tbl"),
        name: table.name || "Table",
        range: range,
        columns: columns,
        tableStyleId: table.tableStyleId || "TableStyleMedium2",
        hasHeader: table.hasHeader !== false,
        hasTotalRow: !!table.hasTotalRow,
    };
}

const tableCtrl = {
    getSheetTables: function(sheetIndex) {
        const index = getSheetIndex(sheetIndex != null ? sheetIndex : Store.currentSheetIndex);
        if (index == null) {
            return [];
        }
        const file = Store.luckysheetfile[index];
        if (!file) {
            return [];
        }
        if (!Array.isArray(file.table)) {
            file.table = [];
        }
        file.table = file.table.map((item) => normalizeTable(item, file.data)).filter(Boolean);
        return file.table;
    },
    getTableById: function(tableId, sheetIndex) {
        return this.getSheetTables(sheetIndex).find((item) => item.id === tableId) || null;
    },
    getTableByName: function(name, sheetIndex) {
        return this.getSheetTables(sheetIndex).find((item) => item.name === name) || null;
    },
    getTableByCell: function(r, c, sheetIndex) {
        const tables = this.getSheetTables(sheetIndex);
        for (let i = 0; i < tables.length; i++) {
            const range = tables[i].range;
            if (r >= range.row[0] && r <= range.row[1] && c >= range.column[0] && c <= range.column[1]) {
                return tables[i];
            }
        }
        return null;
    },
    addTable: function(input, sheetIndex) {
        const current = sheetIndex != null ? sheetIndex : Store.currentSheetIndex;
        const index = getSheetIndex(current);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return null;
        }
        if (!Array.isArray(file.table)) {
            file.table = [];
        }

        let range = input && input.range;
        if (!range && Store.luckysheet_select_save && Store.luckysheet_select_save.length) {
            const last = Store.luckysheet_select_save[Store.luckysheet_select_save.length - 1];
            range = { row: last.row.slice(), column: last.column.slice() };
        }
        if (!range) {
            return null;
        }

        const table = normalizeTable({
            id: input && input.id,
            name: (input && input.name) || ("Table" + (file.table.length + 1)),
            range: range,
            columns: input && input.columns,
            tableStyleId: input && input.tableStyleId,
            hasHeader: input && input.hasHeader,
            hasTotalRow: input && input.hasTotalRow,
        }, file.data);

        file.table.push(table);
        this.persist(current);
        return table;
    },
    removeTable: function(tableId, sheetIndex) {
        const current = sheetIndex != null ? sheetIndex : Store.currentSheetIndex;
        const index = getSheetIndex(current);
        const file = Store.luckysheetfile[index];
        if (!file || !Array.isArray(file.table)) {
            return false;
        }
        const next = file.table.filter((item) => item.id !== tableId);
        if (next.length === file.table.length) {
            return false;
        }
        file.table = next;
        this.persist(current);
        return true;
    },
    bindFilter: function(tableId, sheetIndex) {
        const table = this.getTableById(tableId, sheetIndex);
        if (!table) {
            return false;
        }
        Store.luckysheet_select_save = [cloneRange(table.range)];
        createFilter();
        return true;
    },
    getColumnByName: function(table, columnName) {
        if (!table || !table.columns) {
            return null;
        }
        return table.columns.find((col) => col.name === columnName || col.id === columnName) || null;
    },
    getRangeTxt: function(table, sheetIndex) {
        if (!table || !table.range) {
            return "";
        }
        const file = Store.luckysheetfile[getSheetIndex(sheetIndex != null ? sheetIndex : Store.currentSheetIndex)];
        return getRangetxt(file ? file.index : Store.currentSheetIndex, table.range);
    },
    persist: function(sheetIndex) {
        const current = sheetIndex != null ? sheetIndex : Store.currentSheetIndex;
        const index = getSheetIndex(current);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return;
        }
        if (server.allowUpdate) {
            server.saveParam("all", current, file.table || [], { k: "table" });
        }
    },
    init: function() {
        if (!Store.luckysheetfile) {
            return;
        }
        Store.luckysheetfile.forEach((file) => {
            if (file && Array.isArray(file.table)) {
                file.table = file.table.map((item) => normalizeTable(item, file.data)).filter(Boolean);
            }
        });
    },
};

export default tableCtrl;
