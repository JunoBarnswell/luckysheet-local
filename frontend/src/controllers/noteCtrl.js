/**
 * Cell / Sheet Note。
 * 与 postil `cell.ps` 批注并存：ps 是可编辑批注框，note 是角标 + hover 文本。
 * 存储：
 * - cell.note: { text } | string
 * - sheet.notes: { "r_c": { text } }
 * - sheet.note: { text } 工作表级备注
 */

import { rowLocation, colLocation, mouseposition } from "../global/location";
import luckysheetFreezen from "./freezen";
import menuButton from "./menuButton";
import { getSheetIndex } from "../methods/get";
import server from "./server";
import Store from "../store";
import locale from "../locale/locale";

function asNote(value) {
    if (value == null) {
        return null;
    }
    if (typeof value === "string") {
        const text = value.replace(/\s/g, "") === "" ? "" : value;
        return text === "" ? null : { text: value };
    }
    if (typeof value === "object") {
        const text = value.text != null ? value.text : value.content;
        if (text == null || String(text).replace(/\s/g, "") === "") {
            return null;
        }
        return {
            text: String(text),
            width: value.width,
            height: value.height,
        };
    }
    return { text: String(value) };
}

const noteCtrl = {
    getCellNote: function(r, c, sheetIndex) {
        const index = getSheetIndex(sheetIndex != null ? sheetIndex : Store.currentSheetIndex);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return null;
        }

        const data = file.data || Store.flowdata;
        const cell = data && data[r] ? data[r][c] : null;
        const fromCell = asNote(cell && cell.note);
        if (fromCell) {
            return fromCell;
        }

        if (file.notes && file.notes[r + "_" + c]) {
            return asNote(file.notes[r + "_" + c]);
        }
        return null;
    },
    getSheetNote: function(sheetIndex) {
        const index = getSheetIndex(sheetIndex != null ? sheetIndex : Store.currentSheetIndex);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return null;
        }
        return asNote(file.note);
    },
    setCellNote: function(r, c, note, sheetIndex) {
        const current = sheetIndex != null ? sheetIndex : Store.currentSheetIndex;
        const index = getSheetIndex(current);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return;
        }
        if (!file.data) {
            file.data = Store.flowdata;
        }
        if (!file.data[r]) {
            file.data[r] = [];
        }
        if (file.data[r][c] == null) {
            file.data[r][c] = {};
        }
        const normalized = asNote(note);
        if (normalized) {
            file.data[r][c].note = normalized;
        } else {
            delete file.data[r][c].note;
        }
        if (!file.notes) {
            file.notes = {};
        }
        if (normalized) {
            file.notes[r + "_" + c] = normalized;
        } else {
            delete file.notes[r + "_" + c];
        }
        if (current === Store.currentSheetIndex && Store.flowdata[r]) {
            if (Store.flowdata[r][c] == null) {
                Store.flowdata[r][c] = {};
            }
            if (normalized) {
                Store.flowdata[r][c].note = normalized;
            } else if (Store.flowdata[r][c]) {
                delete Store.flowdata[r][c].note;
            }
        }
        if (server.allowUpdate) {
            server.saveParam("all", current, file.notes, { k: "notes" });
        }
    },
    setSheetNote: function(note, sheetIndex) {
        const current = sheetIndex != null ? sheetIndex : Store.currentSheetIndex;
        const index = getSheetIndex(current);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return;
        }
        const normalized = asNote(note);
        if (normalized) {
            file.note = normalized;
        } else {
            delete file.note;
        }
        this.syncSheetTabTitle(current);
        if (server.allowUpdate) {
            server.saveParam("all", current, file.note || null, { k: "note" });
        }
    },
    hasNote: function(r, c, sheetIndex) {
        return this.getCellNote(r, c, sheetIndex) != null;
    },
    syncSheetTabTitle: function(sheetIndex) {
        const index = getSheetIndex(sheetIndex != null ? sheetIndex : Store.currentSheetIndex);
        const file = Store.luckysheetfile[index];
        if (!file) {
            return;
        }
        const $item = $("#luckysheet-sheets-item" + file.index);
        if (!$item.length) {
            return;
        }
        const note = asNote(file.note);
        const _locale = locale();
        const tip = note ? note.text : "";
        $item.attr("data-sheet-note", tip);
        $item.attr("title", tip || file.name);
        if (tip) {
            $item.addClass("luckysheet-sheets-item-has-note");
        } else {
            $item.removeClass("luckysheet-sheets-item-has-note");
        }
        if (tip && _locale.note) {
            $item.attr("title", file.name + " — " + tip);
        }
    },
    overshow: function(event) {
        $("#luckysheet-note-overshow").remove();

        if ($(event.target).closest("#luckysheet-cell-main").length == 0) {
            const $sheet = $(event.target).closest(".luckysheet-sheets-item");
            if ($sheet.length) {
                const tip = $sheet.attr("data-sheet-note");
                if (tip) {
                    const html = `<div id="luckysheet-note-overshow" class="luckysheet-mousedown-cancel" style="background:#fffbe6;padding:6px 10px;border:1px solid #d4b106;box-shadow:2px 2px #999;position:fixed;left:${event.pageX + 12}px;top:${event.pageY + 12}px;z-index:100;max-width:280px;white-space:pre-wrap;">${$("<div/>").text(tip).html()}</div>`;
                    $(html).appendTo($("body"));
                }
            }
            return;
        }

        const mouse = mouseposition(event.pageX, event.pageY);
        const scrollLeft = $("#luckysheet-cell-main").scrollLeft();
        const scrollTop = $("#luckysheet-cell-main").scrollTop();
        const x = mouse[0] + scrollLeft;
        const y = mouse[1] + scrollTop;

        if (luckysheetFreezen.freezenverticaldata != null && mouse[0] < (luckysheetFreezen.freezenverticaldata[0] - luckysheetFreezen.freezenverticaldata[2])) {
            return;
        }
        if (luckysheetFreezen.freezenhorizontaldata != null && mouse[1] < (luckysheetFreezen.freezenhorizontaldata[0] - luckysheetFreezen.freezenhorizontaldata[2])) {
            return;
        }

        let row_index = rowLocation(y)[2];
        let col_index = colLocation(x)[2];
        const margeset = menuButton.mergeborer(Store.flowdata, row_index, col_index);
        if (margeset) {
            row_index = margeset.row[2];
            col_index = margeset.column[2];
        }

        const note = this.getCellNote(row_index, col_index);
        if (!note) {
            return;
        }

        let row = Store.visibledatarow[row_index];
        let row_pre = row_index - 1 == -1 ? 0 : Store.visibledatarow[row_index - 1];
        let col = Store.visibledatacolumn[col_index];
        let col_pre = col_index - 1 == -1 ? 0 : Store.visibledatacolumn[col_index - 1];
        if (margeset) {
            row = margeset.row[1];
            row_pre = margeset.row[0];
            col = margeset.column[1];
            col_pre = margeset.column[0];
        }

        const html = `<div id="luckysheet-note-overshow" class="luckysheet-mousedown-cancel" style="background:#fffbe6;padding:6px 10px;border:1px solid #d4b106;box-shadow:2px 2px #999;position:absolute;left:${col_pre}px;top:${row + 5}px;z-index:100;max-width:280px;white-space:pre-wrap;">${$("<div/>").text(note.text).html()}</div>`;
        $(html).appendTo($("#luckysheet-cell-main"));
    },
    init: function() {
        if (!Store.luckysheetfile) {
            return;
        }
        Store.luckysheetfile.forEach((file) => {
            if (!file) {
                return;
            }
            if (file.notes && typeof file.notes === "object" && file.data) {
                Object.keys(file.notes).forEach((key) => {
                    const parts = key.split("_");
                    const r = parseInt(parts[0], 10);
                    const c = parseInt(parts[1], 10);
                    if (isNaN(r) || isNaN(c)) {
                        return;
                    }
                    if (!file.data[r]) {
                        return;
                    }
                    if (file.data[r][c] == null) {
                        file.data[r][c] = {};
                    }
                    if (file.data[r][c].note == null) {
                        file.data[r][c].note = asNote(file.notes[key]);
                    }
                });
            }
            this.syncSheetTabTitle(file.index);
        });
    },
};

export default noteCtrl;
