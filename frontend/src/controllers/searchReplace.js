import { replaceHtml, chatatABC } from "../utils/util";
import { getSheetIndex } from "../methods/get";
import { modelHTML, keycode } from "./constant";
import { selectHightlightShow } from "./select";
import sheetmanage from "./sheetmanage";
import { isEditMode } from "../global/validate";
import { valueShowEs } from "../global/format";
import { setcellvalue } from "../global/setdata";
import { jfrefreshgrid } from "../global/refresh";
import editor from "../global/editor";
import tooltip from "../global/tooltip";
import func_methods from "../global/func_methods";
import Store from "../store";
import locale from "../locale/locale";
import { checkProtectionLocked } from "./protection";
import escapeHtml from "escape-html";

//查找替换
const luckysheetSearchReplace = {
    createDialog: function(source) {
        $("#luckysheet-modal-dialog-mask").hide();
        $("#luckysheet-search-replace").remove();

        const _locale = locale();
        const locale_findAndReplace = _locale.findAndReplace;
        const locale_button = _locale.button;

        let content =
            '<div class="tabBox">' +
            '<span id="searchTab">' +
            locale_findAndReplace.find +
            "</span>" +
            '<span id="replaceTab">' +
            locale_findAndReplace.replace +
            "</span>" +
            "</div>" +
            '<div class="ctBox">' +
            '<div class="inputBox">' +
            '<div class="textboxs" id="searchInput">' +
            locale_findAndReplace.findTextbox +
            '：<input class="formulaInputFocus" spellcheck="false" value=""/></div>' +
            '<div class="textboxs" id="replaceInput">' +
            locale_findAndReplace.replaceTextbox +
            '：<input class="formulaInputFocus" spellcheck="false" value=""/></div>' +
            '<div class="checkboxs">' +
            '<div id="regCheck">' +
            '<input type="checkbox"/>' +
            "<span>" +
            locale_findAndReplace.regexTextbox +
            "</span>" +
            "</div>" +
            '<div id="wordCheck">' +
            '<input type="checkbox"/>' +
            "<span>" +
            locale_findAndReplace.wholeTextbox +
            "</span>" +
            "</div>" +
            '<div id="caseCheck">' +
            '<input type="checkbox"/>' +
            "<span>" +
            locale_findAndReplace.distinguishTextbox +
            "</span>" +
            "</div>" +
            '<div id="formulaCheck">' +
            '<input type="checkbox"/>' +
            "<span>" +
            (locale_findAndReplace.formulaTextbox || "Find in formula") +
            "</span>" +
            "</div>" +
            '<div id="workbookCheck">' +
            '<input type="checkbox"/>' +
            "<span>" +
            (locale_findAndReplace.workbookTextbox || "Search workbook") +
            "</span>" +
            "</div>" +
            "</div>" +
            "</div>" +
            '<div class="btnBox">' +
            '<button id="replaceAllBtn" class="btn btn-default">' +
            locale_findAndReplace.allReplaceBtn +
            "</button>" +
            '<button id="replaceBtn" class="btn btn-default">' +
            locale_findAndReplace.replaceBtn +
            "</button>" +
            '<button id="searchAllBtn" class="btn btn-default">' +
            locale_findAndReplace.allFindBtn +
            "</button>" +
            '<button id="searchNextBtn" class="btn btn-default">' +
            locale_findAndReplace.findBtn +
            "</button>" +
            "</div>" +
            "</div>";

        $("body").append(
            replaceHtml(modelHTML, {
                id: "luckysheet-search-replace",
                addclass: "luckysheet-search-replace",
                title: "",
                content: content,
                botton:
                    '<button class="btn btn-default luckysheet-model-close-btn">' + locale_button.close + "</button>",
                style: "z-index:100003",
                close: locale_button.close,
            }),
        );
        let $t = $("#luckysheet-search-replace")
                .find(".luckysheet-modal-dialog-content")
                .css("min-width", 500)
                .end(),
            myh = $t.outerHeight(),
            myw = $t.outerWidth();
        let winw = $(window).width(),
            winh = $(window).height();
        let scrollLeft = $(document).scrollLeft(),
            scrollTop = $(document).scrollTop();
        $("#luckysheet-search-replace")
            .attr("data-ls-instance", Store.instanceId || "")
            .css({ left: (winw + scrollLeft - myw) / 2, top: (winh + scrollTop - myh) / 3 })
            .show();

        if (source == "0") {
            $("#luckysheet-search-replace #searchTab")
                .addClass("on")
                .siblings()
                .removeClass("on");
            $("#luckysheet-search-replace #replaceInput").hide();
            $("#luckysheet-search-replace #replaceAllBtn").hide();
            $("#luckysheet-search-replace #replaceBtn").hide();
        } else if (source == "1") {
            $("#luckysheet-search-replace #replaceTab")
                .addClass("on")
                .siblings()
                .removeClass("on");
            $("#luckysheet-search-replace #replaceInput").show();
            $("#luckysheet-search-replace #replaceAllBtn").show();
            $("#luckysheet-search-replace #replaceBtn").show();
        }
    },
    init: function() {
        let _this = this;

        //查找替换 切换
        $(document)
            .off("click.SRtabBoxspan")
            .on("click.SRtabBoxspan", "#luckysheet-search-replace .tabBox span", function() {
                $(this)
                    .addClass("on")
                    .siblings()
                    .removeClass("on");

                let $id = $(this).attr("id");
                if ($id == "searchTab") {
                    $("#luckysheet-search-replace #replaceInput").hide();
                    $("#luckysheet-search-replace #replaceAllBtn").hide();
                    $("#luckysheet-search-replace #replaceBtn").hide();

                    $("#luckysheet-search-replace #searchInput input").focus();
                } else if ($id == "replaceTab") {
                    $("#luckysheet-search-replace #replaceInput").show();
                    $("#luckysheet-search-replace #replaceAllBtn").show();
                    $("#luckysheet-search-replace #replaceBtn").show();

                    $("#luckysheet-search-replace #replaceInput input").focus();
                }
            });

        //查找下一个
        $(document)
            .off("keyup.SRsearchInput")
            .on("keyup.SRsearchInput", "#luckysheet-search-replace #searchInput input", function(event) {
                let kcode = event.keyCode;
                if (kcode == keycode.ENTER) {
                    _this.searchNext();
                }
            });
        $(document)
            .off("click.SRsearchNextBtn")
            .on("click.SRsearchNextBtn", "#luckysheet-search-replace #searchNextBtn", function() {
                _this.searchNext();
            });

        //查找全部
        $(document)
            .off("click.SRsearchAllBtn")
            .on("click.SRsearchAllBtn", "#luckysheet-search-replace #searchAllBtn", function() {
                _this.searchAll();
            });
        $(document)
            .off("click.SRsearchAllboxItem")
            .on("click.SRsearchAllboxItem", "#luckysheet-search-replace #searchAllbox .boxItem", function() {
                $(this)
                    .addClass("on")
                    .siblings()
                    .removeClass("on");

                let r = $(this).attr("data-row");
                let c = $(this).attr("data-col");
                let sheetIndex = $(this).attr("data-sheetIndex");

                _this.locateSearchHit(r, c, sheetIndex);

                let scrollLeft = $("#luckysheet-cell-main").scrollLeft(),
                    scrollTop = $("#luckysheet-cell-main").scrollTop();
                let winH = $("#luckysheet-cell-main").height(),
                    winW = $("#luckysheet-cell-main").width();

                let row = Store.visibledatarow[r],
                    row_pre = r - 1 == -1 ? 0 : Store.visibledatarow[r - 1];
                let col = Store.visibledatacolumn[c],
                    col_pre = c - 1 == -1 ? 0 : Store.visibledatacolumn[c - 1];

                if (col - scrollLeft - winW + 20 > 0) {
                    $("#luckysheet-scrollbar-x").scrollLeft(col - winW + 20);
                } else if (col_pre - scrollLeft - 20 < 0) {
                    $("#luckysheet-scrollbar-x").scrollLeft(col_pre - 20);
                }

                if (row - scrollTop - winH + 20 > 0) {
                    $("#luckysheet-scrollbar-y").scrollTop(row - winH + 20);
                } else if (row_pre - scrollTop - 20 < 0) {
                    $("#luckysheet-scrollbar-y").scrollTop(row_pre - 20);
                }
            });

        //替换
        $(document)
            .off("click.SRreplaceBtn")
            .on("click.SRreplaceBtn", "#luckysheet-search-replace #replaceBtn", function() {
                _this.replace();
            });

        //全部替换
        $(document)
            .off("click.SRreplaceAllBtn")
            .on("click.SRreplaceAllBtn", "#luckysheet-search-replace #replaceAllBtn", function() {
                _this.replaceAll();
            });
    },
    locateSearchHit: function(r, c, sheetIndex) {
        r = parseInt(r, 10);
        c = parseInt(c, 10);

        if (sheetIndex != null && String(sheetIndex) != String(Store.currentSheetIndex)) {
            sheetmanage.changeSheetExec(sheetIndex);
        }

        Store.luckysheet_select_save = [{ row: [r, r], column: [c, c] }];
        selectHightlightShow();

        let scrollLeft = $("#luckysheet-cell-main").scrollLeft(),
            scrollTop = $("#luckysheet-cell-main").scrollTop();
        let winH = $("#luckysheet-cell-main").height(),
            winW = $("#luckysheet-cell-main").width();

        let row = Store.visibledatarow[r],
            row_pre = r - 1 == -1 ? 0 : Store.visibledatarow[r - 1];
        let col = Store.visibledatacolumn[c],
            col_pre = c - 1 == -1 ? 0 : Store.visibledatacolumn[c - 1];

        if (col - scrollLeft - winW + 20 > 0) {
            $("#luckysheet-scrollbar-x").scrollLeft(col - winW + 20);
        } else if (col_pre - scrollLeft - 20 < 0) {
            $("#luckysheet-scrollbar-x").scrollLeft(col_pre - 20);
        }

        if (row - scrollTop - winH + 20 > 0) {
            $("#luckysheet-scrollbar-y").scrollTop(row - winH + 20);
        } else if (row_pre - scrollTop - 20 < 0) {
            $("#luckysheet-scrollbar-y").scrollTop(row_pre - 20);
        }
    },
    searchNext: function() {
        let _this = this;

        let searchText = $("#luckysheet-search-replace #searchInput input").val();
        if (searchText == "" || searchText == null) {
            return;
        }
        const _locale = locale();
        const locale_findAndReplace = _locale.findAndReplace;
        let range;
        if (
            Store.luckysheet_select_save.length == 0 ||
            (Store.luckysheet_select_save.length == 1 &&
                Store.luckysheet_select_save[0].row[0] == Store.luckysheet_select_save[0].row[1] &&
                Store.luckysheet_select_save[0].column[0] == Store.luckysheet_select_save[0].column[1])
        ) {
            range = [
                {
                    row: [0, Store.flowdata.length - 1],
                    column: [0, Store.flowdata[0].length - 1],
                },
            ];
        } else {
            range = $.extend(true, [], Store.luckysheet_select_save);
        }

        let searchIndexArr = _this.getSearchHits(searchText, range);

        if (searchIndexArr.length == 0) {
            if (isEditMode()) {
                alert(locale_findAndReplace.noFindTip);
            } else {
                tooltip.info(locale_findAndReplace.noFindTip, "");
            }

            return;
        }

        let count = 0;

        if (
            Store.luckysheet_select_save.length == 0 ||
            (Store.luckysheet_select_save.length == 1 &&
                Store.luckysheet_select_save[0].row[0] == Store.luckysheet_select_save[0].row[1] &&
                Store.luckysheet_select_save[0].column[0] == Store.luckysheet_select_save[0].column[1])
        ) {
            if (Store.luckysheet_select_save.length == 0) {
                count = 0;
            } else {
                for (let i = 0; i < searchIndexArr.length; i++) {
                    if (
                        searchIndexArr[i].r == Store.luckysheet_select_save[0].row[0] &&
                        searchIndexArr[i].c == Store.luckysheet_select_save[0].column[0] &&
                        String(searchIndexArr[i].sheetIndex) == String(Store.currentSheetIndex)
                    ) {
                        if (i == searchIndexArr.length - 1) {
                            count = 0;
                        } else {
                            count = i + 1;
                        }

                        break;
                    }
                }
            }

            this.locateSearchHit(
                searchIndexArr[count].r,
                searchIndexArr[count].c,
                searchIndexArr[count].sheetIndex,
            );
            return;
        } else {
            let rf = range[range.length - 1].row_focus;
            let cf = range[range.length - 1].column_focus;

            for (let i = 0; i < searchIndexArr.length; i++) {
                if (searchIndexArr[i].r == rf && searchIndexArr[i].c == cf) {
                    if (i == searchIndexArr.length - 1) {
                        count = 0;
                    } else {
                        count = i + 1;
                    }

                    break;
                }
            }

            for (let s = 0; s < range.length; s++) {
                let r1 = range[s].row[0],
                    r2 = range[s].row[1];
                let c1 = range[s].column[0],
                    c2 = range[s].column[1];

                if (
                    searchIndexArr[count].r >= r1 &&
                    searchIndexArr[count].r <= r2 &&
                    searchIndexArr[count].c >= c1 &&
                    searchIndexArr[count].c <= c2
                ) {
                    let obj = range[s];
                    obj["row_focus"] = searchIndexArr[count].r;
                    obj["column_focus"] = searchIndexArr[count].c;
                    range.splice(s, 1);
                    range.push(obj);

                    break;
                }
            }

            Store.luckysheet_select_save = range;
        }

        this.locateSearchHit(
            searchIndexArr[count].r,
            searchIndexArr[count].c,
            searchIndexArr[count].sheetIndex,
        );

        if ($("#searchAllbox").is(":visible")) {
            $("#luckysheet-search-replace #searchAllbox .boxItem").removeClass("on");
        }
    },
    searchAll: function() {
        let _this = this;

        const _locale = locale();
        const locale_findAndReplace = _locale.findAndReplace;

        $("#luckysheet-search-replace #searchAllbox").remove();

        let searchText = $("#luckysheet-search-replace #searchInput input").val();
        if (searchText == "" || searchText == null) {
            return;
        }

        /**
         * fix #1115 查找改为全局查找（todo: 后续可以传入range）
         */
        let range;
        // if(Store.luckysheet_select_save.length == 0 || (Store.luckysheet_select_save.length == 1 && Store.luckysheet_select_save[0].row[0] == Store.luckysheet_select_save[0].row[1] && Store.luckysheet_select_save[0].column[0] == Store.luckysheet_select_save[0].column[1])){
        range = [
            {
                row: [0, Store.flowdata.length - 1],
                column: [0, Store.flowdata[0].length - 1],
            },
        ];
        // }
        // else{
        //     range = $.extend(true, [], Store.luckysheet_select_save);
        // }

        let searchIndexArr = _this.getSearchHits(searchText, range);

        if (searchIndexArr.length == 0) {
            if (isEditMode()) {
                alert(locale_findAndReplace.noFindTip);
            } else {
                tooltip.info(locale_findAndReplace.noFindTip, "");
            }

            return;
        }

        let searchAllHtml = "";

        for (let i = 0; i < searchIndexArr.length; i++) {
            const hit = searchIndexArr[i];
            const file = Store.luckysheetfile[getSheetIndex(hit.sheetIndex != null ? hit.sheetIndex : Store.currentSheetIndex)];
            const data = file && file.data ? file.data : Store.flowdata;
            let value_ShowEs = valueShowEs(hit.r, hit.c, data);
            if (value_ShowEs == null && data[hit.r] && data[hit.r][hit.c] && data[hit.r][hit.c].f) {
                value_ShowEs = data[hit.r][hit.c].f;
            }
            value_ShowEs = value_ShowEs == null ? "" : value_ShowEs.toString();
            const sheetName = file ? file.name : "";
            const sheetIndex = hit.sheetIndex != null ? hit.sheetIndex : Store.currentSheetIndex;

            searchAllHtml +=
                '<div class="boxItem" data-row="' +
                hit.r +
                '" data-col="' +
                hit.c +
                '" data-sheetIndex="' +
                sheetIndex +
                '">' +
                "<span>" +
                escapeHtml(sheetName) +
                "</span>" +
                "<span>" +
                chatatABC(hit.c) +
                (hit.r + 1) +
                "</span>" +
                '<span title="' +
                escapeHtml(value_ShowEs) +
                '">' +
                escapeHtml(value_ShowEs) +
                "</span>" +
                "</div>";
        }

        $(
            `<div id="searchAllbox"><div class="boxTitle"><span>${locale_findAndReplace.searchTargetSheet}</span><span>${locale_findAndReplace.searchTargetCell}</span><span>${locale_findAndReplace.searchTargetValue}</span></div><div class="boxMain">${searchAllHtml}</div></div>`,
        ).appendTo($("#luckysheet-search-replace"));

        $("#luckysheet-search-replace #searchAllbox .boxItem")
            .eq(0)
            .addClass("on")
            .siblings()
            .removeClass("on");

        this.locateSearchHit(
            searchIndexArr[0].r,
            searchIndexArr[0].c,
            searchIndexArr[0].sheetIndex,
        );
    },
    getSearchOptions: function() {
        const $container = $("#luckysheet-search-replace");
        const isChecked = (inputId) => $container.find(`#${inputId} input[type='checkbox']`).is(":checked");
        return {
            regCheck: isChecked("regCheck"),
            wordCheck: isChecked("wordCheck"),
            caseCheck: isChecked("caseCheck"),
            formulaCheck: isChecked("formulaCheck"),
            workbookCheck: isChecked("workbookCheck"),
        };
    },
    matchSearchText: function(value, searchText, options) {
        if (value == null || value === "") {
            return false;
        }
        value = value.toString();
        if (value == 0) {
            value = value.toString();
        }
        let needle = searchText;
        if (!options.caseCheck) {
            value = value.toLowerCase();
            needle = searchText.toLowerCase();
        }
        if (options.wordCheck) {
            return needle == value;
        }
        if (options.regCheck) {
            let flags = "g";
            if (!options.caseCheck) {
                flags += "i";
            }
            let reg = new RegExp(func_methods.getRegExpStr(needle), flags);
            return reg.test(value);
        }
        return value.indexOf(needle) != -1;
    },
    getSearchIndexArr: function(searchText, range, data, sheetIndex) {
        const arr = [];
        const obj = {};
        const options = this.getSearchOptions();
        data = data || Store.flowdata;
        if (!data || !data.length) {
            return arr;
        }

        const addResult = (r, c) => {
            const key = (sheetIndex != null ? sheetIndex + "_" : "") + r + "_" + c;
            if (!(key in obj)) {
                obj[key] = 0;
                arr.push({ r: r, c: c, sheetIndex: sheetIndex != null ? sheetIndex : Store.currentSheetIndex });
            }
        };

        const matchCell = (r, c) => {
            const cell = data[r] && data[r][c];
            if (cell == null) {
                return;
            }
            let value = valueShowEs(r, c, data);
            if (this.matchSearchText(value, searchText, options)) {
                addResult(r, c);
                return;
            }
            if (options.formulaCheck && cell.f != null && this.matchSearchText(cell.f, searchText, options)) {
                addResult(r, c);
            }
        };

        for (let s = 0; s < range.length; s++) {
            const r1 = range[s].row[0],
                r2 = Math.min(range[s].row[1], data.length - 1);
            const c1 = range[s].column[0],
                c2 = Math.min(range[s].column[1], (data[0] && data[0].length ? data[0].length : 1) - 1);

            for (let r = r1; r <= r2; r++) {
                for (let c = c1; c <= c2; c++) {
                    matchCell(r, c);
                }
            }
        }

        return arr;
    },
    getSearchHits: function(searchText, range) {
        const options = this.getSearchOptions();
        if (!options.workbookCheck) {
            return this.getSearchIndexArr(searchText, range, Store.flowdata, Store.currentSheetIndex);
        }

        const hits = [];
        const files = Store.luckysheetfile || [];
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (!file || file.hide == 1) {
                continue;
            }
            const data = file.data;
            if (!data || !data.length) {
                continue;
            }
            const sheetRange = [
                {
                    row: [0, data.length - 1],
                    column: [0, data[0].length - 1],
                },
            ];
            hits.push.apply(hits, this.getSearchIndexArr(searchText, sheetRange, data, file.index));
        }
        return hits;
    },
    replace: function() {
        let _this = this;

        const _locale = locale();
        const locale_findAndReplace = _locale.findAndReplace;

        if (!Store.allowEdit) {
            tooltip.info(locale_findAndReplace.modeTip, "");
            return;
        }

        let searchText = $("#luckysheet-search-replace #searchInput input").val();
        if (searchText == "" || searchText == null) {
            if (isEditMode()) {
                alert(locale_findAndReplace.searchInputTip);
            } else {
                tooltip.info(locale_findAndReplace.searchInputTip, "");
            }

            return;
        }

        let range;
        if (
            Store.luckysheet_select_save.length == 0 ||
            (Store.luckysheet_select_save.length == 1 &&
                Store.luckysheet_select_save[0].row[0] == Store.luckysheet_select_save[0].row[1] &&
                Store.luckysheet_select_save[0].column[0] == Store.luckysheet_select_save[0].column[1])
        ) {
            range = [
                {
                    row: [0, Store.flowdata.length - 1],
                    column: [0, Store.flowdata[0].length - 1],
                },
            ];
        } else {
            range = $.extend(true, [], Store.luckysheet_select_save);
        }

        let searchIndexArr = _this.getSearchHits(searchText, range);

        if (searchIndexArr.length == 0) {
            if (isEditMode()) {
                alert(locale_findAndReplace.noReplceTip);
            } else {
                tooltip.info(locale_findAndReplace.noReplceTip, "");
            }

            return;
        }

        let count = null;

        let last = Store.luckysheet_select_save[Store.luckysheet_select_save.length - 1];
        let rf = last.row_focus;
        let cf = last.column_focus;

        for (let i = 0; i < searchIndexArr.length; i++) {
            if (
                searchIndexArr[i].r == rf &&
                searchIndexArr[i].c == cf &&
                String(searchIndexArr[i].sheetIndex != null ? searchIndexArr[i].sheetIndex : Store.currentSheetIndex) == String(Store.currentSheetIndex)
            ) {
                count = i;
                break;
            }
        }

        if (count == null) {
            if (searchIndexArr.length == 0) {
                if (isEditMode()) {
                    alert(locale_findAndReplace.noMatchTip);
                } else {
                    tooltip.info(locale_findAndReplace.noMatchTip, "");
                }

                return;
            } else {
                count = 0;
            }
        }

        //正则表达式匹配
        let regCheck = false;
        if ($("#luckysheet-search-replace #regCheck input[type='checkbox']").is(":checked")) {
            regCheck = true;
        }

        //整词匹配
        let wordCheck = false;
        if ($("#luckysheet-search-replace #wordCheck input[type='checkbox']").is(":checked")) {
            wordCheck = true;
        }

        //区分大小写匹配
        let caseCheck = false;
        if ($("#luckysheet-search-replace #caseCheck input[type='checkbox']").is(":checked")) {
            caseCheck = true;
        }

        let replaceText = $("#luckysheet-search-replace #replaceInput input").val();
        const hit = searchIndexArr[count];
        const sheetIndex = hit.sheetIndex != null ? hit.sheetIndex : Store.currentSheetIndex;
        _this.locateSearchHit(hit.r, hit.c, sheetIndex);

        if (!checkProtectionLocked(hit.r, hit.c, sheetIndex)) {
            return;
        }

        let d = editor.deepCopyFlowData(Store.flowdata);
        if (wordCheck) {
            setcellvalue(hit.r, hit.c, d, replaceText);
        } else {
            let reg;
            if (caseCheck) {
                reg = new RegExp(func_methods.getRegExpStr(searchText), "g");
            } else {
                reg = new RegExp(func_methods.getRegExpStr(searchText), "ig");
            }
            let v = valueShowEs(hit.r, hit.c, d)
                .toString()
                .replace(reg, replaceText);
            setcellvalue(hit.r, hit.c, d, v);
        }

        Store.luckysheet_select_save = [{ row: [hit.r, hit.r], column: [hit.c, hit.c] }];

        if ($("#luckysheet-search-replace #searchAllbox").is(":visible")) {
            $("#luckysheet-search-replace #searchAllbox").hide();
        }

        jfrefreshgrid(d, Store.luckysheet_select_save);
        selectHightlightShow();
        _this.locateSearchHit(hit.r, hit.c, sheetIndex);
    },
    replaceAll: function() {
        let _this = this;

        const _locale = locale();
        const locale_findAndReplace = _locale.findAndReplace;

        if (!Store.allowEdit) {
            tooltip.info(locale_findAndReplace.modeTip, "");
            return;
        }

        let searchText = $("#luckysheet-search-replace #searchInput input").val();
        if (searchText == "" || searchText == null) {
            if (isEditMode()) {
                alert(locale_findAndReplace.searchInputTip);
            } else {
                tooltip.info(locale_findAndReplace.searchInputTip, "");
            }

            return;
        }

        let range;
        if (
            Store.luckysheet_select_save.length == 0 ||
            (Store.luckysheet_select_save.length == 1 &&
                Store.luckysheet_select_save[0].row[0] == Store.luckysheet_select_save[0].row[1] &&
                Store.luckysheet_select_save[0].column[0] == Store.luckysheet_select_save[0].column[1])
        ) {
            range = [
                {
                    row: [0, Store.flowdata.length - 1],
                    column: [0, Store.flowdata[0].length - 1],
                },
            ];
        } else {
            range = $.extend(true, [], Store.luckysheet_select_save);
        }

        let searchIndexArr = _this.getSearchHits(searchText, range);

        if (searchIndexArr.length == 0) {
            if (isEditMode()) {
                alert(locale_findAndReplace.noReplceTip);
            } else {
                tooltip.info(locale_findAndReplace.noReplceTip, "");
            }

            return;
        }

        //正则表达式匹配
        let regCheck = false;
        if ($("#luckysheet-search-replace #regCheck input[type='checkbox']").is(":checked")) {
            regCheck = true;
        }

        //整词匹配
        let wordCheck = false;
        if ($("#luckysheet-search-replace #wordCheck input[type='checkbox']").is(":checked")) {
            wordCheck = true;
        }

        //区分大小写匹配
        let caseCheck = false;
        if ($("#luckysheet-search-replace #caseCheck input[type='checkbox']").is(":checked")) {
            caseCheck = true;
        }

        let replaceText = $("#luckysheet-search-replace #replaceInput input").val();

        let replaceCount = 0;
        const groups = {};
        for (let i = 0; i < searchIndexArr.length; i++) {
            const idx = searchIndexArr[i].sheetIndex != null ? searchIndexArr[i].sheetIndex : Store.currentSheetIndex;
            if (!groups[idx]) {
                groups[idx] = [];
            }
            groups[idx].push(searchIndexArr[i]);
        }

        const groupKeys = Object.keys(groups);
        for (let g = 0; g < groupKeys.length; g++) {
            const sheetIndex = groupKeys[g];
            if (String(sheetIndex) != String(Store.currentSheetIndex)) {
                sheetmanage.changeSheetExec(sheetIndex);
            }

            let d = editor.deepCopyFlowData(Store.flowdata);
            let sheetRange = [];
            for (let i = 0; i < groups[sheetIndex].length; i++) {
                let r = groups[sheetIndex][i].r;
                let c = groups[sheetIndex][i].c;

                if (!checkProtectionLocked(r, c, sheetIndex, false)) {
                    continue;
                }

                if (wordCheck) {
                    setcellvalue(r, c, d, replaceText);
                } else {
                    let reg;
                    if (caseCheck) {
                        reg = new RegExp(func_methods.getRegExpStr(searchText), "g");
                    } else {
                        reg = new RegExp(func_methods.getRegExpStr(searchText), "ig");
                    }
                    let v = valueShowEs(r, c, d)
                        .toString()
                        .replace(reg, replaceText);
                    setcellvalue(r, c, d, v);
                }

                sheetRange.push({ row: [r, r], column: [c, c] });
                replaceCount++;
            }

            if (sheetRange.length) {
                jfrefreshgrid(d, sheetRange);
            }
        }

        if ($("#luckysheet-search-replace #searchAllbox").is(":visible")) {
            $("#luckysheet-search-replace #searchAllbox").hide();
        }

        selectHightlightShow();

        let succeedInfo = replaceHtml(locale_findAndReplace.successTip, {
            xlength: replaceCount,
        });
        if (isEditMode()) {
            alert(succeedInfo);
        } else {
            tooltip.info(succeedInfo, "");
        }
    },
};

export default luckysheetSearchReplace;
