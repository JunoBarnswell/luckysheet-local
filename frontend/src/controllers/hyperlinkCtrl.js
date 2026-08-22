import { replaceHtml } from '../utils/util';
import { getcellvalue } from '../global/getdata';
import { luckysheetrefreshgrid } from '../global/refresh';
import { rowLocation, colLocation, mouseposition } from '../global/location';
import formula from '../global/formula';
import tooltip from '../global/tooltip';
import editor from '../global/editor';
import { modelHTML } from './constant';
import { selectHightlightShow } from './select';
import server from './server';
import sheetmanage from './sheetmanage';
import luckysheetFreezen from './freezen';
import menuButton from './menuButton';
import { getSheetIndex } from '../methods/get';
import locale from '../locale/locale';
import Store from '../store';

function collectDefinedNames() {
    const names = [];
    const pushItem = (item, sheetIndex, sheetName) => {
        if (item == null) {
            return;
        }
        if (typeof item === "string") {
            names.push({ name: item, sheetIndex: sheetIndex, sheetName: sheetName });
            return;
        }
        names.push({
            name: item.name || item.id,
            formulaOrRef: item.formulaOrRef || item.ref || item.formula || item.linkAddress,
            sheetIndex: item.sheetIndex != null ? item.sheetIndex : sheetIndex,
            sheetName: item.sheetName || sheetName,
            row: item.row,
            column: item.column,
        });
    };

    if (Array.isArray(Store.definedNames)) {
        Store.definedNames.forEach((item) => pushItem(item));
    }

    (Store.luckysheetfile || []).forEach((file) => {
        if (!file) {
            return;
        }
        if (Array.isArray(file.definedNames)) {
            file.definedNames.forEach((item) => pushItem(item, file.index, file.name));
        }
        if (Array.isArray(file.table)) {
            file.table.forEach((table) => {
                if (!table || !table.name || !table.range) {
                    return;
                }
                names.push({
                    name: table.name,
                    sheetIndex: file.index,
                    sheetName: file.name,
                    row: table.range.row,
                    column: table.range.column,
                });
            });
        }
    });
    return names;
}

function resolveDefinedName(name) {
    const target = String(name || "").trim();
    if (!target) {
        return null;
    }
    const found = collectDefinedNames().find((item) => item.name && String(item.name).toLowerCase() === target.toLowerCase());
    if (!found) {
        return null;
    }
    if (found.row && found.column) {
        return {
            row: found.row,
            column: found.column,
            sheetIndex: found.sheetIndex != null ? found.sheetIndex : Store.currentSheetIndex,
        };
    }
    const ref = found.formulaOrRef;
    if (ref && formula.iscelldata(ref)) {
        return formula.getcellrange(ref);
    }
    if (ref && found.sheetName && formula.iscelldata(found.sheetName + "!" + ref)) {
        return formula.getcellrange(found.sheetName + "!" + ref);
    }
    return null;
}

function sheetIndexByName(name) {
    const files = Store.luckysheetfile || [];
    for (let i = 0; i < files.length; i++) {
        if (files[i] && files[i].name == name) {
            return files[i].index;
        }
    }
    return null;
}

function navigateToCellRange(cellrange) {
    if (!cellrange || !cellrange.row || !cellrange.column) {
        return false;
    }
    const sheetIndex = cellrange.sheetIndex != null ? cellrange.sheetIndex : Store.currentSheetIndex;
    const range = [{
        row: cellrange.row,
        column: cellrange.column
    }];

    if (sheetIndex != Store.currentSheetIndex) {
        $("#luckysheet-sheet-area div.luckysheet-sheets-item").removeClass("luckysheet-sheets-item-active");
        $("#luckysheet-sheets-item" + sheetIndex).addClass("luckysheet-sheets-item-active");
        sheetmanage.changeSheet(sheetIndex);
    }

    Store.luckysheet_select_save = range;
    selectHightlightShow(true);

    let row_pre = cellrange.row[0] - 1 == -1 ? 0 : Store.visibledatarow[cellrange.row[0] - 1];
    let col_pre = cellrange.column[0] - 1 == -1 ? 0 : Store.visibledatacolumn[cellrange.column[0] - 1];

    $("#luckysheet-scrollbar-x").scrollLeft(col_pre);
    $("#luckysheet-scrollbar-y").scrollTop(row_pre);
    return true;
}

const hyperlinkCtrl = {
    item: {
        linkType: 'external', //链接类型 external外部链接，internal内部链接
        linkAddress: '',  //链接地址 网页地址或工作表单元格引用
        linkTooltip: '',  //提示
    },
    hyperlink: null,
    createDialog: function(){
        let _this = this;

        const _locale = locale();
        const hyperlinkText = _locale.insertLink;
        const toolbarText = _locale.toolbar;
        const buttonText = _locale.button;

        $("#luckysheet-modal-dialog-mask").show();
        $("#luckysheet-insertLink-dialog").remove();

        let sheetListOption = '';
        Store.luckysheetfile.forEach(item => {
            sheetListOption += `<option value="${item.name}">${item.name}</option>`;
        })

        let content =  `<div class="box">
                            <div class="box-item">
                                <label for="luckysheet-insertLink-dialog-linkText">${hyperlinkText.linkText}：</label>
                                <input type="text" id="luckysheet-insertLink-dialog-linkText"/>
                            </div>
                            <div class="box-item">
                                <label for="luckysheet-insertLink-dialog-linkType">${hyperlinkText.linkType}：</label>
                                <select id="luckysheet-insertLink-dialog-linkType">
                                    <option value="external">${hyperlinkText.external}</option>
                                    <option value="internal">${hyperlinkText.internal}</option>
                                    <option value="DEFINE_NAME">${hyperlinkText.defineName || "DEFINE_NAME"}</option>
                                    <option value="RANGE">${hyperlinkText.range || "RANGE"}</option>
                                </select>
                            </div>
                            <div class="show-box show-box-external">
                                <div class="box-item">
                                    <label for="luckysheet-insertLink-dialog-linkAddress">${hyperlinkText.linkAddress}：</label>
                                    <input type="text" id="luckysheet-insertLink-dialog-linkAddress" placeholder="${hyperlinkText.placeholder1}" />
                                </div>
                            </div>
                            <div class="show-box show-box-internal">
                                <div class="box-item">
                                    <label for="luckysheet-insertLink-dialog-linkSheet">${hyperlinkText.linkSheet}：</label>
                                    <select id="luckysheet-insertLink-dialog-linkSheet">
                                        ${sheetListOption}
                                    </select>
                                </div>
                                <div class="box-item">
                                    <label for="luckysheet-insertLink-dialog-linkCell">${hyperlinkText.linkCell}：</label>
                                    <input type="text" id="luckysheet-insertLink-dialog-linkCell" value="A1" placeholder="${hyperlinkText.placeholder2}" />
                                </div>
                            </div>
                            <div class="show-box show-box-DEFINE_NAME" style="display:none;">
                                <div class="box-item">
                                    <label for="luckysheet-insertLink-dialog-linkDefineName">${hyperlinkText.linkDefineName || hyperlinkText.defineName}：</label>
                                    <input type="text" id="luckysheet-insertLink-dialog-linkDefineName" placeholder="${hyperlinkText.placeholderDefineName || ""}" />
                                </div>
                            </div>
                            <div class="show-box show-box-RANGE" style="display:none;">
                                <div class="box-item">
                                    <label for="luckysheet-insertLink-dialog-rangeSheet">${hyperlinkText.linkSheet}：</label>
                                    <select id="luckysheet-insertLink-dialog-rangeSheet">
                                        ${sheetListOption}
                                    </select>
                                </div>
                                <div class="box-item">
                                    <label for="luckysheet-insertLink-dialog-rangeCell">${hyperlinkText.linkCell}：</label>
                                    <input type="text" id="luckysheet-insertLink-dialog-rangeCell" value="A1:C3" placeholder="${hyperlinkText.placeholder2}" />
                                </div>
                            </div>
                            <div class="box-item">
                                <label for="luckysheet-insertLink-dialog-linkTooltip">${hyperlinkText.linkTooltip}：</label>
                                <input type="text" id="luckysheet-insertLink-dialog-linkTooltip" placeholder="${hyperlinkText.placeholder3}" />
                            </div>
                        </div>`;

        $("body").append(replaceHtml(modelHTML, { 
            "id": "luckysheet-insertLink-dialog", 
            "addclass": "luckysheet-insertLink-dialog", 
            "title": toolbarText.insertLink, 
            "content": content, 
            "botton":  `<button id="luckysheet-insertLink-dialog-confirm" class="btn btn-primary">${buttonText.confirm}</button>
                        <button class="btn btn-default luckysheet-model-close-btn">${buttonText.cancel}</button>`, 
            "style": "z-index:100003" 
        }));
        let $t = $("#luckysheet-insertLink-dialog").find(".luckysheet-modal-dialog-content").css("min-width", 350).end(), 
            myh = $t.outerHeight(), 
            myw = $t.outerWidth();
        let winw = $(window).width(), 
            winh = $(window).height();
        let scrollLeft = $(document).scrollLeft(), 
            scrollTop = $(document).scrollTop();
        $("#luckysheet-insertLink-dialog").css({ 
            "left": (winw + scrollLeft - myw) / 2, 
            "top": (winh + scrollTop - myh) / 3 
        }).show();

        _this.dataAllocation();
    },
    init: function (){
        let _this = this;

        const _locale = locale();
        const hyperlinkText = _locale.insertLink;

        //链接类型
        $(document).off("change.linkType").on("change.linkType", "#luckysheet-insertLink-dialog-linkType", function(e){
            let value = this.value;

            $("#luckysheet-insertLink-dialog .show-box").hide();
            $("#luckysheet-insertLink-dialog .show-box-" + value).show();
        })

        //确认按钮
        $(document).off("click.confirm").on("click.confirm", "#luckysheet-insertLink-dialog-confirm", function(e){
            let last = Store.luckysheet_select_save[Store.luckysheet_select_save.length - 1];
            let rowIndex = last.row_focus || last.row[0];
            let colIndex = last.column_focus || last.column[0];

            //文本
            let linkText = $("#luckysheet-insertLink-dialog-linkText").val();

            let linkType = $("#luckysheet-insertLink-dialog-linkType").val();
            let linkAddress = $("#luckysheet-insertLink-dialog-linkAddress").val();
            let linkSheet = $("#luckysheet-insertLink-dialog-linkSheet").val();
            let linkCell = $("#luckysheet-insertLink-dialog-linkCell").val();
            let linkTooltip = $("#luckysheet-insertLink-dialog-linkTooltip").val();
            let payload = null;

            if(linkType == 'external'){
                if(!/^http[s]?:\/\//.test(linkAddress)){
                    linkAddress = 'https://' + linkAddress;
                }

                if(!/^http[s]?:\/\/([\w\-\.]+)+[\w-]*([\w\-\.\/\?%&=]+)?$/ig.test(linkAddress)){
                    tooltip.info('<i class="fa fa-exclamation-triangle"></i>', hyperlinkText.tooltipInfo1);
                    return;
                }
            }
            else if(linkType == 'DEFINE_NAME'){
                let defineName = $("#luckysheet-insertLink-dialog-linkDefineName").val();
                let resolved = resolveDefinedName(defineName);
                if(!resolved){
                    tooltip.info('<i class="fa fa-exclamation-triangle"></i>', hyperlinkText.tooltipInfoDefineName || hyperlinkText.tooltipInfo2);
                    return;
                }
                linkAddress = String(defineName).trim();
                payload = {
                    sheetName: (Store.luckysheetfile[getSheetIndex(resolved.sheetIndex)] || {}).name,
                    row: resolved.row,
                    column: resolved.column,
                    sheetIndex: resolved.sheetIndex,
                };
            }
            else if(linkType == 'RANGE'){
                let rangeSheet = $("#luckysheet-insertLink-dialog-rangeSheet").val();
                let rangeCell = $("#luckysheet-insertLink-dialog-rangeCell").val();
                if(!formula.iscelldata(rangeCell)){
                    tooltip.info('<i class="fa fa-exclamation-triangle"></i>', hyperlinkText.tooltipInfo2);
                    return;
                }
                linkAddress = rangeSheet + "!" + rangeCell;
                let cellrange = formula.getcellrange(linkAddress);
                if(!cellrange){
                    tooltip.info('<i class="fa fa-exclamation-triangle"></i>', hyperlinkText.tooltipInfo2);
                    return;
                }
                payload = {
                    sheetName: rangeSheet,
                    row: cellrange.row,
                    column: cellrange.column,
                    sheetIndex: cellrange.sheetIndex != null ? cellrange.sheetIndex : sheetIndexByName(rangeSheet),
                };
            }
            else{
                if(!formula.iscelldata(linkCell)){
                    tooltip.info('<i class="fa fa-exclamation-triangle"></i>', hyperlinkText.tooltipInfo2);
                    return;
                }

                linkAddress = linkSheet + "!" + linkCell;
            }

            if(linkText == null || linkText.replace(/\s/g, '') == ''){
                linkText = linkAddress;
            }

            let item = {
                linkType: linkType,
                linkAddress: linkAddress,
                linkTooltip: linkTooltip,
            }
            if(payload){
                item.payload = payload;
            }

            let historyHyperlink = $.extend(true, {}, _this.hyperlink);
            let currentHyperlink = $.extend(true, {}, _this.hyperlink);

            currentHyperlink[rowIndex + "_" + colIndex] = item;

            let d = editor.deepCopyFlowData(Store.flowdata);
            let cell = d[rowIndex][colIndex];

            if(cell == null){
                cell = {};
            }

            cell.fc = 'rgb(0, 0, 255)';
            cell.un = 1;
            cell.v = cell.m = linkText;

            d[rowIndex][colIndex] = cell;

            _this.ref(
                historyHyperlink, 
                currentHyperlink, 
                Store.currentSheetIndex, 
                d, 
                [{ row: [rowIndex, rowIndex], column: [colIndex, colIndex] }]
            );

            $("#luckysheet-modal-dialog-mask").hide();
            $("#luckysheet-insertLink-dialog").hide();
        })
    },
    dataAllocation: function(){
        let _this = this;

        let last = Store.luckysheet_select_save[Store.luckysheet_select_save.length - 1];
        let rowIndex = last.row_focus || last.row[0];
        let colIndex = last.column_focus || last.column[0];

        let hyperlink = _this.hyperlink || {};
        let item = hyperlink[rowIndex + "_" + colIndex] || {};

        //文本
        let text = getcellvalue(rowIndex, colIndex, null, 'm');
        $("#luckysheet-insertLink-dialog-linkText").val(text);

        //链接类型
        let linkType = item.linkType || 'external';
        $("#luckysheet-insertLink-dialog-linkType").val(linkType);

        $("#luckysheet-insertLink-dialog .show-box").hide();
        $("#luckysheet-insertLink-dialog .show-box-" + linkType).show();

        //链接地址
        let linkAddress = item.linkAddress || '';

        if(linkType == 'external'){
            $("#luckysheet-insertLink-dialog-linkAddress").val(linkAddress);
        }
        else if(linkType == 'DEFINE_NAME'){
            $("#luckysheet-insertLink-dialog-linkDefineName").val(linkAddress);
        }
        else if(linkType == 'RANGE'){
            let sheettxt = item.payload && item.payload.sheetName;
            let rangetxt = "";
            if(formula.iscelldata(linkAddress)){
                sheettxt = sheettxt || linkAddress.split("!")[0];
                rangetxt = linkAddress.split("!")[1] || "";
            }
            $("#luckysheet-insertLink-dialog-rangeSheet").val(sheettxt || "");
            $("#luckysheet-insertLink-dialog-rangeCell").val(rangetxt || "A1:C3");
        }
        else{
            if(formula.iscelldata(linkAddress)){
                let sheettxt = linkAddress.split("!")[0];
                let rangetxt = linkAddress.split("!")[1];

                $("#luckysheet-insertLink-dialog-linkSheet").val(sheettxt);
                $("#luckysheet-insertLink-dialog-linkCell").val(rangetxt);
            }
        }

        //提示
        let linkTooltip = item.linkTooltip || '';
        $("#luckysheet-insertLink-dialog-linkTooltip").val(linkTooltip);
    },
    cellFocus: function(r, c){
        let _this = this;

        if(_this.hyperlink == null || _this.hyperlink[r + '_' + c] == null){
            return;
        }

        let item = _this.hyperlink[r + '_' + c];

        if(item.linkType == 'external'){
            window.open(item.linkAddress);
            return;
        }

        if(item.linkType == 'DEFINE_NAME'){
            let cellrange = resolveDefinedName(item.linkAddress);
            if(!cellrange && item.payload && item.payload.row && item.payload.column){
                cellrange = {
                    row: item.payload.row,
                    column: item.payload.column,
                    sheetIndex: item.payload.sheetIndex != null ? item.payload.sheetIndex : sheetIndexByName(item.payload.sheetName),
                };
            }
            if(!cellrange){
                return;
            }
            navigateToCellRange(cellrange);
            return;
        }

        if(item.linkType == 'RANGE'){
            let cellrange = null;
            if(item.payload && item.payload.row && item.payload.column){
                cellrange = {
                    row: item.payload.row,
                    column: item.payload.column,
                    sheetIndex: item.payload.sheetIndex != null ? item.payload.sheetIndex : sheetIndexByName(item.payload.sheetName),
                };
            }
            else if(formula.iscelldata(item.linkAddress)){
                cellrange = formula.getcellrange(item.linkAddress);
            }
            if(!cellrange){
                return;
            }
            navigateToCellRange(cellrange);
            return;
        }

        let cellrange = formula.getcellrange(item.linkAddress);
        navigateToCellRange(cellrange);
    },
    overshow: function(event){
        let _this = this;

        $("#luckysheet-hyperlink-overshow").remove();

        if($(event.target).closest("#luckysheet-cell-main").length == 0){
            return;
        }

        let mouse = mouseposition(event.pageX, event.pageY);
        let scrollLeft = $("#luckysheet-cell-main").scrollLeft();
        let scrollTop = $("#luckysheet-cell-main").scrollTop();
        let x = mouse[0] + scrollLeft;
        let y = mouse[1] + scrollTop;

        if(luckysheetFreezen.freezenverticaldata != null && mouse[0] < (luckysheetFreezen.freezenverticaldata[0] - luckysheetFreezen.freezenverticaldata[2])){
            return;
        }

        if(luckysheetFreezen.freezenhorizontaldata != null && mouse[1] < (luckysheetFreezen.freezenhorizontaldata[0] - luckysheetFreezen.freezenhorizontaldata[2])){
            return;
        }

        let row_index = rowLocation(y)[2];
        let col_index = colLocation(x)[2];

        let margeset = menuButton.mergeborer(Store.flowdata, row_index, col_index);
        if(!!margeset){
            row_index = margeset.row[2];
            col_index = margeset.column[2];
        }

        if(_this.hyperlink == null || _this.hyperlink[row_index + "_" + col_index] == null){
            return;
        }

        let item = _this.hyperlink[row_index + "_" + col_index];
        let linkTooltip = item.linkTooltip;

        if(linkTooltip == null || linkTooltip.replace(/\s/g, '') == ''){
            linkTooltip = item.linkAddress;
        }

        let row = Store.visibledatarow[row_index], 
            row_pre = row_index - 1 == -1 ? 0 : Store.visibledatarow[row_index - 1];
        let col = Store.visibledatacolumn[col_index], 
            col_pre = col_index - 1 == -1 ? 0 : Store.visibledatacolumn[col_index - 1];

        if(!!margeset){
            row = margeset.row[1];
            row_pre = margeset.row[0];
            
            col = margeset.column[1];
            col_pre = margeset.column[0];
        }

        let html = `<div id="luckysheet-hyperlink-overshow" style="background:#fff;padding:5px 10px;border:1px solid #000;box-shadow:2px 2px #999;position:absolute;left:${col_pre}px;top:${row + 5}px;z-index:100;">
                        <div>${linkTooltip}</div>
                        <div>单击鼠标可以追踪</div>
                    </div>`;

        $(html).appendTo($("#luckysheet-cell-main"));
    },
    ref: function(historyHyperlink, currentHyperlink, sheetIndex, d, range){
        let _this = this;

        if (Store.clearjfundo) {
            Store.jfundo.length  = 0;

            let redo = {};
            redo["type"] = "updateHyperlink";
            redo["sheetIndex"] = sheetIndex;
            redo["historyHyperlink"] = historyHyperlink;
            redo["currentHyperlink"] = currentHyperlink;
            redo["data"] = Store.flowdata; 
            redo["curData"] = d;
            redo["range"] = range; 
            Store.jfredo.push(redo); 
        }

        _this.hyperlink = currentHyperlink;
        Store.luckysheetfile[getSheetIndex(sheetIndex)].hyperlink = currentHyperlink;

        Store.flowdata = d;
        editor.webWorkerFlowDataCache(Store.flowdata);//worker存数据
        Store.luckysheetfile[getSheetIndex(sheetIndex)].data = Store.flowdata;

        //共享编辑模式
        if(server.allowUpdate){ 
            server.saveParam("all", sheetIndex, currentHyperlink, { "k": "hyperlink" });
            server.historyParam(Store.flowdata, sheetIndex, range[0]);
        }

        setTimeout(function () {
            luckysheetrefreshgrid();
        }, 1);
    }
}

export default hyperlinkCtrl;