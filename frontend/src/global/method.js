import server from '../controllers/server';
import { luckysheetlodingHTML, luckyColor } from '../controllers/constant';
import sheetmanage from '../controllers/sheetmanage';
import luckysheetformula from './formula';
import imageCtrl from '../controllers/imageCtrl';
import dataVerificationCtrl from '../controllers/dataVerificationCtrl';
import pivotTable from '../controllers/pivotTable';
import luckysheetFreezen from '../controllers/freezen';
import { getSheetIndex } from '../methods/get';
import { luckysheetextendData } from './extend';
import luckysheetConfigsetting from '../controllers/luckysheetConfigsetting';
import editor from './editor';
import luckysheetcreatesheet from './createsheet';
import Store from '../store';
import defaultConfig from '../store/defaults';
import { getFocusedId, getUnit, listUnits, disposeUnit, getFocusedContext } from '../store/registry';
import { focusUnit } from '../store/registry';

export { defaultConfig };

const method = {
    //翻页
    addDataAjax: function(param, index, url, func){
        let _this = this;

        if(index == null){
            index = Store.currentSheetIndex;
        }

        if(url == null){
            url = server.loadSheetUrl;
        }

        $("#luckysheet-grid-window-1").append(luckysheetlodingHTML());
        param.currentPage++;
        
        let dataType = 'application/json;charset=UTF-8';
        let token = sessionStorage.getItem('x-auth-token');

        $.ajax({
            method: 'POST',
            url: url,
            headers: { "x-auth-token": token },
            data: JSON.stringify(param),
            contentType: dataType,
            success: function(d) {
                //d可能为json字符串
                if(typeof d == "string"){
                    d = JSON.parse(d);
                }

                let dataset = d.data;
                
                let newData = dataset.celldata;
                luckysheetextendData(dataset["row"], newData);

                setTimeout(function(){
                    Store.loadingObj.close()
                }, 500);

                if(func && typeof(func)=="function"){ 
                    func(dataset);
                }
            }
        })
    },
    //重载
    reload: function(param, index, url, func){
        let _this = this;

        if(index == null){
            index = Store.currentSheetIndex;
        }

        if(url == null){
            url = server.loadSheetUrl;
        }

        $("#luckysheet-grid-window-1").append(luckysheetlodingHTML());

        let arg = {"gridKey" : server.gridKey, "index": index};
        param = $.extend(true, param, arg);
        let file = Store.luckysheetfile[getSheetIndex(index)];

        $.post(url, param, function (d) {
            let dataset = new Function("return " + d)();
            file.celldata = dataset[index.toString()];
            let data = sheetmanage.buildGridData(file);

            setTimeout(function(){
                Store.loadingObj.close()
            }, 500);

            file["data"] = data;
            Store.flowdata = data;
            editor.webWorkerFlowDataCache(data);//worker存数据

            luckysheetcreatesheet(data[0].length, data.length, data, null, false);
            file["load"] = "1";

            Store.luckysheet_select_save.length = 0;
            Store.luckysheet_selection_range = [];

            server.saveParam("shs", null, Store.currentSheetIndex);

            sheetmanage.changeSheet(index);

            if(func && typeof(func)=="function"){ 
                func();
            }
        });
    },
    clearSheetByIndex: function(i){
        let index = getSheetIndex(i);
        let sheetfile = Store.luckysheetfile[index];

        if(!sheetfile.isPivotTable){
            sheetfile.data = [];
            sheetfile.row = Store.defaultrowNum;
            sheetfile.column = Store.defaultcolumnNum;

            sheetfile.chart = [];
            sheetfile.config = null;
            sheetfile.filter = null;
            sheetfile.filter_select = null;
            sheetfile.celldata = [];
            sheetfile.pivotTable = {};
            sheetfile.calcChain = [];
            sheetfile.status = 0;
            sheetfile.load = 0;

            Store.flowdata = [];
            editor.webWorkerFlowDataCache(Store.flowdata);//worker存数据

            $("#"+ Store.container +" .luckysheet-data-visualization-chart").remove();
            $("#"+ Store.container +" .luckysheet-datavisual-selection-set").remove();

            $("#luckysheet-row-count-show, #luckysheet-formula-functionrange-select, #luckysheet-row-count-show, #luckysheet-column-count-show, #luckysheet-change-size-line, #luckysheet-cell-selected-focus, #luckysheet-selection-copy, #luckysheet-cell-selected-extend, #luckysheet-cell-selected-move, #luckysheet-cell-selected").hide();

            delete sheetfile.load;
        }
        else {
            delete Store.luckysheetfile[index];
        }
    },
    clear: function(index){
        let _this = this;

        if(index == "all"){
            for(let i = 0; i < Store.luckysheetfile.length; i++){
                let sheetfile = Store.luckysheetfile[i];
                _this.clearSheetByIndex(sheetfile.index);
            }
            
        }
        else{
            if(index == null){
                index = Store.currentSheetIndex;
            }
            _this.clearSheetByIndex(index);
        }

        sheetmanage.changeSheet(Store.luckysheetfile[0].index);
    },
    destroy:function(id){
        const targetId = id != null ? id : getFocusedId();
        if (targetId && getUnit(targetId) && getFocusedId() !== targetId) {
            focusUnit(targetId);
        }

        if (Store.luckysheetfile != null && $("#luckysheet-scrollbar-x").length && $("#luckysheet-scrollbar-y").length) {
            sheetmanage.storeSheetParam(false);
        }

        const container = Store.container;
        const instId = Store.instanceId || targetId;
        if (server.websocket != null) {
            try {
                server.intentionalClose = true;
                server.websocket.onopen = null;
                server.websocket.onmessage = null;
                server.websocket.onerror = null;
                server.websocket.onclose = null;
                if (server.websocket.readyState === 0 || server.websocket.readyState === 1) {
                    server.websocket.close(1000, "instance destroy");
                }
            } catch (e) { /* best-effort socket teardown */ }
            server.websocket = null;
        }
        if (server.retryTimer) {
            clearInterval(server.retryTimer);
            server.retryTimer = null;
        }
        if (server.reconnectTimer) {
            clearTimeout(server.reconnectTimer);
            server.reconnectTimer = null;
        }

        if (container) {
            $("#" + container).empty();
        }

        if (instId) {
            $('[data-ls-instance="' + instId + '"]').remove();
        } else {
            $("body > .luckysheet-cols-menu").remove();
            $("#luckysheet-modal-dialog-mask, #luckysheetTextSizeTest, #luckysheet-icon-morebtn-div").remove();
            $("#luckysheet-input-box").parent().remove();
            $("#luckysheet-formula-help-c").remove();
            $(".chartSetting, .luckysheet-modal-dialog-slider").remove();
        }

        const targetContext = targetId ? getUnit(targetId) : null;
        if (targetContext && targetContext.portals) {
            targetContext.portals.forEach(function (node) {
                if (node && node.parentNode) {
                    node.parentNode.removeChild(node);
                }
            });
            targetContext.portals.clear();
        }

        if (targetId) {
            disposeUnit(targetId);
        }

        const remaining = listUnits();
        if (remaining.length > 0) {
            focusUnit(remaining[remaining.length - 1]);
        }
    },
    destroyAll:function(){
        const ids = listUnits().slice();
        if (!ids.length) {
            this.destroy();
            return;
        }
        for (let i = 0; i < ids.length; i++) {
            this.destroy(ids[i]);
        }
    },
    editorChart:function(c){
        let chart_selection_color = luckyColor[0];
        let chart_id = "luckysheetEditMode-datav-chart";
        let chart_selection_id = chart_id + "_selection";
        c.chart_id = chart_id;
        let chartTheme = c.chartTheme;
        chartTheme = chartTheme == null ? "default0000" : chartTheme;

        luckysheet.insertChartTosheet(c.sheetIndex, c.dataSheetIndex, c.option, c.chartType, c.selfOption, c.defaultOption, c.row, c.column, chart_selection_color, chart_id, chart_selection_id, c.chartStyle, c.rangeConfigCheck, c.rangeRowCheck, c.rangeColCheck, c.chartMarkConfig, c.chartTitleConfig, c.winWidth, c.winHeight, c.scrollLeft, c.scrollTop, chartTheme, c.myWidth, c.myHeight, c.myLeft!=null?parseFloat(c.myLeft):null, c.myTop!=null?parseFloat(c.myTop):null, c.myindexrank, true);

        $("#"+chart_id).find(".luckysheet-modal-controll-update").click();
    },
    /**
     * 获取单元格的值
     * @param {name} 函数名称
     * @param {arguments} 函数参数
     */
    createHookFunction:function(){
        let hookName = arguments[0];
        if(luckysheetConfigsetting.hook && luckysheetConfigsetting.hook[hookName]!=null && (typeof luckysheetConfigsetting.hook[hookName] == "function")){
            var args = Array.prototype.slice.apply(arguments);
            args.shift();
            let ret = luckysheetConfigsetting.hook[hookName].apply(this, args);
            if(ret===false){
                return false;
            }
            else{
                return true;
            }
        }

        return true;
    }

}

export default method;
