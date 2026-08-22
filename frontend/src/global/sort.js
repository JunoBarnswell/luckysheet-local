import { getObjType } from '../utils/util';
import { isRealNull, isRealNum, isEditMode } from './validate';
import { isdatetime, diff } from './datecontroll';
import tooltip from './tooltip';
import editor from './editor';
import { rowlenByRange } from './getRowlen';
import { jfrefreshgrid } from './refresh';
import {checkProtectionAuthorityNormal} from '../controllers/protection';
import Store from '../store';
import locale from '../locale/locale';
import numeral from 'numeral';
import { getSheetIndex } from '../methods/get';

function isRowHidden(r, config) {
    config = config || Store.config || {};
    return !!(config.rowhidden && (r in config.rowhidden));
}

function rangesIntersect(a, b) {
    if (!a || !b || !a.row || !b.row || !a.column || !b.column) {
        return false;
    }
    return !(a.row[1] < b.row[0] || a.row[0] > b.row[1] || a.column[1] < b.column[0] || a.column[0] > b.column[1]);
}

function getArrayFormulaRanges(sheetIndex) {
    const ranges = [];
    const file = Store.luckysheetfile[getSheetIndex(sheetIndex != null ? sheetIndex : Store.currentSheetIndex)];
    if (!file) {
        return ranges;
    }

    const dynamicArray = file.dynamicArray;
    if (Array.isArray(dynamicArray)) {
        for (let i = 0; i < dynamicArray.length; i++) {
            const item = dynamicArray[i];
            if (!item) {
                continue;
            }
            let rowlen = 1;
            let collen = 1;
            if (item.data != null) {
                rowlen = item.data.length || 1;
                if (getObjType(item.data[0]) == "array") {
                    collen = item.data[0].length || 1;
                }
            }
            ranges.push({
                row: [item.r, item.r + rowlen - 1],
                column: [item.c, item.c + collen - 1],
            });
        }
    }

    const data = file.data || Store.flowdata;
    if (data) {
        for (let r = 0; r < data.length; r++) {
            if (!data[r]) {
                continue;
            }
            for (let c = 0; c < data[r].length; c++) {
                const cell = data[r][c];
                if (cell == null || cell.f == null) {
                    continue;
                }
                const f = String(cell.f);
                if (f.charAt(0) === "{" || /^{=/.test(f)) {
                    ranges.push({ row: [r, r], column: [c, c] });
                }
            }
        }
    }

    return ranges;
}

function sortRangeIntersectsArrayFormula(str, edr, c1, c2, sheetIndex) {
    const target = { row: [str, edr], column: [c1, c2] };
    const ranges = getArrayFormulaRanges(sheetIndex);
    for (let i = 0; i < ranges.length; i++) {
        if (rangesIntersect(target, ranges[i])) {
            return true;
        }
    }
    return false;
}

function collectVisibleSortRows(d, str, edr, c1, c2, config) {
    const visible = [];
    const visibleIndex = [];
    let hasMc = false;

    for (let r = str; r <= edr; r++) {
        if (isRowHidden(r, config)) {
            continue;
        }
        const data_row = [];
        for (let c = c1; c <= c2; c++) {
            if (d[r][c] != null && d[r][c].mc != null) {
                hasMc = true;
                break;
            }
            data_row.push(d[r][c]);
        }
        visible.push(data_row);
        visibleIndex.push(r);
    }

    return { visible, visibleIndex, hasMc };
}

function writeVisibleSortRows(d, visibleIndex, c1, c2, sorted) {
    for (let i = 0; i < visibleIndex.length; i++) {
        const r = visibleIndex[i];
        for (let c = c1; c <= c2; c++) {
            d[r][c] = sorted[i][c - c1];
        }
    }
}

function abortSortIfFormulaIntersect(str, edr, c1, c2) {
    const locale_sort = locale().sort;
    if (sortRangeIntersectsArrayFormula(str, edr, c1, c2, Store.currentSheetIndex)) {
        const msg = locale_sort.formulaIntersectError || locale_sort.mergeError;
        if (isEditMode()) {
            alert(msg);
        } else {
            tooltip.info(msg, "");
        }
        return true;
    }
    return false;
}

//数据排序方法
function orderbydata(data, index, isAsc) {
    if (isAsc == null) {
        isAsc = true;
    }

    let a = function (x, y) {
        let x1 = x[index] , y1 = y[index];

        if(getObjType(x[index]) == "object"){
            x1 = x[index].v;
        }

        if(getObjType(y[index]) == "object"){
            y1 = y[index].v;
        }

        if(isRealNull(x1)){
            return 1;
        }

        if(isRealNull(y1)){
            return -1;
        }

        if (isdatetime(x1) && isdatetime(y1)) {
            return diff(x1, y1);
        }
        else if (isRealNum(x1) && isRealNum(y1)) {
            return numeral(x1).value() - numeral(y1).value();
        }
        else if (!isRealNum(x1) && !isRealNum(y1)) {
            return x1.localeCompare(y1, "zh");
        }
        else if (!isRealNum(x1)) {
            return 1;
        }
        else if (!isRealNum(y1)) {
            return -1;
        }
    }

    let d = function (x, y) {
        let x1 = x[index] , y1 = y[index];

        if(getObjType(x[index]) == "object"){
            x1 = x[index].v;
        }

        if(getObjType(y[index]) == "object"){
            y1 = y[index].v;
        }

        if(isRealNull(x1)){
            return 1;
        }

        if(isRealNull(y1)){
            return -1;
        }

        if (isdatetime(x1) && isdatetime(y1)) {
            return diff(y1, x1);
        }
        else if (isRealNum(x1) && isRealNum(y1)) {
            return numeral(y1).value() - numeral(x1).value();
        }
        else if (!isRealNum(x1) && !isRealNum(y1)) {
            return y1.localeCompare(x1, "zh");
        }
        else if (!isRealNum(x1)) {
            return -1;
        }
        else if (!isRealNum(y1)) {
            return 1;
        }
    }

    if (isAsc) {
        return data.sort(a);
    }
    else {
        return data.sort(d);
    }
}

function orderbydata1D(data, isAsc) {
    if (isAsc == null) {
        isAsc = true;
    }

    let a = function (x, y) {
        let x1 = x, y1 = y;

        if(getObjType(x) == "object"){
            x1 = x.v;
        }

        if(getObjType(y) == "object"){
            y1 = y.v;
        }

        if(x1 == null){
            x1 = "";
        }

        if(y1 == null){
            y1 = "";
        }

        if (isdatetime(x1) && isdatetime(y1)) {
            return diff(x1, y1);
        }
        else if (isRealNum(x1) && isRealNum(y1)) {
            return numeral(x1).value() - numeral(y1).value();
        }
        else if (!isRealNum(x1) && !isRealNum(y1)) {
            return x1.localeCompare(y1, "zh");
        }
        else if (!isRealNum(x1)) {
            return 1;
        }
        else if (!isRealNum(y1)) {
            return -1;
        }
    }

    let d = function (x, y) {
        let x1 = x, y1 = y;

        if(getObjType(x) == "object"){
            x1 = x.v;
        }

        if(getObjType(y) == "object"){
            y1 = y.v;
        }

        if(x1 == null){
            x1 = "";
        }

        if(y1 == null){
            y1 = "";
        }

        if (isdatetime(x1) && isdatetime(y1)) {
            return diff(y1, x1);
        }
        else if (isRealNum(x1) && isRealNum(y1)) {
            return numeral(y1).value() - numeral(x1).value();
        }
        else if (!isRealNum(x1) && !isRealNum(y1)) {
            return y1.localeCompare(x1, "zh");
        }
        else if (!isRealNum(x1)) {
            return -1;
        }
        else if (!isRealNum(y1)) {
            return 1;
        }
    }

    if (isAsc) {
        return data.sort(a);
    }
    else {
        return data.sort(d);
    }
}

//排序选区数据
function sortSelection(isAsc) {
    if(!checkProtectionAuthorityNormal(Store.currentSheetIndex, "sort")){
        return;
    }

    const _locale = locale();
    const locale_sort = _locale.sort;

    if(Store.luckysheet_select_save.length > 1){
        if(isEditMode()){
            alert(locale_sort.noRangeError);
        }
        else{
            tooltip.info(locale_sort.noRangeError, "");
        }

        return;
    }

    if(isAsc == null){
        isAsc = true;
    }

    let d = editor.deepCopyFlowData(Store.flowdata);

    let r1 = Store.luckysheet_select_save[0].row[0], 
        r2 = Store.luckysheet_select_save[0].row[1];
    let c1 = Store.luckysheet_select_save[0].column[0], 
        c2 = Store.luckysheet_select_save[0].column[1];

    let str, edr;

    for(let r = r1; r <= r2; r++){
        if(d[r] != null && d[r][c1] != null){
            let cell = d[r][c1];

            if(cell.mc != null || isRealNull(cell.v)){
                continue;
            }

            if(str == null && /[\u4e00-\u9fa5]+/g.test(cell.v)){
                str = r + 1;
                edr = r + 1;
                continue;
            }
            
            if(str == null){
                str = r;    
            }

            edr = r;
        }
    }

    if(str == null || str > r2){
        return;
    }

    if(abortSortIfFormulaIntersect(str, edr, c1, c2)){
        return;
    }

    // 跳过 config.rowhidden 与手动隐藏行，只在可见行之间重排
    let collected = collectVisibleSortRows(d, str, edr, c1, c2, Store.config);
    let hasMc = collected.hasMc;
    let data = collected.visible;

    if(hasMc){
        if(isEditMode()){
            alert(locale_sort.mergeError);
        }
        else{
            tooltip.info(locale_sort.mergeError, "");
        }

        return;
    }

    data = orderbydata(data, 0, isAsc);
    writeVisibleSortRows(d, collected.visibleIndex, c1, c2, data);

    let allParam = {};
    if(Store.config["rowlen"] != null){
        let cfg = $.extend(true, {}, Store.config);
        cfg = rowlenByRange(d, str, edr, cfg);

        allParam = {
            "cfg": cfg,
            "RowlChange": true
        }
    }

    jfrefreshgrid(d, [{ "row": [str, edr], "column": [c1, c2] }], allParam);
}

//排序一列数据
function sortColumnSeletion(colIndex, isAsc) {
    if(!checkProtectionAuthorityNormal(Store.currentSheetIndex, "sort")){
        return;
    }
    if(isAsc == null){
        isAsc = true;
    }

    const _locale = locale();
    const locale_sort = _locale.sort;

    let d = editor.deepCopyFlowData(Store.flowdata);

    let r1 = 0, r2 = d.length - 1;
    let c1 = 0, c2 = d[0].length - 1;

    let str, edr;

    for(let r = r1; r <= r2; r++){
        if(d[r][colIndex] != null && d[r][colIndex].mc != null){
            continue;
        }

        if(d[r][colIndex] != null && !isRealNull(d[r][colIndex].v) && /[\u4e00-\u9fa5]+/g.test(d[r][colIndex].v) && str == null){
            str = r + 1;
            edr = r + 1;
            continue;
        }

        if(str == null){
            str = r;    
        }

        if(d[r][colIndex] != null && !isRealNull(d[r][colIndex].v)){
            edr = r;
        }
    }

    if(str == null || str > r2){
        return;
    }

    if(abortSortIfFormulaIntersect(str, edr, c1, c2)){
        return;
    }

    // 跳过 config.rowhidden 与手动隐藏行，只在可见行之间重排
    let collected = collectVisibleSortRows(d, str, edr, c1, c2, Store.config);
    let hasMc = collected.hasMc;
    let data = collected.visible;

    if(hasMc){
        if(isEditMode()){
            alert(locale_sort.columnSortMergeError);
        }
        else{
            tooltip.info(locale_sort.columnSortMergeError, "");
        }

        return;
    }

    data = orderbydata(data, colIndex, isAsc);
    writeVisibleSortRows(d, collected.visibleIndex, c1, c2, data);

    let allParam = {};
    if(Store.config["rowlen"] != null){
        let cfg = $.extend(true, {}, Store.config);
        cfg = rowlenByRange(d, str, edr, cfg);

        allParam = {
            "cfg": cfg,
            "RowlChange": true
        }
    }

    jfrefreshgrid(d, [{ "row": [str, edr], "column": [c1, c2] }], allParam);
}

export {
    orderbydata,
    orderbydata1D,
    sortSelection,
    sortColumnSeletion,
    isRowHidden,
    sortRangeIntersectsArrayFormula,
    collectVisibleSortRows,
    writeVisibleSortRows,
    abortSortIfFormulaIntersect,
}