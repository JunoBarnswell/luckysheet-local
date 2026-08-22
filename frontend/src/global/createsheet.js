import editor from './editor';
import rhchInit from './rhchInit';
import formula from './formula';
import { luckysheetrefreshgrid } from './refresh';
import sheetmanage from '../controllers/sheetmanage';
import Store from '../store';
import { asSparseGrid, ensureSparseSize, estimateViewportRect, growSparseToRect } from './sparseGrid';

export default function luckysheetcreatesheet(colwidth, rowheight, data, cfg, active) {
    if(active == null){
        active = true;
    }

    Store.visibledatarow = [];
    Store.visibledatacolumn = [];
    Store.ch_width = 0;
    Store.rh_height = 0;
    Store.zoomRatio = 1;
    Store.drawDirtyRect = null;
    Store.drawViewportRect = null;

    if(cfg != null){
        Store.config = cfg;
    }
    else{
        Store.config = {};
    }

    let grid = asSparseGrid(data, rowheight, colwidth);
    ensureSparseSize(grid, Math.max(rowheight, 1) - 1, Math.max(colwidth, 1) - 1);
    growSparseToRect(grid, estimateViewportRect(rowheight, colwidth, Store));
    Store.flowdata = grid;

    editor.webWorkerFlowDataCache(Store.flowdata);//worker存数据

    rhchInit(rowheight, colwidth);

    if(active){
        sheetmanage.showSheet();

        setTimeout(function () {
            sheetmanage.restoreCache();
            formula.execFunctionGroup();
            sheetmanage.restoreSheetAll(Store.currentSheetIndex);
            luckysheetrefreshgrid();
        }, 1);
    }
}
