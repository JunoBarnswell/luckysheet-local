# LuckySheet 对外 API 清单

> 对照源码：`frontend/src/global/api.js`（v2.1.13，109 个 `export function` + 行首空格的 `setRangeFormat`）
> 对照文档：`frontend/docs/zh/guide/api.md`、`frontend/docs/guide/api.md`
> 挂载：`frontend/src/core.js:45` `common_extend(api, luckysheet)`；`create`/`destroy` 在 core.js 另行赋值。

## 列含义

| 列 | 判定规则（本函数体，经脚本扫描后人工校对关键路径） |
|---|---|
| 公式 | 是否调用 `luckysheetformula.updatecell` / `formula.delFunctionGroup` / `execFunctionGroup*` |
| 刷新 | 是否调用 `jfrefreshgrid` / `luckysheetrefreshgrid` / `jfrefreshgrid_rhcw` |
| server | 是否直接 `server.saveParam`。`jfrefreshgrid` 在 `allowUpdate` 时还会走 `historyParam`，表中不把间接广播标成「是」 |
| 文档 | 中/英文 `api.md` 是否有同名小节 |
| 间接 | 包装函数体内无上述调用，但会进入被包装函数 |

返回值：成功路径以 JSDoc / `return` 为准；失败普遍 `tooltip.info(...)` 后 `return`。`options.success` 为回调，不是返回值。

## 生命周期（不在 api.js 导出）

| 方法 | 源码 | 签名 | 返回值 | 公式 | 刷新 | server | 文档 | 说明 |
|---|---|---|---|---|---|---|---|---|
| `create` | `core.js:48` | `create(setting)` | void | 是（`functionlist` + `execFunctionGroupForce`） | 是（`initialjfFile`） | 条件（`allowUpdate` 时 `openWebSocket`） | 有 | 首行 `method.destroy()`；写入 Store；无 loadUrl 走 `sheetmanage.initialjfFile` |
| `destroy` | `core.js:246` → `method.js:431` | `destroy()` | void | 重置 formula 模块 | 清空 DOM | 否（不断开逻辑见 `closeWebsocket`） | 有 | 用 defaultStore 覆盖 Store（#529） |

旧版挂载（`core.js:203-243`，文档「旧版 API」）：`getluckysheetfile` `getluckysheet_select_save` `setluckysheet_select_save` `getconfig` `getGridData` `buildGridData` `luckysheetrefreshgrid` `jfrefreshgrid` `getcellvalue` `setcellvalue` `getdatabyselection` `flowdata` `selectHightlightShow`。

## api.js 方法总表

| # | 方法 | 行 | 签名 | 返回值 | 公式 | 刷新 | server | 文档 | 说明 |
|---|---|---:|---|---|---|---|---|---|---|
| 1 | `getCellValue` | 56 | `getCellValue(row, column, options = {})` | 单元格值（`options.type` 默认 `v`，可为 `m`/`f`）；失败 `tooltip.info` | 否 | 否 | 否 | 有 | 获取单元格的值 |
| 2 | `setCellValue` | 115 | `setCellValue(row, column, value, options = {})` | void；`options.success(data)` | 是 | 是 | 否 | 有 | 仅当 value 为对象且 f!=null && v==null 才 updatecell（#1004）；当前 sheet 且 isRefresh 时 jfrefreshgrid。协同经 refresh.historyParam，体内无 saveParam。 |
| 3 | `clearCell` | 270 | `clearCell(row, column, options = {})` | invalid: tooltip.info | 是 | 是 | 否 | 有 | 清除指定工作表指定单元格的内容，返回清除掉的数据，不同于删除单元格的功能，不需要设定单元格移动情况 |
| 4 | `deleteCell` | 324 | `deleteCell(move, row, column, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 删除指定工作表指定单元格，返回删除掉的数据，同时，指定是右侧单元格左移还是下方单元格上移 |
| 5 | `setCellFormat` | 366 | `setCellFormat(row, column, attr, value, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 设置某个单元格的属性，如果要设置单元格的值或者同时设置多个单元格属性，推荐使用setCellValue |
| 6 | `find` | 450 | `find(content, options = {})` | invalid: tooltip.info / result | 否 | 否 | 否 | 有 | 查找一个工作表中的指定内容，返回查找到的内容组成的单元格一位数组，数据格式同celldata |
| 7 | `replace` | 529 | `replace(content, replaceContent, options = {})` | invalid: tooltip.info / celldata[] | 否 | 是 | 否 | 有 | 查找一个工作表中的指定内容并替换成新的内容，返回替换后的内容组成的单元格一位数组，数据格式同celldata。 |
| 8 | `exitEditMode` | 570 | `exitEditMode(options = {})` | void + success() | 是 | 否 | 否 | 有 | 手动触发退出编辑模式 |
| 9 | `enterEditMode` | 608 | `enterEditMode(options = {})` | void + success() | 否 | 否 | 否 | 有 | 手动触发进入编辑模式 |
| 10 | `frozenFirstRow` | 632 | `frozenFirstRow(order)` | void | 否 | 是 | 否 | 无 | 冻结首行 |
| 11 | `frozenFirstColumn` | 685 | `frozenFirstColumn(order)` | void | 否 | 是 | 否 | 无 | 冻结首列 |
| 12 | `frozenRowRange` | 739 | `frozenRowRange(range, order)` | void | 否 | 是 | 否 | 无 | 冻结行选区 |
| 13 | `frozenColumnRange` | 801 | `frozenColumnRange(range, order)` | void | 否 | 是 | 否 | 无 | 冻结列选区 |
| 14 | `cancelFrozen` | 862 | `cancelFrozen(order)` | void + success() | 否 | 是 | 否 | 有 | 取消冻结 |
| 15 | `setHorizontalFrozen` | 886 | `setHorizontalFrozen(isRange, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 frozenFirstRow / frozenRowRange。冻结行操作。特别注意，只有在isRange设置为true的时候，才需要设置setting中的range，且与一般的range格式不同。 |
| 16 | `setVerticalFrozen` | 916 | `setVerticalFrozen(isRange, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 frozenFirstColumn / frozenColumnRange。冻结列操作。特别注意，只有在isRange设置为true的时候，才需要设置setting中的range，且与一般的range格式不同。 |
| 17 | `setBothFrozen` | 946 | `setBothFrozen(isRange, options = {})` | void + success() | 否 | 是 | 否 | 有 | 冻结行列操作。特别注意，只有在isRange设置为true的时候，才需要设置setting中的range，且与一般的range格式不同。 |
| 18 | `insertRowOrColumn` | 1091 | `insertRowOrColumn(type, index = 0, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 无 | 在第index行或列的位置，插入number行或列 |
| 19 | `insertRowBottomOrColumnRight` | 1147 | `insertRowBottomOrColumnRight(type, index = 0, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 无 | 在第index行或列的位置，插入number行或列 |
| 20 | `insertRow` | 1202 | `insertRow(row = 0, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 insertRowOrColumn。在第row行的位置，插入number行空白行 |
| 21 | `insertRowBottom` | 1213 | `insertRowBottom(row = 0, options = {})` | void + success() | 间接 | 间接 | 间接 | 无 | 包装 insertRowBottomOrColumnRight。在第row行的位置，插入number行空白行 |
| 22 | `insertColumn` | 1224 | `insertColumn(column = 0, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 insertRowOrColumn。在第column列的位置，插入number列空白列 |
| 23 | `insertColumnRight` | 1235 | `insertColumnRight(column = 0, options = {})` | void + success() | 间接 | 间接 | 间接 | 无 | 包装 insertRowBottomOrColumnRight。在第column列的位置，插入number列空白列 |
| 24 | `deleteRowOrColumn` | 1247 | `deleteRowOrColumn(type, startIndex, endIndex, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 无 | 删除指定的行或列。删除行列之后，行列的序号并不会变化，下面的行（右侧的列）会补充到上（左）面，注意观察数据是否被正确删除即可。 |
| 25 | `deleteRow` | 1280 | `deleteRow(rowStart, rowEnd, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 deleteRowOrColumn。删除指定的行。 |
| 26 | `deleteColumn` | 1292 | `deleteColumn(columnStart, columnEnd, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 deleteRowOrColumn。删除指定的列。 |
| 27 | `hideRowOrColumn` | 1305 | `hideRowOrColumn(type, startIndex, endIndex, options = {})` | invalid: tooltip.info | 否 | 是 | 是 | 无 | 隐藏行或列 |
| 28 | `showRowOrColumn` | 1367 | `showRowOrColumn(type, startIndex, endIndex, options = {})` | invalid: tooltip.info | 否 | 是 | 是 | 无 | 显示隐藏的行或列 |
| 29 | `hideRow` | 1428 | `hideRow(startIndex, endIndex, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 hideRowOrColumn。隐藏行 |
| 30 | `showRow` | 1440 | `showRow(startIndex, endIndex, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 showRowOrColumn。显示行 |
| 31 | `hideColumn` | 1452 | `hideColumn(startIndex, endIndex, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 hideRowOrColumn。隐藏列 |
| 32 | `showColumn` | 1464 | `showColumn(startIndex, endIndex, options = {})` | void + success() | 间接 | 间接 | 间接 | 有 | 包装 showRowOrColumn。显示列 |
| 33 | `setRowHeight` | 1476 | `setRowHeight(rowInfo, options = {})` | invalid: tooltip.info | 否 | 是 | 是 | 有 | 设置指定行的高度。优先级最高，高于默认行高和用户自定义行高。 |
| 34 | `setColumnWidth` | 1533 | `setColumnWidth(columnInfo, options = {})` | invalid: tooltip.info | 否 | 是 | 是 | 有 | 设置指定列的宽度 |
| 35 | `getRowHeight` | 1590 | `getRowHeight(rowInfo, options = {})` | invalid: tooltip.info / rowlen map | 否 | 否 | 否 | 有 | 获取指定工作表指定行的高度，得到行号和高度对应关系的对象 |
| 36 | `getColumnWidth` | 1635 | `getColumnWidth(columnInfo, options = {})` | invalid: tooltip.info / columnlen map | 否 | 否 | 否 | 有 | 获取指定工作表指定列的宽度，得到列号和宽度对应关系的对象 |
| 37 | `getDefaultRowHeight` | 1679 | `getDefaultRowHeight(options = {})` | Store field | 否 | 否 | 否 | 有 | 获取工作表的默认行高 |
| 38 | `getDefaultColWidth` | 1702 | `getDefaultColWidth(options = {})` | Store field | 否 | 否 | 否 | 有 | 获取工作表的默认列宽 |
| 39 | `getRange` | 1724 | `getRange()` | `{row,column}[]` | 否 | 否 | 否 | 有 | 返回当前选区对象的数组，可能存在多个选区。 |
| 40 | `getRangeWithFlatten` | 1746 | `getRangeWithFlatten(range)` | result | 否 | 否 | 否 | 有 | 返回表示指定区域内所有单元格位置的数组，区别getRange方法，该方法以cell单元格(而非某块连续的区域)为单位来组织选区的数据 |
| 41 | `getRangeValuesWithFlatte` | 1770 | `getRangeValuesWithFlatte(range)` | void | 否 | 否 | 否 | 有 | 返回表示指定区域内所有单元格内容的对象数组 |
| 42 | `getRangeAxis` | 1788 | `getRangeAxis()` | result | 否 | 否 | 否 | 有 | 返回对应当前选区的坐标字符串数组，可能存在多个选区。 |
| 43 | `getRangeValue` | 1807 | `getRangeValue(options = {})` | void | 否 | 否 | 否 | 有 | 返回指定工作表指定范围的单元格二维数组数据，每个单元格为一个对象 |
| 44 | `getRangeHtml` | 1834 | `getRangeHtml(options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 复制指定工作表指定单元格区域的数据，返回包含`<table>`html格式的数据，可用于粘贴到excel中保持单元格样式。 |
| 45 | `getRangeArray` | 2354 | `getRangeArray(dimensional, options = {})` | invalid: tooltip.info / data | 否 | 否 | 否 | 有 | 复制指定工作表指定单元格区域的数据，返回一维、二维或者自定义行列数的二维数组的数据。只有在dimensional设置为custom的时候，才需要设置setting中的row和column |
| 46 | `getRangeJson` | 2449 | `getRangeJson(isFirstRowTitle, options = {})` | void | 否 | 否 | 否 | 有 | 复制指定工作表指定单元格区域的数据，返回json格式的数据 |
| 47 | `getRangeDiagonal` | 2538 | `getRangeDiagonal(type, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | / |
| 48 | `getRangeBoolean` | 2646 | `getRangeBoolean(options = {})` | void + success() | 否 | 否 | 否 | 有 | 复制指定工作表指定单元格区域的数据，返回布尔值的数据 |
| 49 | `setRangeShow` | 2732 | `setRangeShow(range, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 指定工作表选中一个或多个选区为选中状态并选择是否高亮，支持多种格式设置。 |
| 50 | `setRangeValue` | 2836 | `setRangeValue(data, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 循环 setCellValue(..., {isRefresh:false}) 后把 file.data 回滚到循环前深拷贝（2870-2882），再 jfrefreshgrid(..., isRunExecFunction 默认路径被 true,false 覆盖）。#1004 根因。 |
| 51 | `setSingleRangeFormat` | 2908 | `setSingleRangeFormat(attr, value, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 无 | 设置指定范围的单元格格式，一般用作处理格式，赋值操作推荐使用setRangeValue方法 |
| 52 | `setRangeFilter` | 3032 | `setRangeFilter(type, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 为指定索引的工作表，选定的范围开启或关闭筛选功能 |
| 53 | `setRangeMerge` | 3104 | `setRangeMerge(type, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 为指定索引的工作表，选定的范围设定合并单元格 |
| 54 | `cancelRangeMerge` | 3320 | `cancelRangeMerge(options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 为指定索引的工作表，选定的范围取消合并单元格 |
| 55 | `setRangeSort` | 3460 | `setRangeSort(type, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 为指定索引的工作表，选定的范围开启排序功能，返回选定范围排序后的数据。 |
| 56 | `setRangeSortMulti` | 3555 | `setRangeSortMulti(hasTitle, sort, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 为指定索引的工作表，选定的范围开启多列自定义排序功能，返回选定范围排序后的数据。 |
| 57 | `setRangeConditionalFormatDefault` | 3662 | `setRangeConditionalFormatDefault(conditionName, conditionValue, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 为指定索引的工作表，选定的范围开启条件格式，根据设置的条件格式规则突出显示部分单元格，返回开启条件格式后的数据。 |
| 58 | `setRangeConditionalFormat` | 3955 | `setRangeConditionalFormat(type, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 为指定索引的工作表，选定的范围开启条件格式，返回开启条件格式后的数据。 |
| 59 | `deleteRangeConditionalFormat` | 4224 | `deleteRangeConditionalFormat(itemIndex, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 为指定下标的工作表，删除条件格式规则，返回被删除的条件格式规则 |
| 60 | `clearRange` | 4288 | `clearRange(options = {})` | invalid: tooltip.info | 是 | 是 | 否 | 有 | 清除指定工作表指定单元格区域的内容，不同于删除选区的功能，不需要设定单元格移动情况 |
| 61 | `deleteRange` | 4407 | `deleteRange(move, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 删除指定工作表指定单元格区域，返回删除掉的数据，同时，指定是右侧单元格左移还是下方单元格上移 |
| 62 | `matrixOperation` | 4467 | `matrixOperation(type, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 指定工作表指定单元格区域的数据进行矩阵操作，返回操作成功后的结果数据 |
| 63 | `matrixCalculation` | 4787 | `matrixCalculation(type, number, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 指定工作表指定单元格区域的数据进行矩阵计算，返回计算成功后的结果数据 |
| 64 | `setSheetAdd` | 4892 | `setSheetAdd(options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 新增一个sheet，返回新增的工作表对象 |
| 65 | `setSheetDelete` | 5015 | `setSheetDelete(options = {})` | invalid: tooltip.info / sheet file | 否 | 否 | 否 | 有 | 删除指定下标的工作表，返回已删除的工作表对象 |
| 66 | `setSheetCopy` | 5050 | `setSheetCopy(options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 复制指定下标的工作表到指定下标位置 |
| 67 | `setSheetHide` | 5143 | `setSheetHide(options = {})` | invalid: tooltip.info / sheet file | 否 | 否 | 否 | 有 | 隐藏指定下标的工作表，返回被隐藏的工作表对象 |
| 68 | `setSheetShow` | 5173 | `setSheetShow(options = {})` | invalid: tooltip.info / sheet file | 否 | 否 | 否 | 有 | 取消隐藏指定下标的工作表，返回被取消隐藏的工作表对象 |
| 69 | `setSheetActive` | 5203 | `setSheetActive(order, options = {})` | invalid: tooltip.info / sheet file | 否 | 否 | 否 | 有 | 设置指定下标的工作表为当前工作表（激活态），即切换到指定的工作表，返回被激活的工作表对象 |
| 70 | `setSheetName` | 5236 | `setSheetName(name, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 修改工作表名称 |
| 71 | `setSheetColor` | 5284 | `setSheetColor(color, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 设置工作表名称处的颜色 |
| 72 | `setSheetMove` | 5333 | `setSheetMove(type, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 指定工作表向左边或右边移动一个位置，或者指定索引，返回指定的工作表对象 |
| 73 | `setSheetOrder` | 5425 | `setSheetOrder(orderList, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 重新排序所有工作表的位置，指定工作表顺序的数组。 |
| 74 | `setSheetZoom` | 5484 | `setSheetZoom(zoom, options = {})` | invalid: tooltip.info | 否 | 否 | 是 | 有 | 设置工作表缩放比例 |
| 75 | `showGridLines` | 5528 | `showGridLines(options = {})` | invalid: tooltip.info / sheet file | 否 | 是 | 否 | 有 | 显示指定下标工作表的网格线，返回操作的工作表对象 |
| 76 | `hideGridLines` | 5566 | `hideGridLines(options = {})` | invalid: tooltip.info / sheet file | 否 | 是 | 否 | 有 | 隐藏指定下标工作表的网格线，返回操作的工作表对象 |
| 77 | `refresh` | 5603 | `refresh(options = {})` | void + success() | 否 | 是 | 否 | 有 | 刷新canvas |
| 78 | `scroll` | 5626 | `scroll(options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 滚动当前工作表位置 |
| 79 | `resize` | 5683 | `resize(options = {})` | void + success() | 否 | 否 | 否 | 有 | 根据窗口大小自动resize画布 |
| 80 | `getScreenshot` | 5701 | `getScreenshot(options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 返回指定选区截图后生成的base64格式的图片 |
| 81 | `setWorkbookName` | 5805 | `setWorkbookName(name, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 设置工作簿名称 |
| 82 | `getWorkbookName` | 5827 | `getWorkbookName(options = {})` | void + success() | 否 | 否 | 否 | 有 | 获取工作簿名称 |
| 83 | `undo` | 5860 | `undo(options = {})` | 刚撤销的 history 对象（来自 `jfredo`） | 否 | 间接 | 间接 | 有 | 实际调用 controlHistory.redo；读 Store.jfredo。命名与直觉相反，禁止改名。 |
| 84 | `redo` | 5885 | `redo(options = {})` | 刚重做的 history 对象（来自 `jfundo`） | 否 | 间接 | 间接 | 有 | 实际调用 controlHistory.undo；读 Store.jfundo。 |
| 85 | `getAllSheets` | 5908 | `getAllSheets()` | sheet 数组（含 celldata） | 否 | 否 | 否 | 有 | 把 data 转回 celldata 并删除 load/freezen。 |
| 86 | `getAllChartsBase64` | 5923 | `getAllChartsBase64(cb)` | void | 否 | 否 | 否 | 无 | 返回所有工作表配置 |
| 87 | `getSheet` | 5971 | `getSheet(options = {})` | Store field / sheetmanage result | 否 | 否 | 否 | 有 | 根据index获取sheet页配置 |
| 88 | `getSheetData` | 5996 | `getSheetData(options = {})` | invalid: tooltip.info / data | 否 | 否 | 否 | 有 | 快捷返回指定工作表的数据 |
| 89 | `getConfig` | 6021 | `getConfig(options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 快捷返回指定工作表的config配置 |
| 90 | `setConfig` | 6043 | `setConfig(cfg, options = {})` | invalid: tooltip.info | 否 | 是 | 否 | 有 | 快捷设置指定工作表config配置 |
| 91 | `getLuckysheetfile` | 6081 | `getLuckysheetfile()` | void + success() | 否 | 否 | 否 | 有 | 返回所有表格数据结构的一维数组luckysheetfile |
| 92 | `setDataVerification` | 6103 | `setDataVerification(optionItem, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 指定工作表范围设置数据验证功能，并设置参数 |
| 93 | `deleteDataVerification` | 6328 | `deleteDataVerification(options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 指定工作表范围删除数据验证功能 |
| 94 | `insertImage` | 6393 | `insertImage(src, options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 在指定的工作表中指定单元格位置插入图片 |
| 95 | `deleteImage` | 6590 | `deleteImage(options = {})` | invalid: tooltip.info | 否 | 否 | 否 | 有 | 删除指定工作表中的图片 |
| 96 | `getImageOption` | 6648 | `getImageOption(options = {})` | invalid: tooltip.info / sheet file | 否 | 否 | 否 | 有 | 获取指定工作表的图片配置 |
| 97 | `transToCellData` | 6677 | `transToCellData(data, options = {})` | sheetmanage result | 否 | 否 | 否 | 有 | data => celldata ，data二维数组数据转化成 {r, c, v}格式一维数组 |
| 98 | `transToData` | 6699 | `transToData(celldata, options = {})` | sheetmanage result | 否 | 否 | 否 | 有 | celldata => data ，celldata一维数组数据转化成表格所需二维数组 |
| 99 | `toJson` | 6719 | `toJson()` | 可再传入 `create` 的 options 对象 | 否 | 否 | 否 | 有 | 返回 Store.toJsonOptions + getAllSheets()，无刷新。 |
| 100 | `changLang` | 6747 | `changLang(lang = 'zh')` | invalid: tooltip.info | 否 | 否 | 否 | 有 | toJson() 后 luckysheet.create，会 destroy 当前簿。 |
| 101 | `closeWebsocket` | 6761 | `closeWebsocket()` | void | 否 | 否 | 否 | 有 | 关闭协同连接，无公式/刷新。 |
| 102 | `getRangeByTxt` | 6773 | `getRangeByTxt(txt)` | void | 否 | 否 | 否 | 有 | 根据范围字符串转换为range数组 |
| 103 | `getTxtByRange` | 6796 | `getTxtByRange(range=Store.luckysheet_select_save)` | void | 否 | 否 | 否 | 有 | 根据范围数组转换为范围字符串 |
| 104 | `pagerInit` | 6813 | `pagerInit(config)` | void | 否 | 否 | 否 | 有 | 初始化分页器 |
| 105 | `refreshFormula` | 6839 | `refreshFormula(success)` | void + success() | 是 | 是 | 否 | 有 | formula.execFunctionGroupForce(true) + luckysheetrefreshgrid。 |
| 106 | `updataSheet` | 6856 | `updataSheet(options = {})` | void + success() | 是 | 是 | 是 | 有 | 可 forceCalculation；server.saveParam('shs')。 |
| 107 | `refreshMenuButtonFocus` | 6903 | `refreshMenuButtonFocus(data ,r,c , success)` | void + success() | 否 | 否 | 否 | 有 | 刷新状态栏的状态 |
| 108 | `checkTheStatusOfTheSelectedCells` | 6927 | `checkTheStatusOfTheSelectedCells(type,status)` | void | 否 | 否 | 否 | 有 | 检查选区内所有cell指定类型的状态是否满足条件（主要是粗体、斜体、删除线和下划线等等） |
| 109 | `openSearchDialog` | 6947 | `openSearchDialog(source = 1)` | void | 否 | 否 | 否 | 有 | 调用查找/替换 dialog |
| 110 | `setRangeFormat` | 2961 | `setRangeFormat(attr, value, options = {})` | invalid: tooltip.info / void + success() | 否 | 是 | 否 | 有 | 行首有空格的 export；多选区循环 `setSingleRangeFormat` 后 `jfrefreshgrid` |

## 与文档的差异

### 文档有、api.js 无独立导出

| 文档方法 | 现状 |
|---|---|
| `insertRange(move)` | `zh/guide/api.md` 有小节；`api.js` **无** `export function insertRange` |
| `insertChart` / `setChart` / `getChart` / `deleteChart` | 文档「图表」章；api.js 无。图表走 `expendPlugins/chart` + `method.editorChart` / `insertChartTosheet` |
| `setProtection` | 文档「工作表保护」；api.js 无。保护在 `config.authority` + protection 事件命名空间 |
| `enterEditMode` | 源码有（`api.js:608`），英文/中文 api.md **单元格章未列**（仅有 exitEditMode） |
| `frozenFirstRow` 等内部冻结 | 源码导出；文档只写 `setHorizontalFrozen` 包装 |

### 源码有、文档无或标题不一致

| 源码方法 | 说明 |
|---|---|
| `frozenFirstRow` `frozenFirstColumn` `frozenRowRange` `frozenColumnRange` | 冻结实现函数 |
| `insertRowOrColumn` `insertRowBottomOrColumnRight` `insertRowBottom` `insertColumnRight` | 行列插入实现/方向变体 |
| `deleteRowOrColumn` `hideRowOrColumn` `showRowOrColumn` | 行列显隐实现 |
| `setSingleRangeFormat` | `setRangeFormat` 的单区内核 |
| `getAllChartsBase64` | 依赖页面上的 echarts 实例 |
| `getLuckysheetfile` | 文档有；与旧版 `getluckysheetfile` 并存 |

### 关键协同/公式路径（人工核对）

| 方法 | 公式 | 刷新 | server.saveParam |
|---|---|---|---|
| `setCellValue` | 条件触发 `updatecell`（`api.js:189-194`）或字符串以 `=` 开头（`228-229`） | 当前 sheet 且 `isRefresh`：`jfrefreshgrid`（250-251） | 否（经 refresh） |
| `setRangeValue` | 经 `setCellValue`，但 2881-2882 **回滚 file.data** | `jfrefreshgrid(..., true, false)` + 可选 `luckysheetrefreshgrid` | 否 |
| `clearCell` / `clearRange` | `delFunctionGroup` | `jfrefreshgrid` | 否 |
| `hideRowOrColumn` / `showRowOrColumn` | 否 | `jfrefreshgrid_rhcw` | `cg` + `rowhidden/colhidden` |
| `setRowHeight` / `setColumnWidth` | 否 | `jfrefreshgrid_rhcw` | `cg` + `rowlen/columnlen` |
| `setRangeConditionalFormat*` / `deleteRangeConditionalFormat` | 否 | 条件格式重算+绘制 | `all` + `k:luckysheet_conditionformat_save` |
| `setSheetAdd` | 改 `calcChain[].index` | sheet 切换绘制 | `sha` / `shr` |
| `setSheetName` / `setSheetColor` / `setSheetZoom` | 否 | 部分 | `all` + 对应 `k` |
| `setSheetMove` / `setSheetOrder` | 否 | sheetBar | `shr` |
| `refreshFormula` | `execFunctionGroupForce(true)` | `luckysheetrefreshgrid` | 否 |
| `updataSheet` | 可选 `execFunctionGroupForce` | `luckysheetrefreshgrid` | `shs` |
| `undo` / `redo` | 经 history 恢复 | 经 controlHistory | 经 history |
| `create` | `functionlist` + 初始化重算 | 全量 | `allowUpdate` 开 WS |
| `destroy` | 重置 formula 模块字段 | 清空容器 | 否 |

## 验收说明

本清单覆盖 `api.js` 全部 `export function`（含 `setRangeFormat`）以及 core 的 `create`/`destroy`。后续 Phase 1 改 `setCellValue`/`setRangeValue` 时以本表「公式/刷新」列为回归对照。
