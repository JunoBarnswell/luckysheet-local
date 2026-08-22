# OPEN [BUG] Issue → 本地模块映射

> 出处：[GitHub缺陷深挖V2](64bb840c-9956-4e98-b843-3428a3534c7e)。  
> Issue 正文来自 GitHub 公开页；**文件路径与行号已在本仓库 2026-08-22 复核**。  
> 官方仓 `dream-num/Luckysheet` 已 archived（2025-10-30），下列 OPEN 不会在 upstream 关闭。

## 必含缺陷（复现步骤）

### #1004 — setRangeValue 公式单元格不联动（P0）

- URL：https://github.com/dream-num/Luckysheet/issues/1004
- 状态：OPEN
- 本地：`frontend/src/global/api.js:189-214`、`api.js:2836-2897`、`frontend/src/global/setdata.js:23-27`、`frontend/src/global/refresh.js:25-36`
- 复现：
  1. `luckysheet.create` 一张表，A1 放入数值 `1`。
  2. 用 `setRangeValue` 或初始化把 B1 写成 `{f:'=A1*2', v:2, m:'2'}`（**同时带 f 与 v**）。
  3. `luckysheet.setCellValue(0, 0, 10)` 或 `setRangeValue` 改 A1。
  4. **期望** B1 变为 20；**实际** B1 仍显示 2。
- 根因：`setCellValue` 只在 `value.f!=null && value.v==null` 时走 `luckysheetformula.updatecell`；`setRangeValue` 循环写入后还把 `file.data` 回滚到循环前深拷贝（`2880-2882`）。

### #504 — 表格偏下时首次点击滚顶（P0）

- URL：https://github.com/dream-num/Luckysheet/issues/504
- 状态：OPEN
- 本地：`frontend/src/controllers/updateCell.js:98`、`frontend/src/controllers/formulaBar.js:63,71`、`frontend/src/global/formula.js:3487`、`frontend/src/controllers/filter.js:149`
- 复现：
  1. 页面上方放超过一屏的内容，Luckysheet 在视口下方（或 `fullscreenmode:false`）。
  2. 滚动到能看见表格。
  3. **第一次单击/进入编辑**。
  4. **期望** `window.scrollY` 不变；**实际** 整页滚回顶部。
- 根因：编辑入口 `$("#luckysheet-rich-text-editor").focus().select()` 无 `preventScroll`。`util.js:490,507` 仅部分路径修过。

### #794 — 嵌套页单击跳到顶部（P0，upstream CLOSED，本 fork 仍需防回归）

- URL：https://github.com/dream-num/Luckysheet/issues/794
- 状态：upstream CLOSED（Dushusir 在 container focus 加了 preventScroll）
- 本地仍裸 `focus()`：`updateCell.js:98`、`formulaBar.js:63,71`、`formula.js:3487`、`filter.js:149`、`selection.js:38,550,592,618`
- 复现：
  1. iframe 或上方固定工具栏的嵌套页打开 Luckysheet。
  2. 第一次单击单元格。
  3. **期望** 不跳顶、第二次才能编辑的现象消失；**实际** 本 fork 编辑路径仍可能跳顶。
- 对照：`frontend/src/utils/util.js:506-507` 注释写明 `fix #794 #152`，只覆盖 `luckysheetContainerFocus`。

### #529 — destroy/create 后纵向滚动条位置错（P1）

- URL：https://github.com/dream-num/Luckysheet/issues/529
- 状态：OPEN
- 本地：`frontend/src/global/method.js:431-453`、`frontend/src/controllers/sheetmanage.js:940-977`（`execF`）、`sheetmanage.js:1146-1166`（`restoreselect` 再次 `scrollTop`）
- 复现：
  1. create 一张较高的表，把纵向滚动条拉离顶部（或不写 `scrollTop`）。
  2. `luckysheet.destroy()`。
  3. 用**另一份** data 再 `create`。
  4. **期望** 新表滚动从 0 或新 options.`scrollTop` 开始；**实际** 可能停在底部/旧位置。
- 根因：destroy 重置 Store，但 sheet 对象上的 `scrollTop` 与 `execF`/`restoreselect` 双写恢复产生竞争。

## 映射表（≥20 个 OPEN [BUG]）

| Issue | 标题（摘要） | P | 模块 | 本地文件（已核实） | 复现要点 |
|---|---|---|---|---|---|
| [#1004](https://github.com/dream-num/Luckysheet/issues/1004) | setRangeValue 公式不更新 | P0 | 公式 / API | `api.js` `setdata.js` `refresh.js` `formula.js` | 见上 |
| [#504](https://github.com/dream-num/Luckysheet/issues/504) | 首次点击滚顶 | P0 | 交互 / IME | `updateCell.js:98` `formulaBar.js:63,71` | 见上 |
| [#794](https://github.com/dream-num/Luckysheet/issues/794) | 嵌套页单击跳顶 | P0 | 交互 | 同上 + `util.js:506-507` `filter.js:149` | 见上；upstream 已关 |
| [#529](https://github.com/dream-num/Luckysheet/issues/529) | destroy/create 滚动条 | P1 | 生命周期 | `method.js:431-453` `sheetmanage.js:940-977,1146-1166` | 见上 |
| [#479](https://github.com/dream-num/Luckysheet/issues/479) | 跨表引用公式不生效 | P0 | 公式 | `formula.js` `sheetmanage.js` `demoData` 跨 sheet `f` | sheet2 公式带错误 `v` 时引用吃缓存值 |
| [#477](https://github.com/dream-num/Luckysheet/issues/477) | 有数据单元格保护失效 | P0 | 保护 / 协同 | `config.authority`（sheet 文档）protection 事件 `method.js:442` | 空白格锁、有值格仍可编 |
| [#451](https://github.com/dream-num/Luckysheet/issues/451) | 6000 行筛选数据丢失 | P0 | 筛选 | `filter.js`（`rowhidden`）`sheetmanage.js:715-749` | 大表筛选后数据不可恢复 |
| [#483](https://github.com/dream-num/Luckysheet/issues/483) | 初始化/协同 JSON 公式不执行 | P1 | 公式 / 协同 | `sheetmanage.js:948` `core.js:170-181` `server.js:87` | 含公式单元格初始化不重算 |
| [#213](https://github.com/dream-num/Luckysheet/issues/213) | 协同离线更改丢失 | P0 | 协同 | `server.js:25-86` `server.js` WebSocket | 断网指令未达后台 |
| [#214](https://github.com/dream-num/Luckysheet/issues/214) | 切 sheet 不应协同 | P1 | 协同 | `server.js` `sheetmanage.js` 切表；`api.js:6893` `saveParam("shs")` | 切表广播干扰他人 |
| [#1435](https://github.com/dream-num/Luckysheet/issues/1435) | 拖拉复制公式数值错 | P0 | 公式 / 选区 | `handler.js` 填充 `formula.js` `selection.js` | 公式文本未变，计算结果错 |
| [#1420](https://github.com/dream-num/Luckysheet/issues/1420) | setRangeValue 首格公式无格式 | P1 | 公式 / API | `api.js:2836-2897` `setCellValue` | 批量第一格只显示公式 |
| [#181](https://github.com/dream-num/Luckysheet/issues/181) | setCellValue 超行报错 | P1 | API | `api.js:115` `setdata.js:9-15` `createsheet.js` | `setCellValue(100,0,12)` 行未 grow |
| [#412](https://github.com/dream-num/Luckysheet/issues/412) | 公式不支持 `%` | P1 | 公式 | `formula.js` `functionImplementation.js` | `=A1*5%` 当字符串 |
| [#417](https://github.com/dream-num/Luckysheet/issues/417) | getRangeHtml 错位 | P1 | 导出 / 渲染 | `api.js:1834` `getRangeHtml`（合并+隐藏列） | 复制 HTML 行错位 |
| [#1436](https://github.com/dream-num/Luckysheet/issues/1436) | 公式栏中文显示拼音 | P1 | IME / 编辑 | `formulaBar.js` `formula.js` `updateCell.js` | Win 下公式栏输入中文 |
| [#183](https://github.com/dream-num/Luckysheet/issues/183) | 换行后方向键移格 | P2 | 编辑 | `keyboard.js` `updateCell.js` | 格内换行后上下键跳格 |
| [#425](https://github.com/dream-num/Luckysheet/issues/425) | Firefox 编辑多空白行 | P2 | 编辑 | `updateCell.js` `editor.js` | 仅 Firefox Enter 进编辑 |
| [#459](https://github.com/dream-num/Luckysheet/issues/459) | 换行单元格改字号出错 | P2 | 渲染 | `format.js` `draw.js` `inlineString` | 换行+`=` 后字体变小 |
| [#453](https://github.com/dream-num/Luckysheet/issues/453) | 筛选+冻结调行高按钮错位 | P2 | 筛选 / 冻结 | `filter.js` `freezen.js` `getRowlen` | 筛选项位置错 |
| [#151](https://github.com/dream-num/Luckysheet/issues/151) | 冻结行系列错位 | P2 | 冻结 / 滚动 | `freezen.js` `scroll.js` | 冻行后调行高/复制框 |
| [#127](https://github.com/dream-num/Luckysheet/issues/127) | 生成图表报错 | P2 | 图表 | `expendPlugins/chart/plugin.js:26-34`（unpkg CDN） | 未开 `plugins:['chart']` 或 CDN 失败 |
| [#1422](https://github.com/dream-num/Luckysheet/issues/1422) | 图表拖动选区错 | P2 | 图表 | `chart/plugin.js` `store/index.js:91-113` `chartparam` | 无行列标题时选区错乱 |
| [#178](https://github.com/dream-num/Luckysheet/issues/178) | inline 改纯文本内容消失 | P2 | 编辑 | `setdata.js` `updateCell.js` 纯文本 `ct` | A6 设 Plain Text 后空 |
| [#521](https://github.com/dream-num/Luckysheet/issues/521) | 局部嵌入菜单 width 不准 | P2 | 布局 | `controllers/resize.js` `createdom.js` `store` 尺寸字段 | 非全屏菜单宽度 |
| [#185](https://github.com/dream-num/Luckysheet/issues/185) | label:BUG OPEN（目录收录） | P2 | 待对照正文 | 先按标题再打开 GitHub 正文定位 | V2 统计列入 23 条 label:BUG |
| [#165](https://github.com/dream-num/Luckysheet/issues/165) | label:BUG OPEN | P2 | 待对照正文 | 同上 | 同上 |
| [#145](https://github.com/dream-num/Luckysheet/issues/145) | label:BUG OPEN | P2 | 待对照正文 | 同上 | 同上 |
| [#124](https://github.com/dream-num/Luckysheet/issues/124) | label:BUG OPEN | P2 | 待对照正文 | 同上 | 同上 |
| [#449](https://github.com/dream-num/Luckysheet/issues/449) | label:BUG OPEN | P2 | 待对照正文 | 同上 | 同上 |

V2 统计：`label:BUG` + `is:open` 页面约 **23** 条。上表已覆盖必含 4 个 + 公式/协同/渲染主干，合计 **28** 行。`#185/#165/#145/#124/#449` 在 V2 目录中有编号、当时未抓全文，修复前需打开 GitHub 补复现，**不编造步骤**。

## 公告（不算功能 bug 队列）

| Issue | 作用 |
|---|---|
| [#799](https://github.com/dream-num/Luckysheet/issues/799) | 停维公告 |
| [#1454](https://github.com/dream-num/Luckysheet/issues/1454) | EOL，推荐 Univer |

## 本 fork 额外 P0（无独立 GitHub 号）

| 项 | 模块 | 本地 | 说明 |
|---|---|---|---|
| SUBTOTAL + 筛选 | 公式 / 筛选 | `functionImplementation.js:2976` `filter.js` 写 `config.rowhidden` | SUBTOTAL 实现 **无** `rowhidden` 读取（全文件 grep 零命中） |

## 结论

P0 开发顺序建议：#1004 公式写入链路 → #504/#794 focus → SUBTOTAL/rowhidden → #479/#1435 计算正确性 → #451/#477/#213 数据丢失与保护。#529 与多实例评估相关，可在单实例生命周期内先修双写 `scrollTop`。
