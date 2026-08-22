# LuckySheet × Univer 模块证据册

> Phase 0.1。合并 16 个 V2 子代理结论；每条证据为本地（或 Univer 对照仓）**已复核**的 `文件:行号`。  
> LuckySheet 根：`C:\Users\kuo13\Projects\luckysheet-local`  
> Univer 对照：`C:\Users\kuo13\Projects\univer`  
> 复核日：2026-08-22。路径未写前缀时默认 `frontend/src/`。

## 出处（16/16）

| # | 模块 | Agent ID |
|---|---|---|
| 1 | Univer 架构 | [92a7c37f-fbac-4bc8-8f61-3c0eff613775](92a7c37f-fbac-4bc8-8f61-3c0eff613775) |
| 2 | LuckySheet 架构 | [bfd7e451-e175-4aed-a314-82666963dd2d](bfd7e451-e175-4aed-a314-82666963dd2d) |
| 3 | 公式引擎 | [18a84e7f-7dc4-49ac-b9be-268ac711c565](18a84e7f-7dc4-49ac-b9be-268ac711c565) |
| 4 | GitHub 缺陷目录 | [64bb840c-9956-4e98-b843-3428a3534c7e](64bb840c-9956-4e98-b843-3428a3534c7e) |
| 5 | 渲染性能 | [6a91f9aa-9e9f-488e-8daa-5559a0dd0c8a](6a91f9aa-9e9f-488e-8daa-5559a0dd0c8a) |
| 6 | 图表透视 | [294a476e-4b79-4c9d-ad90-0c35393620ac](294a476e-4b79-4c9d-ad90-0c35393620ac) |
| 7 | 导入导出协同 | [3edd1c63-efcc-4271-b789-1114145d4fdf](3edd1c63-efcc-4271-b789-1114145d4fdf) |
| 8 | 条件格式 / 数据验证 | [70879f5c-0c77-49f4-93f0-621654153e3c](70879f5c-0c77-49f4-93f0-621654153e3c) |
| 9 | 筛选排序格式 | [035bea10-1306-42f3-98de-45d7dc7995ef](035bea10-1306-42f3-98de-45d7dc7995ef) |
| 10 | 剪贴板撤销 | [676c938c-4860-46e0-a5b8-1bd6f1d0c296](676c938c-4860-46e0-a5b8-1bd6f1d0c296) |
| 11 | 交互缺陷 | [cd32d5e1-cc94-42c6-82b0-b730ca630579](cd32d5e1-cc94-42c6-82b0-b730ca630579) |
| 12 | 数据模型 | [39175912-574f-47f3-84e1-a337b1a83323](39175912-574f-47f3-84e1-a337b1a83323) |
| 13 | 官方文档 | [1e1f1893-cf86-4e57-a778-217a2ce1af09](1e1f1893-cf86-4e57-a778-217a2ce1af09) |
| 14 | 单元格编辑富文本 | [12f287fd-c6b6-4432-8605-7e7a801c255f](12f287fd-c6b6-4432-8605-7e7a801c255f) |
| 15 | 超链接 / 批注 / Table | [0cb85a17-4e3e-4cb1-8dd9-6e217e10fe96](0cb85a17-4e3e-4cb1-8dd9-6e217e10fe96) |
| 16 | Univer OSS 修复 | [06772be9-520b-40ce-baea-3dabcab7916a](06772be9-520b-40ce-baea-3dabcab7916a) |

补充交叉引用（统一计划提及，非 16 席位重复计数）：[9dafadbd-7e3b-4ff2-9cb4-096783f8855e](9dafadbd-7e3b-4ff2-9cb4-096783f8855e) 数据模型/剪贴板续挖。

---

## 1. Univer 架构

出处：`92a7c37f-fbac-4bc8-8f61-3c0eff613775`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `univer/packages/core/README.md:7` | core = 运行时 + command/mutation |
| 2 | `univer/packages/sheets/README.md:7` | sheets = 无 UI 的表格模型 |
| 3 | `univer/packages/sheets-ui/README.md:7` | sheets-ui = 选区/菜单/剪贴板/公式栏 |
| 4 | `univer/packages/engine-render/README.md:7` | canvas 渲染、滚动缩放 |
| 5 | `univer/packages/engine-formula/README.md:7` | 解析、函数、依赖、计算 |
| 6 | `univer/packages/ui/README.md:7` | 工作台 UI 框架 |
| 7 | `univer/packages/design/README.md:7` | 设计系统 |
| 8 | `univer/README.md:141` | 推荐注册：design→ui→render→formula→sheets→sheets-ui |
| 9 | `univer/packages/core/src/services/command/command.service.ts:37-54` | `CommandType` COMMAND/OPERATION/MUTATION |
| 10 | `univer/.../command.service.ts:58-60` | 数据修改必须走 command |
| 11 | `univer/packages/core/src/univer.ts:287` | DI 绑定 `ICommandService` |
| 12 | `univer/packages/sheets/src/plugin.ts:56` | sheets `@DependentOn` 公式引擎 |
| 13 | `univer/packages/sheets-ui/src/plugin.ts:37` | UI 依赖 sheets + render + ui |
| 14 | `univer/packages/core/src/sheets/typedef.ts:29-86` | `IWorkbookData` |
| 15 | `univer/packages/core/src/sheets/typedef.ts:91-160` | `IWorksheetData` |
| 16 | `univer/packages/core/src/sheets/typedef.ts:248-295` | `ICellData`：`v/f/t/s/p` |
| 17 | `univer/packages/core/src/types/interfaces/i-style-data.ts:257-354` | `IStyleBase`/`IStyleData` |
| 18 | `univer/packages/core/src/services/instance/instance.service.ts:76-119` | `IUniverInstanceService` |
| 19 | `univer/.../instance.service.ts:108-112` | `createUnit` / `disposeUnit` |
| 20 | `univer/.../instance.service.ts:97-101` | `focusUnit` / 仅 1 个 focused |
| 21 | `univer/packages/sheets/src/facade/f-range.ts:1304` | `FRange.setValue` |
| 22 | `univer/packages/sheets-formula/src/facade/f-formula.ts` | Facade 注册自定义函数 |
| 23 | `univer/packages/engine-formula/src/controllers/formula.controller.ts:135-147` | `registerExecutors` |
| 24 | `univer/README.md` Sheets 行开源 vs Pro 表 | chart/pivot/print/IO/OT 为 Pro |
| 25 | `univer/packages/core/src/sheets/typedef.ts:69-77` | 插件数据进 `resources` |
| 26 | docs.univer.ai migrate-from-luckysheet | 不可 1:1 迁移 |

**结论**：Univer 是 DI + Command/Mutation/Operation + 多 Unit 实例的插件栈；Sheets 数据在 `IWorkbookData`/`IWorksheetData`/`ICellData`，样式进 `s` 或 styleId。回移只借鉴语义（dirty、依赖图、focus preventScroll），禁止搬 Pro 包，也不把 LS 重写成 Univer monorepo。

---

## 2. LuckySheet 架构

出处：`bfd7e451-e175-4aed-a314-82666963dd2d`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `index.js:13` | `module.exports = luckysheet` |
| 2 | `core.js:45` | `common_extend(api, luckysheet)` |
| 3 | `core.js:48-49` | `create` 先 `destroy` |
| 4 | `core.js:65-66` | `Store.container` / `Store.luckysheetfile` |
| 5 | `core.js:146-150` | 插件名入 `asyncLoad`，`initPlugins` |
| 6 | `core.js:153` | `functionlist(customFunctions)` |
| 7 | `core.js:165-167` | 无 loadUrl → `initialjfFile` + `initialWorkBook` |
| 8 | `core.js:170-181` | 有 loadUrl → POST 后可选 `openWebSocket` |
| 9 | `core.js:188-199` | handler/filter/matrix/sheetBar/formulaBar/keyboard/orderBy/zoom/listener |
| 10 | `core.js:246` | `luckysheet.destroy = method.destroy` |
| 11 | `store/index.js:12-16` | `flowdata`/`config`/`visibledatarow` |
| 12 | `store/index.js:39-43` | 选区与剪贴板 |
| 13 | `store/index.js:84-86` | `clearjfundo`/`jfundo`/`jfredo` |
| 14 | `store/index.js:91-113` | `chartparam` |
| 15 | `store/index.js:142-153` | `cooperativeEdit` / `asyncLoad` |
| 16 | `method.js:431-442` | destroy 清 DOM、解绑 `.luckysheetEvent` |
| 17 | `method.js:448-453` | defaultStore 覆盖 Store（#529） |
| 18 | `method.js:455-488` | 重置 formula/sheetmanage/pivot/image/DV |
| 19 | `sheetmanage.js:715-749` | `buildGridData`：data 优先于 celldata |
| 20 | `sheetmanage.js:848-854` | `initialjfFile` + `buildGridData` |
| 21 | `sheetmanage.js:940-977` | `execF`：param、select、公式、scroll |
| 22 | `global/createdom.js:23` | `luckysheetcreatedom` |
| 23 | `controllers/expendPlugins.js:5-8` | chart/print/exportXlsx 注册表 |
| 24 | `controllers/listener.js:26-34` | `createProxy` 监听 `jfredo`/`jfundo` |
| 25 | `frontend/package.json:3` | 版本 `2.1.13` |
| 26 | `frontend/README.md:14` | 停维，推荐 Univer |

**结论**：LuckySheet 是 UMD 单例：`create` = destroy + 注入 Store + DOM + 公式链 + 事件。`flowdata` 是当前 sheet 真相源；撤销栈命名反转。协同与插件都挂在同一 Store。任何多实例方案都必须先拆这个单例（见 `docs/multi-instance-eval.md`）。

---

## 3. 公式引擎

出处：`18a84e7f-7dc4-49ac-b9be-268ac711c565`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `function/functionListDescriptor.js` | 数组长度 **375**（node require 实测） |
| 2 | `function/functionlist.js:8-16` | locale 列表绑定 `functionImplementation[func.n]` |
| 3 | `function/functionlist.js:19-21` | `customFunctions` push |
| 4 | `function/functionlist.js` | 函数表与当前计算坐标存于 `Store.runtime.formula`，不再写入 `window` |
| 5 | `function/functionImplementation.js:2976` | `SUBTOTAL` 实现入口 |
| 6 | `functionImplementation.js` SUBTOTAL 段 | **无** `rowhidden`（模块内 grep 零命中） |
| 7 | `global/api.js:189-194` | 仅 `f!=null && v==null` 走 `updatecell` |
| 8 | `global/api.js:196-214` | 同时有 f+v 走 `setcellvalue` |
| 9 | `global/api.js:228-229` | 字符串以 `=` 开头才 `updatecell` |
| 10 | `global/api.js:250-251` | 当前 sheet `jfrefreshgrid` |
| 11 | `global/api.js:2870-2882` | `setRangeValue` 深拷贝后回滚 `file.data` |
| 12 | `global/api.js:2884-2888` | `jfrefreshgrid(..., true, false)` |
| 13 | `global/setdata.js:23-27` | 对象写入可只改 `f`/`v`，不触发重算 |
| 14 | `global/refresh.js:25-36` | `runExecFunction` 收集 range 再 `execFunctionGroup` |
| 15 | `global/refresh.js:39` | `jfrefreshgrid` 默认 `isRunExecFunction=true` |
| 16 | `global/formula.js:4380` | `addFunctionGroup` 只 push `calcChain` |
| 17 | `global/formula.js:5253` | `execFunctionGroupForce` |
| 18 | `global/formula.js:5260` | `execFunctionGroup` 扫描式传播 |
| 19 | `sheetmanage.js:948` | 初始化 `execFunctionGroupForce` |
| 20 | `global/api.js:6839-6841` | `refreshFormula` 强制重算 |
| 21 | `univer/.../formula.controller.ts:135-147` | UV 批量 registerExecutors |
| 22 | `univer/.../function.service.ts:23-47` | 执行器 + 描述双轨 |
| 23 | UV CHANGELOG `#7272` | dirty ranges 先收集再渲染 |
| 24 | UV CHANGELOG `#4840` | 筛选隐藏行与公式 |
| 25 | `docs/zh/guide/sheet.md:36` | `calcChain` 为初始化字段 |
| 26 | `frontend/docs/zh/guide/sheet.md` calcChain 节点说明 | `{r,c,index,func,...}` |

**结论**：#1004 在 API 分支而不在 375 个函数实现本身。LS 只有线性 `calcChain` + 扫描 `execFunctionGroup`，无依赖图/dirty range；Worker 不跑公式。P0 应放宽「含合法 `f` 即重算」、删除 `setRangeValue` 回滚，并让 SUBTOTAL 跳过 `config.rowhidden`。

---

## 4. GitHub 缺陷目录

出处：`64bb840c-9956-4e98-b843-3428a3534c7e`  
完整映射见 `docs/issue-module-map.md`。

| # | 证据 | 要点 |
|---|---|---|
| 1 | https://github.com/dream-num/Luckysheet/issues/1004 | OPEN：setRangeValue 公式不联动 |
| 2 | `api.js:189` | 与 #1004 备注一致 |
| 3 | https://github.com/dream-num/Luckysheet/issues/504 | OPEN：首次点击滚顶 |
| 4 | `updateCell.js:98` | 裸 `focus().select()` |
| 5 | https://github.com/dream-num/Luckysheet/issues/794 | CLOSED；本 fork 编辑路径未齐 |
| 6 | `utils/util.js:506-507` | 注释 `fix #794 #152` |
| 7 | https://github.com/dream-num/Luckysheet/issues/529 | OPEN：destroy/create 滚动 |
| 8 | `method.js:448-453` | Store 整表重置 |
| 9 | https://github.com/dream-num/Luckysheet/issues/799 | 停维公告 |
| 10 | https://github.com/dream-num/Luckysheet/issues/1454 | EOL |
| 11 | https://github.com/dream-num/Luckysheet/issues/479 | 跨表公式 |
| 12 | https://github.com/dream-num/Luckysheet/issues/477 | 保护失效 |
| 13 | https://github.com/dream-num/Luckysheet/issues/451 | 6000 行筛选丢数据 |
| 14 | https://github.com/dream-num/Luckysheet/issues/213 | 协同离线丢失 |
| 15 | https://github.com/dream-num/Luckysheet/issues/214 | 切 sheet 广播 |
| 16 | https://github.com/dream-num/Luckysheet/issues/1435 | 拖拉公式值错 |
| 17 | https://github.com/dream-num/Luckysheet/issues/1420 | setRangeValue 首格格式 |
| 18 | https://github.com/dream-num/Luckysheet/issues/181 | setCellValue 超行 |
| 19 | https://github.com/dream-num/Luckysheet/issues/412 | `%` 公式 |
| 20 | https://github.com/dream-num/Luckysheet/issues/417 | getRangeHtml |
| 21 | https://github.com/dream-num/Luckysheet/issues/1436 | IME 拼音 |
| 22 | https://github.com/dream-num/Luckysheet/issues/127 | 图表报错 |
| 23 | GitHub 仓页 | 2025-10-30 archived |
| 24 | `frontend/README.md:16` | Follow #1454 |
| 25 | V2 统计 | label:BUG open ≈ 23 |

**结论**：fork 必须自修。功能 P0 是 #1004/#504（及 #794 回归）/#479/#451/#477/#213；#799/#1454 只作产品边界，不进修复队列。

---

## 5. 渲染性能

出处：`6a91f9aa-9e9f-488e-8daa-5559a0dd0c8a`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `store/index.js:12` | 全量 `flowdata` |
| 2 | `sheetmanage.js:715-749` | 稀疏 celldata 物化稠密网格 |
| 3 | `global/createsheet.js` grow 逻辑 | 按 row/column 扩阵 |
| 4 | `global/draw.js:384` | `luckysheetDrawMain` |
| 5 | `global/draw.js:395` | `flowdata == null` 直接 return |
| 6 | `global/draw.js:456` | `clearRect(0,0,全宽,全高)` |
| 7 | `global/draw.js:461-462` | `visibledatarow` 二分可见行 |
| 8 | `global/refresh.js:1161` | `luckysheetrefreshgrid` |
| 9 | `global/refresh.js:39-42` | 无 data 则用 `Store.flowdata` |
| 10 | `controllers/scroll.js` 滚动回调 | 再入 refresh/draw |
| 11 | `global/rhchInit.js` | 累加 `visibledatarow/column` |
| 12 | `store/index.js:15-16` | 可见行列数组 |
| 13 | `draw.js` `conditionformat.getComputeMap` | 绘制时算 CF |
| 14 | `global/editor.js:30` | `webWorkerFlowDataCache` 整表深拷贝 |
| 15 | `refresh.js` `webWorkerFlowDataCache` | 刷新后整表拷贝 |
| 16 | UV `#7394` | merge-aware incremental scroll |
| 17 | UV engine-render Viewport | 虚拟滚动 |
| 18 | UV SpreadsheetSkeleton | 布局与合并几何 |
| 19 | `frontend/README.md:14` | 「大数据量加载」列为 Univer 已解决问题 |
| 20 | `store/index.js:20-29` | 全表像素尺寸字段 |
| 21 | `api.js:6677-6712` | 稠密↔稀疏转换 API |
| 22 | `getAllSheets` `api.js:5908` | 导出再扫整表 |
| 23 | `jfrefreshgrid` 无 dirtyRect 参数 | `refresh.js:39` 签名 |
| 24 | `store/index.js:120-123` | measureText 缓存，非脏矩形 |
| 25 | `sheetmanage.js:721` | 已有 data 仍逐格 `setcellvalue` |

**结论**：虚拟化只在 canvas 裁剪，数据层仍是稠密 `flowdata`。Phase 2 应先稀疏存储，再给 `jfrefreshgrid`/`luckysheetDrawMain` 加 dirty rect；不要指望先换渲染引擎。

---

## 6. 图表 / 透视

出处：`294a476e-4b79-4c9d-ad90-0c35393620ac`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `expendPlugins/chart/plugin.js:26-29` | unpkg：vue@2.6 / vuex / element-ui / echarts@4.8 |
| 2 | `expendPlugins/chart/plugin.js:34` | element-ui CSS CDN |
| 3 | `expendPlugins.js:6` | `'chart': chart` |
| 4 | `store/index.js:91-113` | `chartparam` 拖动/缩放状态 |
| 5 | `method.js:501` | `insertChartTosheet` |
| 6 | `server.js:938` | 协同插入图表 |
| 7 | `demoData/sheetChart.js` | 官方 demo 图表 JSON |
| 8 | `controllers/pivotTable.js:729` | `createPivotTable` |
| 9 | `controllers/pivotTable.js:787` | `refreshPivotTable` |
| 10 | `docs/zh/guide/sheet.md:37-38` | `isPivotTable` / `pivotTable` |
| 11 | `docs/zh/guide/sheet.md:45` | `chart[]` |
| 12 | `demoData/sheetPivotTable.js` | 透视 demo |
| 13 | `demoData/sheetPivotTableData.js` | 源数据 sheet |
| 14 | Univer OSS glob chart/pivot 包 | 开源仓无 engine |
| 15 | docs.univer.ai pivot-table | Pro 功能 |
| 16 | docs.univer.ai getting-started/pro | 未授权限制 |
| 17 | `api.js:5923` | `getAllChartsBase64` 依赖页面 echarts |
| 18 | `exportXlsx/plugin.js:31-33` | 导出带 `chartMap` |
| 19 | FAQ/CDN | 离线图表失败（#127） |
| 20 | `core.js:198` | print 初始化已注释，图表仍靠 plugin |
| 21 | `store/index.js:88-89` | `createChart` / `highlightChart` 钩子字符串 |
| 22 | `locale` 图表文案 | UI 依赖 ChartMix |
| 23 | Pro `@univerjs-pro/sheets-chart` | 禁止复制 |
| 24 | Pro `@univerjs-pro/sheets-pivot` | 禁止复制 |
| 25 | `sheetmanage.js:954-960` | 透视 sheet 恢复走 `restoreSheetAll` |

**结论**：图表/透视只能自研加固：CDN 本地化 + `pivotTable.refresh` 稳定性。不能从 Univer OSS 抄引擎，也不能把 Pro 当源码。

---

## 7. 导入导出 / 协同

出处：`3edd1c63-efcc-4271-b789-1114145d4fdf`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `expendPlugins/exportXlsx/plugin.js:10-12` | 插件只摘 asyncLoad |
| 2 | `exportXlsx/plugin.js:30-31` | `fetchAndDownloadXlsx` 调 `toJson()` |
| 3 | `backend/luckysheet/.../ExcelIoController.java` | Java 转 xlsx / 导入 |
| 4 | `backend/luckysheet/.../ExcelIoService.java` | luckysheet-lib 适配 |
| 5 | `backend/README-zh.md` | Java 后端说明 |
| 6 | `expendPlugins/print/plugin.js:2` | import 空壳 `print.js` |
| 7 | `expendPlugins/print/print.js` | **0 字节** |
| 8 | `controllers/server.js:34-42` | 单格 `saveParam("v")` |
| 9 | `server.js:79` | 范围 `rv` |
| 10 | `server.js:87-90` | `allowUpdate` 才发送 |
| 11 | `core.js:178-180` | create 后 `openWebSocket` |
| 12 | `api.js:6761` | `closeWebsocket` |
| 13 | `refresh.js` `saveParam("all")` | config/calcChain/CF/DV/filter… |
| 14 | Univer README Pro 表 | Import/Export Server、Print、OT |
| 15 | OSS `importXLSXToSnapshotAsync` | 开源仓 0 实现 |
| 16 | `api.js:6719-6739` | `toJson` 当专有格式 |
| 17 | `core.js:170-172` | `loadUrl` 用 `new Function("return "+d)` 解析 |
| 18 | `method.js` `addDataAjax` | 分页加载 |
| 19 | `server.js` `wsUpdateMsg` | 下行补丁 |
| 20 | `store/index.js:142-150` | 协同光标/行高跟随 |
| 21 | Issue #213 / #214 | 离线丢失、切表广播 |
| 22 | `packages/rpc`（Univer） | Worker RPC，不是 OT 产品 |
| 23 | `api.js:6893` | `saveParam("shs")` 切表 |
| 24 | `frontend/README.md:32` | 导入导出打印请用 Univer |
| 25 | `docs/print-blocked.md` | 打印 Blocked 专文 |

**结论**：IO/打印/OT 必须自研或外购，不能抄 Univer Pro。协同是 LWW WebSocket，先加固版本号、切表边界和重连，而不是上 OT。

---

## 8. 条件格式 / 数据验证

出处：`70879f5c-0c77-49f4-93f0-621654153e3c`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `docs/zh/guide/sheet.md:43` | `luckysheet_conditionformat_save` |
| 2 | `docs/zh/guide/sheet.md:49` | `dataVerification` |
| 3 | `conditionformat.js:70` | `dataBarList` |
| 4 | `conditionformat.js:85` | `colorGradationList` |
| 5 | `conditionformat.js:263-274` | type `dataBar` |
| 6 | `conditionformat.js:279-292` | type `colorGradation` |
| 7 | `conditionformat.js:297-309` | type `icons` |
| 8 | `conditionformat.js:316-320` | `conditionName = type2` |
| 9 | `conditionformat.js:428` | `textContains` |
| 10 | `conditionformat.js:465` | `occurrenceDate` |
| 11 | `conditionformat.js:482-493` | top10 / last10 |
| 12 | `conditionformat.js:509` | `AboveAverage` |
| 13 | `demoData/sheetConditionFormat.js` | 全类型样例 |
| 14 | `locale/zh.js` CF 文案段 | 快捷菜单键 |
| 15 | `dataVerificationCtrl.js:21` | 默认 `type: 'dropdown'` |
| 16 | `dataVerificationCtrl.js:101-109` | dropdown/checkbox/number_*/text_*/validity |
| 17 | `dataVerificationCtrl.js:212` | `remote` 勾选 |
| 18 | `api.js:6103-6120` | `setDataVerification` 9 类型白名单 |
| 19 | `api.js:6120` | typeValues 含 validity |
| 20 | UV CF `CFRuleType` | highlightCell/colorScale/iconSet/dataBar |
| 21 | UV DV `DataValidationType` | ANY/CUSTOM/LIST/… |
| 22 | `draw.js` `checksCF` | 绘制期计算 |
| 23 | `store/index.js:168` | `conditionFormatCells` |
| 24 | `api.js:3937` | CF 协同 `k:luckysheet_conditionformat_save` |
| 25 | `sheet.md` alternateformat | 交替色 ≠ CF |
| 26 | `dataVerificationCtrl.js` `validateCellData` | 运行时校验 |

**结论**：四主类同构，但 LS 扁平 `conditionName` 对 UV 三层枚举。P0 做 adapter；`regExp`/`sort`/`validity`/`remote` 作为 LS 扩展保留。

---

## 9. 筛选 / 排序 / 数字格式

出处：`035bea10-1306-42f3-98de-45d7dc7995ef`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `controllers/filter.js:30-62` | 选项状态带 `rowhidden` |
| 2 | `filter.js:149` | container 裸 `focus()` |
| 3 | `filter.js:489-500` | 多列隐藏行合并 |
| 4 | `filter.js:818` | 读 `Store.config.rowhidden` |
| 5 | `docs/zh/guide/sheet.md:39-40` | `filter_select` / `filter` |
| 6 | `controllers/orderBy.js:187-195` | 有合并则拒绝排序 |
| 7 | `global/sort.js` | `orderbydata` |
| 8 | `api.js:3460` | `setRangeSort` |
| 9 | `api.js:3555` | `setRangeSortMulti` |
| 10 | `global/format.js:1436` | `w`/`w0.00` 万亿注释 |
| 11 | `global/format.js:1759` | `genarate` 万单位 |
| 12 | `frontend/package.json:47` | `numeral` |
| 13 | `functionImplementation.js:2976` | SUBTOTAL 不读筛选 |
| 14 | `api.js:3032` | `setRangeFilter` |
| 15 | UV `#4840` | 筛选↔公式 |
| 16 | UV `#5862` | 数字当字符串比较 |
| 17 | UV `#5797` | 筛选列同步 |
| 18 | UV `#7163` | 科学计数法边界 |
| 19 | UV `#7383` | 空格式 coverable |
| 20 | `searchReplace.js:520` | 查找用 `valueShowEs`，不读 `cell.f` |
| 21 | `api.js:450` | `find` 默认 type `m` |
| 22 | UV `FindBy.FORMULA` `#6970` | 跨表/公式查找 |
| 23 | `orderBy.js` 筛选菜单排序 | 单列 |
| 24 | UV `sheets-sort.command.ts` | 跳过 hidden/filtered |
| 25 | `docs/zh/guide/cell.md:184-198` | fa 含科学计数与万 |

**结论**：筛选三分法可对齐；P0 是 SUBTOTAL 跳过 `rowhidden`。P1：列同步、排序数组公式校验、科学计数、公式查找。`w/W` 必须当 LS 扩展 pattern 留下。

---

## 10. 剪贴板 / 撤销

出处：`676c938c-4860-46e0-a5b8-1bd6f1d0c296`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `store/index.js:42-43` | `luckysheet_copy_save` / `paste_iscut` |
| 2 | `handler.js:5931-5947` | document `paste` 读 html/plain |
| 3 | `handler.js:5275` | 按钮 `selection.paste` |
| 4 | `handler.js:3873` | 格式刷 `pasteHandlerOfPaintModel` |
| 5 | `controllers/selection.js:38` | textarea `focus` 无 preventScroll |
| 6 | `selection.js:550,592,618` | 复制粘贴 focus |
| 7 | `store/index.js:84-86` | jfundo/jfredo |
| 8 | `controlHistory.js:63-68` | 名为 `redo` 的函数 pop `jfredo` → 实际撤销 |
| 9 | `controlHistory.js:448-453` | 名为 `undo` 的函数 pop `jfundo` → 实际重做 |
| 10 | `api.js:5860-5863` | `undo()` 调 `controlHistory.redo` |
| 11 | `api.js:5885-5888` | `redo()` 调 `controlHistory.undo` |
| 12 | `refresh.js:39` | 刷新推历史 |
| 13 | `listener.js` Proxy | 监听栈变化 |
| 14 | UV `IUndoRedoService` | mutation 栈，命名正常 |
| 15 | UV SPECIAL_PASTE_VALUE | 仅值粘贴 |
| 16 | 计划约束 | **禁止改 jfundo/jfredo 名** |
| 17 | `handler.js` 内部 copy table 标记 | 内部粘贴识别 |
| 18 | `refresh.js` `jfrefreshgrid_pastcut` | 剪切粘贴刷新 |
| 19 | UV `#6686` | cut/paste 更新公式引用 |
| 20 | UV `#6738` | HTML 消毒 |
| 21 | UV `#6326` | 粘贴 CF/DV/numfmt |
| 22 | `getRangeHtml` `api.js:1834` | HTML 导出，#417 |
| 23 | `selection.js` copy HTML table | 内部格式 |
| 24 | `keyboard.js` Ctrl+V | 进 paste 链路 |
| 25 | `method.js:441` | destroy `off(.luckysheetEvent)` 含 paste |

**结论**：剪贴板是 jQuery document 粘贴 + 内部 table 标记，没有「仅值/仅格式」独立命令。撤销必须按 `controlHistory` 实际行为写测试，不能按单词 undo/redo 望文生义。

---

## 11. 交互缺陷

出处：`cd32d5e1-cc94-42c6-82b0-b730ca630579`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `updateCell.js:98` | `$("#luckysheet-rich-text-editor").focus().select()` |
| 2 | `formulaBar.js:63` | Enter 后裸 focus |
| 3 | `formulaBar.js:71` | ESC 后裸 focus |
| 4 | `formula.js:3487` | `el.focus()` 无 preventScroll |
| 5 | `filter.js:149` | container 裸 focus |
| 6 | `util.js:486-494` | **仅 fullscreenmode** 才 preventScroll |
| 7 | `util.js:490` | `input.focus({ preventScroll: true })` |
| 8 | `util.js:506-507` | container focus 修 #794 |
| 9 | `config` `fullscreenmode: true` | 可关 |
| 10 | `selection.js:38,550,592,618` | 剪贴板 textarea focus |
| 11 | `api.js:618` | `enterEditMode` → updateCell |
| 12 | `api.js:6950` | 搜索框 focus |
| 13 | `searchReplace.js:144,150` | 查找替换 focus |
| 14 | `postil.js:493,598` | 批注 focus |
| 15 | `cursorPos.js:8,38` | 定位用 focus |
| 16 | `store/index.js:1-170` | 单例 |
| 17 | 约 74 文件 import Store | 多实例互串 |
| 18 | `createdom.js` 全局 id | `#luckysheet-*` |
| 19 | `handler.js` scroll/resize 绑定 | destroy 未对 window 全量 off |
| 20 | `method.js:445-446` | destroy 只重置 freezen initial 旗标 |
| 21 | UV `focus-editor.ts:35` | 统一 preventScroll |
| 22 | Issue #504 / #794 | 嵌套页滚顶 |
| 23 | `core.js:70` | `Store.fullscreenmode = extendsetting.fullscreenmode` |
| 24 | `keyboard.js` 进编辑 | 调 updateCell |
| 25 | `handler.js` 双击编辑 | 调 updateCell |

**结论**：本 fork 仍具备高概率复现 #504/#794。必须统一 `focusEditor(el,{preventScroll:true})`，覆盖 updateCell/formulaBar/formula/filter/selection，不能只修 container。

---

## 12. 数据模型

出处：`39175912-574f-47f3-84e1-a337b1a83323`（交叉 `9dafadbd`）

| # | 证据 | 要点 |
|---|---|---|
| 1 | `docs/zh/guide/sheet.md:9-50` | 初始化 JSON 全字段 |
| 2 | `docs/zh/guide/sheet.md:22` | celldata 初始化 |
| 3 | `docs/zh/guide/cell.md:125-144` | `v`/`m`/`f` |
| 4 | `docs/zh/guide/cell.md:17-24` | `ct` |
| 5 | `docs/zh/guide/cell.md:97-102` | `mc` |
| 6 | `docs/zh/guide/cell.md:146+` | `ps` |
| 7 | `sheetmanage.js:700-713` | `getGridData` |
| 8 | `sheetmanage.js:715-726` | data 优先 |
| 9 | `sheetmanage.js:728-742` | celldata → 二维 |
| 10 | `api.js:6677` | `transToCellData` |
| 11 | `api.js:6699` | `transToData` |
| 12 | `store/index.js:12` | flowdata |
| 13 | `sheetmanage.js` `setSheetParam` | flowdata = file.data |
| 14 | `sheetmanage.js:1173` | storeSheetParamALL 回写 data |
| 15 | `refresh.js` 双写 flowdata+file.data | 刷新契约 |
| 16 | `typedef.ts:248-269` | UV 无 `m` |
| 17 | `i-style-data.ts:319-320` | UV pattern 在 `n` |
| 18 | `typedef.ts:122-126` | UV `mergeData` + `cellData` |
| 19 | `api.js:5908-5917` | 导出转 celldata，删 freezen |
| 20 | `store/index.js:155-167` | `defaultCell` 扁平样式 |
| 21 | migrate-from-luckysheet | 不可 1:1 |
| 22 | `docs/field-mapping-ls-uv.md` | ≥24 行对照 |
| 23 | `calcChain` sheet.md:36 | 公式必需链 |
| 24 | `config.rowhidden` sheet.md:27 | 筛选写入目标 |
| 25 | `IWorkbookData.resources` typedef.ts:77 | UV 插件快照 |

**结论**：契约是「celldata 进、data 存、flowdata 算」。adapter 必须回算 `m`、保住 `calcChain`/`config`，禁止把运行时字段当存储。

---

## 13. 官方文档

出处：`1e1f1893-cf86-4e57-a778-217a2ce1af09`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `frontend/README.md:14` | no longer maintained |
| 2 | `frontend/README.md:16` | Follow #1454 |
| 3 | `frontend/README.md:32` | 导入导出打印用 Univer |
| 4 | `frontend/README-zh.md` | 中文停维 + 大数据/图表/透视/公式 |
| 5 | Issue #1454 | luckysheet is EOL |
| 6 | Issue #799 | stopped maintenance |
| 7 | blog.univer.ai 中文对照文 | 官方推荐迁移 |
| 8 | docs.univer.ai migrate-from-luckysheet | 不能 1:1 |
| 9 | docs.univer.ai guides/sheets | Sheets 产品文档 |
| 10 | docs.univer.ai guides/pro | Pro 边界 |
| 11 | docs.univer.ai import-export | Server + Pro client |
| 12 | docs.univer.ai pivot-table | Pro |
| 13 | docs.univer.ai getting-started/pro | 未授权限制 |
| 14 | `frontend/docs/zh/guide/api.md` | 对外 API 说明书 |
| 15 | `frontend/docs/zh/guide/cell.md` | 单元格字段 |
| 16 | `frontend/docs/zh/guide/sheet.md` | 工作表字段 |
| 17 | `frontend/docs/guide/data.md` | celldata 仅初始化 |
| 18 | github.com/dream-num/migrate-luckysheet | 404 |
| 19 | awesome-univer/migrate-luckysheet | 社区迁移器 |
| 20 | LuckysheetDocs 特性列表 | 历史功能宣传 |
| 21 | Univer Apache-2.0 vs LS MIT | 许可证不同 |
| 22 | `frontend/package.json:3` | 2.1.13 |
| 23 | GitHub archived 2025-10-30 | 只读 |
| 24 | blog：Luckysheet 架构单一耦合高 | 官方评价 |
| 25 | blog：WS 缺复杂冲突处理 | 与 server.js LWW 一致 |

**结论**：官方关系坐实：LS EOL，Univer 是继任产品但不是源码超集。本 fork 必须在 README/CHANGELOG 声明维护范围与 Pro 边界（Phase 5）。

---

## 14. 单元格编辑 / 富文本

出处：`12f287fd-c6b6-4432-8605-7e7a801c255f`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `updateCell.js:98` | 进编辑 focus |
| 2 | `controllers/inlineString`（getdata 引用） | inlineStr `ct.s` |
| 3 | `global/getdata.js:8` | `isInlineStringCT` |
| 4 | `global/getdata.js:283-314` | 读 `ct.s` 拼显示 |
| 5 | `global/getRowlen.js:426-439` | 行高考虑 inline |
| 6 | `store/index.js:133-134` | inline 编辑缓存 |
| 7 | `method.js:148-149` | destroy 默认含 inline 缓存字段 |
| 8 | `api.js:228` | `"<span"` 走 updatecell |
| 9 | `global/editor.js` | 编辑器 / Worker 缓存 |
| 10 | `formulaBar.js` | 公式栏与格子双编辑器 |
| 11 | `formula.js` `functionInputHanddler` | 输入着色 |
| 12 | `cursorPos.js:8` | 光标定位 focus |
| 13 | Issue #1436 | IME 拼音不同步 |
| 14 | Issue #183 | 换行后方向键 |
| 15 | Issue #425 | Firefox 多空行 |
| 16 | Issue #178 | Plain Text 内容消失 |
| 17 | UV sheets-ui 编辑器 | 可复用 docs 引擎需文件证明 |
| 18 | UV `doc-ime-input-manager.service.ts` | composition 状态机 |
| 19 | `store/index.js:65` | `luckysheetCellUpdate` |
| 20 | `api.js:570-578` | `exitEditMode` → `formula.updatecell` |
| 21 | `api.js:608` | `enterEditMode` |
| 22 | `keyboard.js` | 键入进编辑 |
| 23 | `handler.js` 双击 | 进 `luckysheetupdateCell` |
| 24 | `docs/zh/guide/cell.md` inlineStr | `ct.t` 富文本 |
| 25 | UV `ICellData.p` | 富文本文档模型 |

**结论**：LS 编辑是 DOM 富文本 + 公式栏，IME/composition 未做状态机。P1 与 focus 修复绑定；富文本对 UV 走 `p`，回移要还原 `ct.s`。

---

## 15. 超链接 / 批注 / Table

出处：`0cb85a17-4e3e-4cb1-8dd9-6e217e10fe96`

| # | 证据 | 要点 |
|---|---|---|
| 1 | `hyperlinkCtrl.js:20` | `linkType: 'external'` |
| 2 | `hyperlinkCtrl.js:48-50` | 仅 external / internal |
| 3 | `hyperlinkCtrl.js:129-153` | 保存 linkType |
| 4 | `hyperlinkCtrl.js:238` | 点击跳转分支 |
| 5 | UV hyperlink `DEFINE_NAME` `#6950` | LS 无命名区域链接 |
| 6 | `postil.js:951-952` | `cell.ps` 构建批注 |
| 7 | `docs/zh/guide/cell.md:146` | `ps` comment |
| 8 | `demoData/sheetComment.js` | 批注 demo |
| 9 | UV `@univerjs/sheets-note` `#5125` | Cell Note 独立概念 |
| 10 | UV thread-comment `#6042` | 线程评论 ≠ ps |
| 11 | `demoData/sheetTable.js:1-2` | `window.sheetTable` 仅 sheet 名 Table |
| 12 | `sheetTable.js` | 无结构化 table 对象 |
| 13 | UV `@univerjs/sheets-table` | id/name/range/columns |
| 14 | `docs/zh/guide/sheet.md:45-48` | chart/image 有，无 table 数组 |
| 15 | `refresh.js` k 清单 | 有 hyperlink，无 table |
| 16 | `server.js` 类型 | 可广播 hyperlink |
| 17 | `api.js` | 无 insertHyperlink 独立导出 |
| 18 | `filter.js` | 筛选不绑定 Table 列 |
| 19 | `store` defaultCell | 无 ps/note |
| 20 | UV RANGE link payload | LS internal 仅 sheet+单元格 |
| 21 | `postil.js` focus | 批注框滚顶风险 |
| 22 | `locale` hyperlink 文案 | 两种类型 |
| 23 | 计划 Phase 3.12–3.14 | DEFINE_NAME / Table / Note |
| 24 | migrate-luckysheet Supported | 未含 chart/pivot/image/filter |
| 25 | `getAllSheets` 不删 hyperlink | 超链接随 sheet 导出 |

**结论**：超链接只有外链/内链；批注是单体 `ps`；`sheetTable.js` 不是 Excel Table。Table/Note 属于新增 JSON 契约，需产品决策与 UV 插件对照，而不是改现有 demo 名字交差。

---

## 16. Univer OSS 可借鉴修复

出处：`06772be9-520b-40ce-baea-3dabcab7916a`

| # | 证据 | 回移含义 |
|---|---|---|
| 1 | CHANGELOG `#7272` | 先收集 dirty 再渲染 → `runExecFunction` |
| 2 | `#7265` | 增量公式更新 |
| 3 | `#7306` | 非激活 sheet 删行更新公式引用 |
| 4 | `#7359` | Excel 公式兼容 |
| 5 | `#4840` | 筛选隐藏行公式 / SUBTOTAL |
| 6 | `#5506` | SUBTOTAL 忽略嵌套 SUBTOTAL |
| 7 | `#5862` | 筛选数字当字符串 |
| 8 | `#5797` | 筛选列同步 |
| 9 | `#6686` | cut/paste 公式引用 |
| 10 | `#6727` `#6729` | 跨 sheet 复制 CF/DV |
| 11 | `#6326` `#6257` | 粘贴样式/CF/DV |
| 12 | `#7163` | 科学计数法 |
| 13 | `#7383` | 空格式 coverable |
| 14 | `#6970` | 查找跨表滚动 |
| 15 | `#6950` | 超链接 DEFINE_NAME |
| 16 | `#7394` | 合并格增量滚动 |
| 17 | `#6969` | 冻结区 ghosting |
| 18 | `#3196` | 排序 colIndex |
| 19 | `#5125` | Cell Note |
| 20 | `#7344` | sheets-table |
| 21 | `#6909` | 动态数组 spill |
| 22 | `#6738` | 剪贴板 HTML 消毒 |
| 23 | `docs/tldr` formula-engine-architecture | 依赖树 + dirty |
| 24 | `docs/tldr` web-worker-architecture | Worker 公式（LS 不宜整包搬） |
| 25 | 44 条 OSS 清单（V2 全文） | 均非 `@univerjs-pro` |

**结论**：最值得先搬的是 **数据一致性**（公式 dirty、#1004、筛选 SUBTOTAL、引用变换）和 **交互**（preventScroll、冻结滚动），而不是 Worker 公式架构或 editor-scoped undo。每条 CHANGELOG 只作语义对照，在 LS 单文件链路上重做。

---

## 总册结论

16 个 V2 模块指向同一产品决策：**保留 `window.luckysheet` 与 `celldata/data/config/calcChain` 契约**，用 Univer OSS 的修复语义补 P0 正确性与交互，用自研补图表/透视/IO/打印，明确 Pro 不可抄。Phase 1 验收必须能指回本册行号与 `docs/issue-module-map.md` 的复现步骤。
