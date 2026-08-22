# Changelog（本 fork）

官方 [dream-num/Luckysheet](https://github.com/dream-num/Luckysheet) 已 archived（EOL，最后发布 **2.1.13**）。上游历史仍保留在 [`frontend/CHANGELOG.md`](frontend/CHANGELOG.md)。

**本文件只记录本 fork 相对 2.1.13 的变更。**

版本策略见 [README-fork.md](README-fork.md#版本策略)：

- **2.1.x patch**：只收 bugfix（公式联动、focus 滚顶、destroy 滚动、SUBTOTAL+筛选等）。
- **2.2.x minor**：收功能（稀疏网格、CF/DV adapter、Table/Note、图表本地 vendor、导出增强、协同加固等）。

当前工作簿源码仍标 `frontend/package.json` = `2.1.13`。下列条目按策略归类，**尚未打正式 npm tag**。

---

## [2.2.0] — 未发布（Phase 2–4 功能）

### Features

- **稀疏网格**：`buildGridData` / `Store.flowdata` 改为稀疏存储 + `data[r][c]` 外观；`toJson` 仍导出 `celldata`。写入用 `cloneSheetData`，不再 `$.extend(true, [], file.data)` 把整表稠密化。
- **增量重绘**：`jfrefreshgrid` / `luckysheetrefreshgrid` / `luckysheetDrawMain` 携带 `dirtyRect`；局部 `clearRect`；合并格跨越 dirty 边界时 `expandDirtyRectForMerges`。
- **条件格式 adapter**：`cfAdapter.js` 映射 colorScale / iconSet / highlight 算子；补 `notEqual` / `notBetween` / `gte` / `lte`、文本 beginsWith/endsWith/notContains、`stopIfTrue`。保留 LS 扩展 `regExp` / `sort`。
- **数据验证 adapter**：`dvAdapter.js` 映射 9 类型；补 ANY / CUSTOM / LIST_MULTIPLE。保留 `validity(card/phone)` 与 remote 为 LS 扩展。
- **筛选同步 / 双条件 AND-OR**：行列插删同步 `filter_select` 与列索引（#5797 语义）；CUSTOM 可选 AND/OR；数值不再当字符串比较（#5862 语义）。
- **排序**：跳过 `config.rowhidden`；与数组公式 / 动态数组相交则中止。筛选菜单内排序仍为单列，见 `docs/sort-filter-menu-limit.md`。
- **数字格式**：科学计数法边界（#7163 语义）；空格式 `coverable:false`（#7383 语义）；保留 `w`/`W` 万/亿，见 `docs/numfmt-wan-extension.md`。
- **查找替换**：`FindBy.FORMULA` 读 `cell.f`；工作簿级跨 sheet 查找并定位。
- **剪贴板**：外部粘贴「仅值 / 仅格式」；撤销仍走 `jfundo`/`jfredo`（不改命名）。
- **超链接**：`linkType` 支持 `DEFINE_NAME` 与 RANGE payload。
- **结构化 Table**：`sheet.table[]`（`id/name/range/columns/tableStyleId`）；`sheetTable.js` demo 改为真实 Table 对象。
- **Cell Note**：`cell.note` / `sheet.notes` 黄三角 + hover；与批注 `ps` 红三角并存，见 `docs/cell-note-vs-postil.md`。
- **IME**：`formula.js` compositionstart/update/end，组合态暂停公式着色与 Enter 提交。
- **图表**：`expendPlugins/chart` 去掉 unpkg；vue@2.6 / vuex@3.4 / element-ui@2.13 / echarts@4.8 本地 vendor。
- **透视**：源数据变更触发 `pivotDatas` 重算；占位框用 `pivotTableBoundary`，禁止写死 12×6。
- **导出**：`luckyexcel-node` 补样式 / 公式 / 合并；`exportXlsx` 插件失败回调。不依赖 Univer Server。差距见 `docs/export-fidelity.md`。
- **协同加固**：消息版本号、切 sheet 边界、断线重连。协议见 `docs/collab-protocol.md`。**不实现 Pro OT**。

### Docs

- `docs/print-blocked.md`：打印 **Blocked**（`expendPlugins/print/print.js` 为 0 字节）。
- `docs/multi-instance-eval.md`：多实例仅评估，不实施。
- `docs/perf-baseline.md`：性能指标与测法；未测项不填数字。

---

## [2.1.14] — 未发布（Phase 1 bugfix，按 patch 策略）

### Bug Fixes

- **[#1004](https://github.com/dream-num/Luckysheet/issues/1004)**：`setCellValue` / `setRangeValue` 只要 `f` 是合法公式串（以 `=` 开头）就走重算；删除 `setRangeValue` 循环后回滚 `file.data`；`markFormulaDirty` + `execFunctionGroup` 刷新依赖。
- **[#504](https://github.com/dream-num/Luckysheet/issues/504) / [#794](https://github.com/dream-num/Luckysheet/issues/794)**：编辑入口统一 `focusEditor(el, { preventScroll: true })`（`updateCell` / `formulaBar` / `formula.js` / `filter.js`），嵌套页单击不再滚顶。
- **[#529](https://github.com/dream-num/Luckysheet/issues/529)**：`destroy` 不改调用方 `options.data[].scrollTop`；滚动恢复只走 `sheetmanage.applySheetScroll` 一处。
- **SUBTOTAL + 筛选**：`SUBTOTAL` 读取 `config.rowhidden`，聚合跳过筛选隐藏行（对照 Univer #4840 语义）。

---

## 已知未做 / Blocked

| 项 | 状态 |
|---|---|
| 打印引擎 | Blocked，`print.js` 0 字节，见 `docs/print-blocked.md` |
| Univer Pro chart / pivot / print / OT / import-export Server | **不做**，禁止抄闭源 |
| Store 多实例 | 仅评估，见 `docs/multi-instance-eval.md` |
| 与 Excel 完整 fidelity | 图表/图片/CF/DV/批注等未进 xlsx，见 `docs/export-fidelity.md` |
