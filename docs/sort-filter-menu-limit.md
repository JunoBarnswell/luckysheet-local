# 筛选菜单内排序限制

主排序入口（菜单栏「排序」对话框、`orderBy.js` / `sort.js`）已按 Univer `sheets-sort` 语义：

- 跳过 `config.rowhidden` 与手动隐藏行，只在可见行之间重排；
- 与数组公式 / `dynamicArray` spill 范围相交时中止并提示。

**筛选下拉菜单里的升序/降序仍是单列排序。**

入口：`frontend/src/controllers/filter.js` 的 `orderbydatafiler`，由 `#luckysheet-filter-orderby-asc` / `#luckysheet-filter-orderby-desc` 触发。该文件属于 Phase 1 并行改动范围，本 Phase **不改** 筛选菜单实现。

因此：

- 筛选菜单排序只按当前列比较，不会弹出多关键字对话框；
- 若需要多列排序，请使用工具栏/菜单「自定义排序」。

locale 键：`sort.filterMenuSingleColumnLimit`。
