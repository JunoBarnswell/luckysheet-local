# LuckySheet ↔ Univer 字段映射

> 出处：[数据模型深挖V2](39175912-574f-47f3-84e1-a337b1a83323)；剪贴板侧参考 [数据模型剪贴板深挖](9dafadbd-7e3b-4ff2-9cb4-096783f8855e)。  
> LS 字段来自 `frontend/docs/zh/guide/sheet.md`、`frontend/docs/zh/guide/cell.md` 与运行时代码。  
> UV 字段来自 `C:\Users\kuo13\Projects\univer\packages\core\src\sheets\typedef.ts`、`packages/core/src/types/interfaces/i-style-data.ts`。  
> 官方原文：Luckysheet *can no longer be migrated 1:1*（docs.univer.ai migrate-from-luckysheet）。

## 映射等级

| 等级 | 含义 | adapter 策略 |
|---|---|---|
| 1:1 | 同名或可无损互转 | 直接复制 / 0-1 ↔ BooleanNumber |
| 部分 | 语义接近，枚举或结构不同 | 显式对照表；回移时回算 |
| 不可映射 | 对方模型无对等字段 | LS 保留扩展 / UV 进 `custom` 或 `resources`；禁止丢持久化契约 |

## 工作簿 / 工作表（≥ 字段表）

| # | LS | 含义 | UV | 等级 | adapter |
|---|---|---|---|---|---|
| 1 | `options.data[]` | sheet 数组 | `IWorkbookData.sheets` + `sheetOrder` | 部分 | 按 `order` 生成 `sheetOrder`；UV 无顶层数组 |
| 2 | `name` | 表名 | `IWorksheetData.name` | 1:1 | 直拷 |
| 3 | `index` | sheet 业务 id | `IWorksheetData.id` | 部分 | 转字符串；LS `index`≠`order` |
| 4 | `order` | tab 顺序 | `sheetOrder` 下标 | 部分 | UV 无独立 order 字段 |
| 5 | `status` | 激活 0/1 | 无（运行时 focused unit） | 不可映射 | 回移仅一个 `status=1` |
| 6 | `hide` | 隐藏 0/1 | `hidden` BooleanNumber | 部分 | 0→FALSE，1→TRUE |
| 7 | `color` | tab 色 | `tabColor` | 1:1 | 直拷 |
| 8 | `row` / `column` | 行列数 | `rowCount` / `columnCount` | 1:1 | 直拷 |
| 9 | `defaultRowHeight` | 默认行高 px | `defaultRowHeight` | 1:1 | 直拷 |
| 10 | `defaultColWidth` | 默认列宽 px | `defaultColumnWidth` | 1:1 | 改名 |
| 11 | `scrollLeft` / `scrollTop` | 滚动 | 同名 | 1:1 | 直拷；运行时 `freezen` 不回写 |
| 12 | `zoomRatio` | 缩放 | `zoomRatio` | 1:1 | 直拷 |
| 13 | `showGridLines` | 0/1 | `showGridlines` | 部分 | 1→TRUE |
| 14 | `celldata[]` `{r,c,v}` | **仅初始化**稀疏格 | `cellData` 稀疏矩阵 | 部分 | 遍历写入 `cellData[r][c]`；`buildGridData` **优先已有 `data`**（`sheetmanage.js:721`） |
| 15 | `data[][]` | 运行时/持久主数据 | `cellData` | 部分 | 导出可 `getGridData` 回 `celldata`；协同写 `data` |
| 16 | `Store.flowdata` | 当前 sheet 内存副本 | 无 JSON 字段 | 不可映射 | **禁止当持久化契约**；`store/index.js:12` |
| 17 | `calcChain[]` | 公式链 | 无（engine-formula 依赖图） | 不可映射 | **回移必须保留**；`sheet.md` 声明有公式需带链 |
| 18 | `config.merge` / cell.`mc` | 合并 | `mergeData: IRange[]` | 部分 | `r_c` 键 ↔ start/end；UV 合并不在 cell 上 |
| 19 | `config.rowlen` / `columnlen` | 行高列宽 | `rowData[].h` / `columnData[].w` | 部分 | 索引对象 ↔ 稀疏数组 |
| 20 | `config.rowhidden` / `colhidden` | 隐藏行列 | `rowData[].hd` / `columnData[].hd` | 部分 | LS 用 `0` 占位 |
| 21 | `config.borderInfo[]` | 边框 | `IStyleData.bd` + workbook `styles` | 部分 | rangeType 拆到单元格样式 id |
| 22 | `config.authority` | 工作表保护 | core 无；需 range-protection 插件 | 不可映射 | 保留 LS JSON |
| 23 | `frozen` | 冻结语义 | `IWorksheetData.freeze` | 部分 | LS type 枚举 ↔ xSplit/ySplit |
| 24 | `freezen` | 冻结渲染缓存 | 无 | 不可映射 | 不持久化；`getAllSheets` 会 delete |
| 25 | `filter` / `filter_select` | 筛选 | `@univerjs/sheets-filter` | 部分 | 保留 LS；列模型不同 |
| 26 | `luckysheet_conditionformat_save` | 条件格式 | `resources[SHEET_CONDITIONAL_FORMATTING_PLUGIN]` | 部分 | 规则三层模型，见 CF adapter |
| 27 | `dataVerification` | 数据验证 | `resources[DATA_VALIDATION_PLUGIN]` | 部分 | type 对照表 |
| 28 | `chart[]` / `pivotTable` | 图表透视 | OSS 无；Pro 闭源 | 不可映射 | 保留 LS JSON，禁止抄 Pro |
| 29 | `image[]` | 浮动图 | drawing 插件 | 部分 | 保留 LS 结构 |
| 30 | `luckysheet_alternateformat_save` | 交替色 | 非 CF | 不可映射 | 独立模块，勿并入 CF |
| 31 | `IWorkbookData.styles` | — | 样式字典 | 部分（UV→LS） | 回移须 **inline 展开** 到 LS 扁平字段 |
| 32 | `IWorkbookData.resources` | — | 插件快照 | 部分 | 只还原 LS 认识的插件键 |
| 33 | `visibledatarow` / `ch_width` / `load` | 运行时几何 | skeleton | 不可映射 | 不回移 |
| 56 | `table[]` | 结构化表 | `@univerjs/sheets-table` | 部分 | `{id,name,range,columns,tableStyleId}`；LS 新增可选字段，旧 sheet 无此键 |
| 57 | `cell.note` / `sheet.notes` / `sheet.note` | Note | `@univerjs/sheets-note` | 部分 | **不替代 `ps`**；见 `docs/cell-note-vs-postil.md` |
| 58 | `definedNames[]` | 定义名称 | workbook named range | 部分 | 工作簿或 sheet 级可选；超链接 `DEFINE_NAME` 解析 |

## 单元格

| # | LS cell | 含义 | UV `ICellData` / `IStyleData` | 等级 | adapter |
|---|---|---|---|---|---|
| 34 | `v` | 原始值 | `ICellData.v` | 1:1 | bool 需同时写 `t` |
| 35 | `m` | 显示值 | **无** | 不可映射 | 回移用 `v` + `ct.fa` / `IStyleData.n.pattern` **重算**（`format.js`） |
| 36 | `f` | 公式串 | `ICellData.f` | 1:1 | 跳过 LS `spl` sparkline 伪公式 |
| 37 | `ct.fa` | 数字格式 | `IStyleData.n.pattern` | 部分 | `@`→`@@@`；`w`/`W` 万亿为 LS 扩展 |
| 38 | `ct.t` | g/n/s/d/inlineStr | `ICellData.t` + 可选 `p` | 部分 | inlineStr → `ICellData.p`（`IDocumentData`） |
| 39 | `bg` | 背景 | `IStyleData.bg.rgb` | 部分 | hex/rgb |
| 40 | `ff` | 字体 0–12 或名 | `IStyleData.ff` 字符串 | 部分 | 索引表互查 |
| 41 | `fc` | 字体色 | `IStyleData.cl` | 部分 | UV `cl` 是前景色，不是删除线 |
| 42 | `bl` / `it` / `fs` | 粗体/斜体/字号 | 同名 | 1:1 | BooleanNumber |
| 43 | `cl` | 删除线 0/1 | `IStyleData.st` | 部分 | LS `cl`≠ UV `cl` |
| 44 | `un` | 下划线 | `IStyleData.ul` | 部分 | UV 为装饰对象 |
| 45 | `ht` 0/1/2 | 水平对齐 | `IStyleData.ht` | 部分 | 0 居中 / 1 左 / 2 右 ↔ UV 枚举 |
| 46 | `vt` 0/1/2 | 垂直对齐 | `IStyleData.vt` | 部分 | 0 中 / 1 上 / 2 下 |
| 47 | `tb` 0/1/2 | 截断/溢出/换行 | `IStyleData.tb` WrapStrategy | 部分 | 0 CLIP / 1 OVERFLOW / 2 WRAP |
| 48 | `tr` / `rt` | 旋转档位/角度 | `IStyleData.tr` `{a,v}` | 部分 | 离散档 ↔ 角度 |
| 49 | `mc` | 合并锚点 | `mergeData` | 部分 | 见 #18 |
| 50 | `ps` | 批注 | sheets-thread-comment / sheets-note | 部分 | 产品上 Note≠批注；先保留 `ps` |
| 51 | `qp` | 数字当文本 | `t=FORCE_STRING` / `xf` | 部分 | quotePrefix |
| 52 | `spl` | sparkline | 无 | 不可映射 | 保留；migrate 应跳过当公式 |
| 53 | — | 富文本 | `ICellData.p` | 部分 | LS inlineStr `ct.s[]` ↔ document |
| 54 | — | 数组公式 | `ref` / `si` | 不可映射（LS→UV 单向） | LS 用 `dynamicArray`；回移勿丢 LS 字段 |
| 55 | `custom` | — | `ICellData.custom` | 部分 | 仅放无法建模的扩展 |

UV `ICellData` 定义见 `typedef.ts:248-295`：`p,s,v,t,f,ref,xf,si,custom`。无 `m`、无扁平 `bg/ff/ht`。

## 生命周期（adapter 必须遵守）

```text
初始化：celldata（或已有 data）→ sheetmanage.buildGridData → file.data
运行时：Store.flowdata ← setSheetParam(file.data)
持久化：file.data + config + calcChain + 插件 JSON
导出 API：getAllSheets / transToCellData → celldata
```

证据：`sheetmanage.js:715-749`、`1104-1105`、`1173`；`api.js:6677-6712`、`5908-5920`；`zh/guide/sheet.md:9-50`。

## 结论

不能做 1:1 对象替换。adapter 以 **LS JSON 为对外契约**：必须保住 `data`/`celldata` 双轨、`v/m/f/ct`、扁平样式、`config`、`calcChain`、筛选/冻结/CF/DV/图表透视图片。UV 侧样式进 `s` 或 styleId，格式进 `n.pattern`，插件进 `resources`。回移时 **重算 `m`**，**展开 styleId**，**不要把 `flowdata`/`freezen` 写入后端**。
