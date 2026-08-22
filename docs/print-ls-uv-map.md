# 打印 layout / render 字段映射

> 对照 [Univer Print 公开文档](https://docs.univer.ai/guides/sheets/features/print) 与 OSS [`print-interceptor.service.ts`](../../univer/packages/sheets-ui/src/services/print-interceptor.service.ts)。
> **禁止**引入 `@univerjs-pro/sheets-print` / `@univerjs-pro/print` 源码或 npm `lib/*.js`。
> 本 fork 在 `frontend/src/expendPlugins/print/` 自研，字段仍落 LuckySheet `file.config.printoptions`。

## 1. Facade 能力面

| Univer 公开 API | LuckySheet 落点 | 本 fork 行为 |
|---|---|---|
| `FWorkbook.openPrintDialog` | `luckysheetPrint.createDialog()` / `luckysheet.openPrintDialog()` | 打开预览+设置对话框，**不要求 license** |
| `FWorkbook.closePrintDialog` | `luckysheetPrint.closeDialog()` / `luckysheet.closePrintDialog()` | 关闭对话框与预览层 |
| `FWorkbook.updatePrintConfig` | `luckysheetPrint.updatePrintConfig` | 写 `printoptions` + 规范化 `printLayout` |
| `FWorkbook.updatePrintRenderConfig` | `luckysheetPrint.updatePrintRenderConfig` | 映射网格线/对齐/页眉页脚 |
| `FWorkbook.print` | `luckysheetPrint.print()` / `luckysheet.print()` | 预览页 + `window.print()` |
| interceptor `PRINTING_RANGE` | `collectPrintRange` / `resolvePrintRange` | 当前表 / 选区 / `PrintArea` |
| interceptor `resourceCollector.wait` | `waitPrintResources` | `draft≠1` 时等图表 canvas；超时 10s |
| `enforceWatermark` / 未授权水印 | **不做** | 本 fork 无 Pro 配额水印 |
| `saveScreenshotToClipboard` / `getScreenshot` | **不做** | 公开文档标明需商业许可 |

## 2. `ISheetPrintLayoutConfig` → LuckySheet JSON

规范化副本另存 `file.config.printLayout`（只读对照）；权威存储仍是 Excel 风 `printoptions`。

| Univer | 枚举/类型 | LuckySheet `printoptions` | 规范化 `printLayout` |
|---|---|---|---|
| `area` | `CurrentSheet` / `CurrentSelection` / `Workbook` / `AllSelection` | 缺省当前表；`PrintArea` 有值则用之；对话框可选「选中区域」 | `area` |
| `subUnitIds` | sheetId + 可选 range | 当前 `file`；范围见 `PrintArea` | `sheetIndex` + `range` |
| `paperSize` | `A4` / `Letter` / … | `pageSetup.paperSize`（Excel 代码：9=A4, 1=Letter…） | `paperSize` 字符串 |
| `pageSizeCustom` | `{w,h}` | `pageSetup.paperWidth` / `paperHeight` | `pageSizeCustom` |
| `direction` | `Portrait` / `Landscape` | `pageSetup.orientation`：0 默认, 1 landscape, 2 portrait | `direction` |
| `scale` | `Origin` / `FitWidth` / `FitHeight` / `FitPage` / `Custom` | `pageSetup.scale` + `fitToWidth` / `fitToHeight` | `scale` + `customScale` |
| `customScale` | number | `pageSetup.scale`（10–400） | `customScale` |
| `freeze` | `Row` / `Column` | `PrintTitles.row` / `PrintTitles.column` | `freeze` |
| `margin` | `Normal` / `Narrow` / `Wide` / `None` | `pageMargins`（inch） | `margin` + 数值 |
| `maxRowsEachPage` / `maxColumnsEachPage` | number | 无对应字段；分页算法内部使用 | 可选覆盖 |

### 纸张代码

| Excel `paperSize` | Univer `PrintPaperSize` | mm |
|---|---|---|
| 1 | Letter | 215.9 × 279.4 |
| 3 | Tabloid | 279.4 × 431.8 |
| 5 | Legal | 215.9 × 355.6 |
| 6 | Statement | 139.7 × 215.9 |
| 7 | Executive | 184.15 × 266.7 |
| 8 | A3 | 297 × 420 |
| 9 | A4 | 210 × 297 |
| 11 | A5 | 148 × 210 |
| 12 | B4 | 250 × 353 |
| 13 | B5 | 176 × 250 |

像素换算：`96 DPI`，`1mm = 96/25.4 px`。

## 3. `ISheetPrintRenderConfig` → LuckySheet JSON

| Univer | LuckySheet | 规范化 `printRender` |
|---|---|---|
| `gridlines` | `printOptions.gridLines`（0/1） | `gridlines` boolean |
| `hAlign` `Start/Middle/End` | `printOptions.horizontalCentered` | `hAlign` |
| `vAlign` `Start/Middle/End` | `printOptions.verticalCentered` | `vAlign` |
| `headerFooter` 占位符 | `headerFooter` 既有结构 + 字符串拼接 | `headerFooter` 数组 |
| `headerFooterSetting` | `pageMargins.header/footer` + 左中右字符串 | `headerFooterSetting` |
| `watermark` | **不映射** | 忽略 |

## 4. OSS interceptor 语义（可学不可抄 Pro）

`packages/sheets-ui/src/services/print-interceptor.service.ts`：

- `PRINTING_RANGE`：打印范围可被拦截替换。本实现用 `resolvePrintRange(area)` 固定三种来源。
- `SheetPrintingResourceCollector.wait(timeout=10000)`：异步资源（图）完成后再出页。本实现等 `.luckysheet-data-visualization-chart canvas` 就绪。
- `PRINTING_COMPONENT_COLLECT` / `PRINTING_DOM_COLLECT`：Univer 场景图组件收集。LuckySheet 无 Scene，改为离屏 canvas 自绘。

## 5. 打印管线与差距

1. 显示值只打 `m`，缺省回退 `v`，**不打公式串 `f`**。
2. 分页：`visibledatarow` / `visibledatacolumn` + 纸张可用区。
3. 合并格跨页：**裁切绘制**（整格不拆数据，画布按页 clip）。未做「整格推下一页」的重排。
4. `pageSetup.draft=1` 跳过图片与图表。
5. 优先离屏 canvas 自绘（与网格行高列宽一致）；`html2canvas` 仅作后备。
6. **不做**：PDF 直出、图表矢量、页边距与 Excel 像素级对齐、Pro 水印/截图授权。

## 6. Store / API 注意

- 打印 API 打在 **focused** 实例。
- 对话框挂 `document.body` 并带 `data-ls-instance`，避免双实例抢同一 dialog id。
- 禁止 `const { flowdata } = Store` 后长期持有；Proxy 只在属性访问时指向当前实例。
