# Univer Pro 打印参考映射

> 只读参考包落盘于 `.univer-temp/`（gitignore），**禁止**将 Pro `lib/*.js` 提交进 `src/`。

## 参考包版本

| npm 包 | 版本 | 解压目录 |
|---|---|---|
| `@univerjs-pro/sheets-print` | 0.25.1 | `.univer-temp/sheets-print-ref/package/` |
| `@univerjs-pro/print` | 0.25.1 | `.univer-temp/print-ref/package/` |

## Pro 模块 → LuckySheet 自研模块

| Pro（types / 职责） | LS 实现 |
|---|---|
| `SheetPrintManagerService` | [`printManager.js`](../frontend/src/expendPlugins/print/printManager.js) |
| `SheetPrintCanvasView` | [`printRenderer.js`](../frontend/src/expendPlugins/print/printRenderer.js) |
| `SheetPrintingResourceCollector` | [`printResources.js`](../frontend/src/expendPlugins/print/printResources.js) |
| `@univerjs-pro/print` `createPrintStyle` | [`printBrowser.js`](../frontend/src/expendPlugins/print/printBrowser.js) |
| Print UI / `SheetPrintView` | [`printDialog.js`](../frontend/src/expendPlugins/print/printDialog.js) |
| `SheetScreenShotOperation` | [`printScreenshot.js`](../frontend/src/expendPlugins/print/printScreenshot.js) |
| PDF 导出 | [`printPdf.js`](../frontend/src/expendPlugins/print/printPdf.js) |
| Facade `FWorkbook.*` | [`print.js`](../frontend/src/expendPlugins/print/print.js) + [`api.js`](../frontend/src/global/api.js) |
| 事件 `BeforeSheetPrintOpen` 等 | [`printEvents.js`](../frontend/src/expendPlugins/print/printEvents.js) |
| 枚举 / 布局映射 | [`printLayout.js`](../frontend/src/expendPlugins/print/printLayout.js) |

## 公开类型对照

- `ISheetPrintLayoutConfig` → `file.config.printLayout` + `printoptions`
- `ISheetPrintRenderConfig` → `file.config.printRender` + `printoptions.printOptions`
- `PrintHeaderFooterSymbol` → `printLayout.resolveHeaderFooterText`
- `IUniverSheetsPrintConfig.enforceWatermark` → 插件 `plugins:[{name:'print',config:{enforceWatermark}}]`

## 差距声明

- LS 使用离屏 canvas + `visibledatarow/column`，非 Univer Scene 渲染
- 本 fork 不做 Pro license 页数限制（`DEFAULT_PRINT_LIMIT`）
- Excel 像素级分页仍可能有偏差
