# LuckySheet 与 Univer 公开打印语义映射

本实现仅参考 [Univer Print 公开文档](https://docs.univer.ai/guides/sheets/features/print) 与 [OSS print interceptor](https://github.com/dream-num/univer/blob/dev/packages/sheets-ui/src/services/print-interceptor.service.ts) 的行为契约；没有读取、引入或反编译任何 `@univerjs-pro/*` 包。

## 官方公开 layout 映射

| 字段 | LuckySheet 存储与行为 | 状态 |
|---|---|---|
| `area` | `CurrentSheet`、`Workbook`、`CurrentSelection`、`AllSelection`；多选区保持为独立目标 | 已实现 |
| `subUnitIds` | `sheet.index`，可带 `range` | 已实现 |
| `paperSize` / `pageSizeCustom` | `pageSetup.paperSize`、`paperWidth`、`paperHeight` | 已实现 |
| `direction` / `scale` / `customScale` | `pageSetup.orientation`、`fitToWidth`、`fitToHeight`、`scale` | 已实现 |
| `freeze` | `PrintTitles.row/column`，每页重复标题并扣除分页空间 | 已实现 |
| `margin` / `maxRowsEachPage` / `maxColumnsEachPage` | `pageMargins` 与分页覆盖项 | 已实现 |

## 官方公开 render 映射

| 字段 | LuckySheet 行为 | 状态 |
|---|---|---|
| `gridlines`、`hAlign`、`vAlign` | Canvas 单元格绘制 | 已实现 |
| `headerFooter`、六个 header/footer 槽位 | 页码、总页数、簿名、表名、日期、时间替换 | 已实现 |
| `draft` | 跳过图片和图表 bitmap | 已实现 |
| `watermark` | 不实现 | 明确排除 |

## OSS 可验证语义

- 范围、组件和 DOM 扩展点映射为有序 `PrintTarget[]`、页面渲染器和实例 portal。
- `PrintResourceCollector` 支持 `add(Promise)` 并在 10 秒内持续排空后来注册的资源；图片 decode 与图表 canvas 均先等待再出页。
- 每次打印创建不可变 `PrintSession`，冻结实例、sheet、范围、轴、配置和资源引用；焦点切换不会改变在途打印。

## 明确不支持

- Pro 授权水印、截图授权、PDF 直出、图表矢量输出、Pro 内部算法。
- LuckySheet `print()` 是异步 `Promise<PrintResult>` 扩展；`openPrintDialog()` 和 `closePrintDialog()` 仍是 focused 实例操作。
