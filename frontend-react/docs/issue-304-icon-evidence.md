# Issue #304 Ribbon 图标证据与实现契约

本文件把 HOME/INSERT Ribbon 的图标选择从“看起来像”收敛为可追溯的官方语义契约。图标不承载业务逻辑；命令 ID、状态和错误仍由 `ui-command-catalog` 与 `CommandRuntime` 负责。

## 官方依据

- [Microsoft Office Add-in icon guidelines](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-icons)：Ribbon 图标使用简洁、清晰、单色 monoline 视觉；核心图形不超过必要元素；16/20/32 px 是必须覆盖的基准尺寸，modifier 放在右下角，不使用渐变、纹理或光源。
- [Fluent 2 iconography](https://fluent2.microsoft.design/iconography)：命令栏使用 Fluent system icons；regular 用于普通可用动作，filled 用于选中状态或需要更强可读性的状态；图标按字面隐喻命名，modifier 只用于补充语义。
- [Microsoft Fluent System Icons](https://github.com/microsoft/fluentui-system-icons)：Fluent system icons 的官方开源集合，提供 regular/filled 变体和 RTL direction metadata。
- [SpreadJS Home Tab](https://developer.mescius.com/spreadjs/docs/spreadjs-designer-component/designerinterface/spdesigntab/spdesignhometab)：HOME 的产品分组为 Clipboard、Fonts、Alignment、Number、Styles、Cells、Editing，并提供 Format Cells 对话框启动器。
- [SpreadJS Insert Tab](https://developer.mescius.com/spreadjs/docs/spreadjs-designer-component/designerinterface/spdesigntab/spdesigninserttab)：INSERT 覆盖 tables、slicers、pictures、sparklines，并在插入对象后切换到相应 Design 上下文。

## Ribbon 语义映射

实现使用项目的 `AssetIcon` 入口；每个 asset 的命名应保持 Fluent 的字面隐喻，不以业务命令名称创造第二套图标语义。尺寸按 Ribbon 控件落点固定为：大按钮 24 px、小按钮 16 px、菜单项 16 px、状态 modifier 12 px。图标颜色继承 currentColor，选中状态切换 Fluent filled 变体或由控件的 selected token 表达。官方 regular SVG 已 vendor 到 `apps/web/public/icons/fluent/`，不依赖运行时远程加载。

当前使用的 official assets 覆盖表格、表单、图片、形状、图标、3D cube、SmartArt flowchart、截图、筛选器、时间线、链接、评论、签名，以及 HOME 的剪贴板/字体/对齐/合并/换行/搜索图标。

### HOME

| Surface | 官方字面隐喻 | 视觉状态 |
| --- | --- | --- |
| clipboard.paste / pasteSpecial | clipboard-paste | regular；下拉 chevron 是独立 modifier |
| clipboard.cut / copy | scissors / copy | regular |
| clipboard.formatPainter | paint-brush | regular |
| font.family / size | text-font / text-size | regular |
| font.bold / italic / underline | text-bold / text-italic / text-underline | regular；激活时 filled |
| font.borders / fill / color | border-all / paint-bucket / text-color | regular |
| alignment.top/middle/bottom | align-vertical-* | regular |
| alignment.left/center/right | align-horizontal-* | regular |
| alignment.wrap / merge | text-wrap / table-cells-merge | regular |
| number.format / currency / percent | number-format / money / percent | regular |
| styles.conditional / table / cell | format-painter / table / cell-style | regular |
| cells.insert/delete/format | table-insert / delete / table-properties | regular |
| editing.autoSum / fill / clear | calculator-sum / arrow-fill / eraser | regular |
| editing.sort/filter/search | arrow-sort / filter / search | regular |

### INSERT

| Surface | 官方字面隐喻 | 视觉状态 |
| --- | --- | --- |
| tables.pivot / recommendedPivot / table / forms | pivot-table / lightbulb / table / form | regular |
| illustrations.picture / shape / icons / models3d | image / shapes / icons / cube-3d | regular |
| illustrations.smartArt / screenshot | diagram / screenshot | regular |
| controls.checkbox | checkbox | regular；checked 使用 filled |
| charts.recommended / gallery / pivot | lightbulb / chart / pivot-chart | regular |
| sparklines.gallery | sparkline | regular |
| filters.slicer / timeline | filter / timeline | regular |
| links.hyperlink | link | regular |
| insertComments.threaded | comment | regular；有未读状态才使用 filled modifier |
| text.textbox / headerFooter / wordArt | textbox / page-header / word-art | regular |
| text.signatureLine / object | signature / object | regular |
| symbols.equation / symbol | function / omega | regular |

## 禁止事项

1. 禁止用 emoji、文字首字母、渐变或临时手写 SVG 代替官方图标。
2. 禁止因为图标资源缺失而改变命令语义、吞掉错误或插入假对象；资源缺失必须在构建/验收阶段显式失败。
3. 图标大小、颜色和间距只由 Ribbon token/Tailwind class 控制，业务组件不得散落 magic number。
4. RTL 方向按 Fluent metadata 处理；需要镜像的图标不得复制出第二个业务别名。

## 验收证据

最终验收需同时保留：HOME 1783 px 与 INSERT 1905 px 的浏览器截图、DOM surface 顺序、控制台/网络无新增错误，以及本文映射表与实际 asset 文件名的一致性。
