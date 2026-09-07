# Frontend UI Golden Spec

状态：设计参考已锁定，进入实现阶段（2026-09-07）

本文件是 `frontend-react/apps/web/**` 与 `frontend-react/packages/ui-system/**` 的实施基线。视觉参考为 [editor-reference.png](design/shared-kernel/editor-reference.png)，实际生成尺寸 1544×1024；实现几何以用户批准的 36/28/76/32/28px 与 340px shell 约束为准，不能因参考图输出尺寸改变批准尺寸。

## 1. 产品边界与架构判断

当前应用是 Vite + React 19 的单页应用，入口为 `apps/web/src/main.tsx`，路由由 `apps/web/src/app-routing.ts` 负责：

- `/`、`/workbooks`：Workbook Hub，拥有目录、创建、导入、筛选、搜索、回收站与云端状态展示。
- `/workbooks/:unitId`：单工作簿编辑器，唯一 `WorkbookSession` 由 `EditorRoute` 建立和释放。
- `/auth/callback`、`/auth/silent-renew`：OIDC 回调页。

视觉层只负责组合、焦点、菜单与临时 DOM 状态。业务选择、权限、保存状态、筛选结果、工作簿和图表对象必须继续由 `WorkbookSession`、command runtime、canonical model 与 persistence 链路提供。任何视觉改造不能在组件内创建第二份业务状态，也不能通过 UI 兜底修复 model 或 endpoint 错误。

编辑器调用链为：

```text
App -> WorkbookRouteGate -> EditorRoute
  -> useWorkbookSession
  -> useEditorCommandController
  -> EditorShell
     -> RibbonHost -> Ribbon -> RibbonLayoutRenderer
     -> FormulaBar / SheetCanvas / FeaturePanelHost / SheetTabs / StatusBar
  -> session.dispatch / session.dispatchUiSessionIntent
  -> command runtime -> model -> history/collaboration/persistence/render
```

工作簿中心调用链为：

```text
App -> WorkbookHubContainer -> WorkbookHubPage
  -> WorkbookTopBar / WorkbookSidebar / WorkbookCategoryTabs
  -> WorkbookSearch / WorkbookFilterDialog / WorkbookActionBar
  -> WorkbookGrid or WorkbookTable / WorkbookRowMenu
  -> WorkbookCatalogService -> local memory or remote cloud API
```

云端是唯一远程存储来源。登录状态、角色与同步状态必须从 `AuthProvider`、`WorkbookCatalogService` 和真实 commit/revision 状态 derive；不可通过视觉状态、预置 Demo 或 UI-only success 伪造。

## 2. 视觉 golden（设计图到达后锁定像素值）

用户已批准的结构尺寸：

| 区域 | 目标尺寸 | 视觉职责 |
|---|---:|---|
| 文档规格栏 | 36px | 文件名、保存/云端状态、账户与窗口级动作 |
| ribbon tabs | 28px | 文件入口、Home/Insert 等 tab；只显示当前 schema 顺序 |
| ribbon content | 76px | 按 schema 分组的命令面；窄屏整组折叠，不改变业务顺序 |
| formula bar | 32px | 名称框、公式输入、确认/取消与 IME 焦点 |
| sheet tabs | 28px | 工作表切换、添加、上下文菜单 |
| sidebar | 340px 默认 | 透视、筛选、图表、对象等 feature panel |
| sidebar resize | 300–480px | 统一 pointer gesture owner，拖动范围由 shell 约束 |

参考图的状态拓扑为：36px 文档规格栏（自有云表格 mark、工作簿名、已保存状态、撤销/重做、搜索、协作、评论、分享）、28px ribbon tabs、76px Home ribbon、32px formula bar、Canvas 工作区、340px Pivot 字段侧栏、28px sheet tabs 与状态栏。图中的绿色 E 品牌 mark 不作为产品资产，改用自有云表格图标；图中的销售分析表格只作为 golden fixture，不能作为生产 seed 或 mock 数据。

基础 token 由产品基线提供，设计图仅可细化明度、边框和状态层级：

```text
body: 13px
brand: #107C41
ink: #242424
surface: #F5F5F5
line: #D1D1D1
icon family: Fluent system icons
icon sizes: 16 / 20 / 24px
minimum interactive hit area: 28px
```

响应式约束：1024、1366、1920 三个宽度必须使用同一 ribbon schema。禁止整条 ribbon 横向滚动；空间不足时按组折叠为 schema 派生菜单或 gallery，组内顺序与命令归属不变。当前共享 `RibbonShell` 使用 1600px 以上 wide、1200–1599px compact、1200px 以下 narrow；compact/narrow 将完整组收纳到带 Fluent 图标的组菜单，命令仍全部可达。折叠规则放在共享 `RibbonShell`/layout contract，业务 renderer 不判断任意像素并私自重排。

设计图验收前要补充：原图绝对路径、原生尺寸、第一视口布局、所有可见文案、hover/pressed/selected/disabled/focus-visible 状态、响应式截图及至少五个像素对比点。实现前不得添加设计图没有的 badge、pill、统计数字、装饰卡片或新导航。

## 3. 组件 ownership 与收敛计划

### `packages/ui-system`

这是唯一承载业务页面视觉结构的共享组件边界。现有组件分为：

- 基础布局：`Box`、`Stack`、`Inline`、`Text`、`Heading`、`Divider`、`ScrollArea`、`Kbd`。
- 交互控件：`Button`、`TextInput`、`Textarea`、`Select`、`CheckToggle`、`RadioOption`、`DropdownMenu`、`ContextMenu`、`Dialog`、`Tabs`。
- 数据/状态：`DataTable`、`VirtualList`、`StatePanel`、`StatusBadge`、`Panel`、`TemplatePreview`。
- 编辑器 shell：`DesignerShell`、`RibbonShell`、`FormulaBar`、`SidebarShell`、`ScrollBar`、`RichTextInput`。
- 图标/文件：`Icon`、`AssetIcon`、`FileIcon`。

实施时先统一 token、语义尺寸和焦点契约，再调整业务页调用。页面组件不得散落原生 `button/input/select/textarea/table`；如需特殊语义，先在 ui-system 封装。`Box` 是结构原语，但不得用它绕过交互组件边界。

### `apps/web`

- `editor/`：shell 编排和业务边界；不持有重复模型状态。
- `components/HomeRibbon.tsx`、`InsertRibbon.tsx`、`RibbonLayoutRenderer.tsx`：仅从 canonical schema 与 session context 渲染命令。
- `components/SheetCanvas.tsx`：Canvas/render agent 负责渲染链，本任务不改 render-engine；DOM overlay 仍必须使用 ui-system。
- `workbooks/`：Hub 的页面组合与目录交互；状态通过 catalog/auth/persistence 读取。
- `components/panels/`、`components/pivot/`、`components/dialogs/`：feature 内容，不得自定义 shell 尺寸和颜色 token。

## 4. Icon inventory

### Fluent 资产（`apps/web/public/icons/fluent`）

已有 Fluent 24 regular 资产：

`align_left`, `align_center_vertical`, `align_center_horizontal`, `align_bottom`, `align_top`, `align_right`, `arrow_sort`, `border_all`, `calculator`, `chart_multiple`, `clipboard_paste`, `comment`, `copy`, `cube`, `cut`, `data_bar_vertical`, `data_line`, `data_pie`, `data_scatter`, `eraser`, `filter`, `flowchart`, `form`, `icons`, `image`, `link`, `number_symbol`, `paint_brush`, `paint_bucket`, `screenshot`, `search`, `shapes`, `signature`, `table`, `table_cells_merge`, `text_bold`, `text_color`, `text_font`, `text_font_size`, `text_italic`, `text_underline`, `text_wrap`, `timeline`。

用途约束：命令 ribbon 优先使用这些 Fluent 资产；`AssetIcon` 负责固定尺寸和替代文本。设计图若出现未覆盖 metaphor，先补齐 Fluent 资产或生产质量 SVG，再接入 schema；不得用 Unicode 字符替代 icon。

### 共享 `Icon` 名称族

`Icon.tsx` 当前维护约 100 个语义图标，覆盖：导航（`home`, `menu`, `arrow-*`, `chevron-*`）、文件/持久化（`file-*`, `folder*`, `save`, `upload`, `download`, `cloud-check`）、编辑（`undo`, `redo`, `copy`, `scissors`, `clipboard`, `paint-*`, `eraser`）、文字/格式（`bold`, `italic`, `underline`, `strikethrough`, `type`, `percent`, `comma`, `dollar-sign`, `sigma`）、表格/筛选（`table*`, `grid`, `rows`, `columns`, `filter*`, `sort`, `freeze`, `merge-cells`）、对象/图表（`chart-*`, `picture`, `shape-*`, `textbox`, `sparkline`, `timeline`, `gantt-sheet`, `report-sheet`）、状态与操作（`check*`, `alert-circle`, `info`, `help`, `loader`, `lock`, `eye`, `settings`, `search`, `plus`, `x`, `trash`, `more-*`, `share`, `users`）。

统一规则：默认 16px（compact）、20px（shell/action）、24px（large ribbon tile）；颜色来自 token；使用 `currentColor`；所有可交互图标必须有 accessible name、至少 28px 命中区域和 focus-visible 状态。`Icon` 适合语义 SVG，`AssetIcon` 适合已提交的 Fluent 文件，不再新增第三套图标路径。

## 5. 当前问题与实施边界

只读盘点发现：

1. `DesignerShell` 当前使用 `RIBBON_DENSITY` 的 167/48/29/22px 旧几何；需在统一 token 和 schema 下改为本 spec 的 36/28/76/32/28px。
2. `RibbonShell` 目前保留宽度模式和 tab strip 行为，必须改为按 schema 的组折叠；不可在 1024/1366/1920 通过整条 ribbon 横滚隐藏命令。
3. `SidebarShell` 默认 340px，但现有限制 320–360px；应统一为 300–480px，并保留 pointer capture 和真实宽度状态。
4. `styles.css` 存在全局 `Microsoft YaHei`/`Noto Sans SC` 与旧 designer/home CSS 变量、多处 media query 局部字号覆盖；需迁移到 Tailwind token，删除 global font override 与局部 hack。
5. `Button`、`TextInput`、`Select` 等已有共享封装，但部分组件仍在 ui-system 内部直接使用原生元素；这部分可保留为共享封装内部实现，页面业务层必须继续使用共享组件。
6. Home/Insert 命令数量和 domain handler 已存在，视觉改造不可删除 routes、commands、panel 或真实错误路径；schema 是唯一命令顺序和分组来源。

不在本子任务范围：`frontend-react/packages/render-engine/**`；后端 endpoint、runtime API 契约和 persistence 语义由对应 agent 负责。发生契约变化时必须以其实际导出和协议为准，不猜 endpoint。

## 6. 交互与可访问性验收清单

- 公式编辑：Enter 提交并按设置移动，Esc 恢复，Tab 提交并移动，IME composition 不提前提交；焦点只由 `FormulaBar`/session 编辑链拥有。
- 透视区：search、group menus、defer/loading、错误状态都由真实 feature state 驱动；没有空结果静默降级。
- 筛选：menu 使用 `VirtualList`，选中/全选/清除/应用状态来自 canonical filter resolver；隐藏行只是 projection。
- 图表与绘图对象：context menu、selection、rename、z-order、visibility 通过 session/command，不能靠 overlay 自己改状态。
- 全部 tabs、menus、dialog、resize handle 支持键盘焦点和 28px 最小命中区；焦点顺序稳定，disabled reason 可读。
- 文案从 `i18n`/locales 进入，不在新视觉组件中硬编码中文或英文。

## 7. 后续执行顺序

1. 读取并锁定 ImageGen 设计图：补齐路径、原生尺寸、第一视口与状态 golden。
2. 在 `ui-system` 先建立统一 Tailwind token、shell geometry、icon sizing、focus/hit-area primitives。
3. 收敛 `DesignerShell`、`RibbonShell`、`FormulaBar`、`SidebarShell`、tabs/status 结构。
4. 按 schema 更新 Home/Insert/各 contextual ribbon 与 Hub 组件，保持原 command/session 链。
5. 删除旧全局字体和局部视觉 hack，更新组件与测试契约。
6. 统一阶段完成后再执行 typecheck、边界检查、Playwright/IAB 真实交互与设计图对比；不在中途频繁编译。
