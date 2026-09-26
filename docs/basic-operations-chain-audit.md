# 基础操作与图表：全使用链审查及集中整改

## 基线、范围与真实状态

- 2026-09-26 通过 GitHub Connector 重新读取 main：`a2a6140a90351b38f1e6f5fbc167d09b4f6ecc7f`，与本地 main 相同。
- 当前审查工作分支：`codex/structural-reference-integrity`；本轮复核基线 `origin/main=a2a6140a90351b38f1e6f5fbc167d09b4f6ecc7f`，本轮代码改动从 PR head `4a7f77bc5542be465c109b15c60e7ecf52ff4bec` 继续。下面的行号以当前 PR head 的源码为准，不把历史 main 或旧 PR 文字当成当前实现证据。
- 唯一交付 PR：[草稿 #345](https://github.com/JunoBarnswell/luckysheet-local/pull/345)。本轮结构事实、sheet snapshot、chart geometry 和 snapshot owner identity 修复已分别提交；稀疏单元格规划、extent 和 cell-shift history 改动保留在同一 PR 待整体复核。
- 最新用户要求：以整个使用链为单位；至少 40 个操作一起审查、集中修复并复核；**最后实测**。早先“仅静态、不测试”只约束前置审查阶段，不再替代最终运行验收。
- **已确认的产品取舍**：2026-09-26 用户选择“统一由 Java 服务规划，可要求服务在线”。结构操作不再承诺无服务的浏览器离线执行；连接失败时拒绝提交并保留编辑草稿。此授权只改变结构规划归属，不自动扩展为全部普通输入或图表样式都必须远程。
- 建档 60 个操作，操作数不等于缺陷数。下表是追踪清单，**不是 60 项审查完成/通过的声明**。同根因的行列、图表类型、入口变体不重复算 bug。
- 审计底稿形成时只做静态阅读。后续仅有下列明示的定向回归证据；上一已提交 head 的 CI 成功不能作为当前工作树验收。
- 已完成的历史定向测试：core-model 59 项、Java structural facts 10 项、chart layout 15 项及 Canvas drawing 13 项。本轮新增 Java owned-patch 回归尚未执行：环境只有 JDK 17、无 Maven wrapper/可用 `mvn`，项目要求 Java 21。前端项目级 `tsc` 仍因工作区缺少已声明的 `@types/react` 失败，本轮 5 个改动 TS/TSX 文件均无目标诊断。已在 in-app browser 打开 Vite 页面并查 Console/Network；页面因 `GET /api/auth/config → 500` 停在认证配置错误，Console 无 JS error/warn，未能进入工作簿或做图表 UI 操作。完整门禁、性能和 Excel 验收未完成。
- Java `StructuralStateChanges` 目前只提供事实载体和历史迁移捕获/回放，**尚未接入在线结构意图规划和提交**。当前客户端先做 TS 结构变换，Java 再执行 reducer，远端客户端仍按 intent 重放；唯一 Java planner 尚未达成。

## 审查单位与证据要求

每个操作必须走完：真实 UI 入口 → 参数/选区解释 → 命令/权限 → canonical 计划及模型 → 引用/计算 → projection/render/hit → undo/redo → remote/server → save/reload → OOXML。

每项问题记录：最小触发条件、源码位置及调用关系、唯一状态所有者、影响对象、成功/拒绝路径、复杂度、统一修复、实际验证证据。只有定位入口或找到测试名称时，状态保持“待审”；没有实现或不能互操作的类型不能记为支持。

## 60 个操作清单

状态：`贯通` = 已阅读主要调用链但仍需缺陷修复/六轮复核/实测；`局部` = 已阅读部分共用链；`待审` = 尚不能给完成结论。所有行的最终验收均未完成。

| ID | 用户操作 | 真实入口 / 主要所有者 | 静态状态 |
| --- | --- | --- | --- |
| B01 | 普通值输入与提交/取消 | `WorkbookSession` / cell-edit / `sheet.cell.set` | 局部 |
| B02 | 公式输入及跨表取值、首个公式 | `runtime.ts` / FormulaEngine | 局部 |
| B03 | 设置数字格式与单元格样式 | `sheet.style.set` / `style.set` | 局部 |
| B04 | 清除内容 | `range.clear` / `clear-planner.ts` | 局部 |
| B05 | 清除格式 | `range.clear` / `clear-planner.ts` | 局部 |
| B06 | 复制并粘贴 | `sheet.range.paste` / `range.paste` | 局部 |
| B07 | 剪切并粘贴 | `sheet.range.move` / `range.move` | 局部 |
| B08 | 选择性粘贴 | `sheet.range.paste` / clipboard plan | 局部 |
| B09 | 拖动填充柄/序列填充 | `sheet.range.fill` / `fill.applied` | 局部 |
| B10 | 拖动选区边框移动 | interaction / `sheet.range.move` | 局部 |
| B11 | Ctrl+拖动复制 | interaction / clipboard plan | 局部 |
| B12 | 插入整行 | `sheet.rows.insert` / `rows.inserted` | 贯通 |
| B13 | 删除整行 | `sheet.rows.delete` / `rows.deleted` | 贯通 |
| B14 | 插入整列 | `sheet.columns.insert` / `columns.inserted` | 贯通 |
| B15 | 删除整列 | `sheet.columns.delete` / `columns.deleted` | 贯通 |
| B16 | 插入单元格、下移 | `applyCellShift` / `cells.inserted` | 贯通 |
| B17 | 插入单元格、右移 | 同上，column axis | 贯通 |
| B18 | 删除单元格、上移 | `applyCellShift` / `cells.deleted` | 贯通 |
| B19 | 删除单元格、左移 | 同上，column axis | 贯通 |
| B20 | 排序及重排 | `data.sort.*` / `rows.permuted` | 局部 |
| B21 | 筛选、清除条件及重应用 | `sheet.autoFilter.*` / resolved owner | 局部 |
| B22 | 合并与取消合并 | `sheet.merge.*` | 局部 |
| B23 | 隐藏/取消隐藏行列 | rows/columns visibility mutations | 局部 |
| B24 | 冻结及取消冻结窗格 | `sheet.freeze.set` / PaneMap | 局部 |
| B25 | 多 Sheet 切换及懒加载 | WorkbookSession / ProjectionRuntime / CellMatrix | 局部 |
| B26 | Sheet 重命名 | `sheet.rename` / ReferenceTransformDomain | 局部 |
| B27 | 删除 Sheet | `sheet.remove` / reference owners | 局部 |
| B28 | Table 扩缩及表内行列编辑 | `sheet-table-commands.ts` / structural commands | 局部 |
| B29 | 撤销 | CommandRuntime `undo → preflightHistory → applyHistory` | 贯通 |
| B30 | 重做 | CommandRuntime `redo → preflightHistory → applyHistory` | 贯通 |
| B31 | 保存并重新打开 | checkpoint / journal / WorkbookOperationService | 局部 |
| B32 | 远端操作重放及版本冲突 | CollaborationSession / server commit / replay | 贯通 |
| B33 | XLSX 导入、编辑、导出、重导入 | exchange-excel-ooxml / artifact ownership | 局部 |
| C01 | 创建柱形图及子类型 | `chart.insert.column` / chart layout / native codec | 局部 |
| C02 | 创建条形图及子类型 | `chart.insert.bar` / 同一图表链 | 贯通 |
| C03 | 创建折线图及子类型 | `chart.insert.line` / 同一图表链 | 局部 |
| C04 | 创建面积图及子类型 | `chart.insert.area` / 同一图表链 | 局部 |
| C05 | 创建饼图/复合饼图 | `chart.insert.pie` / 同一图表链 | 贯通 |
| C06 | 创建圆环图 | `chart.insert.doughnut` / 同一图表链 | 局部 |
| C07 | 创建 XY 散点图 | `chart.insert.scatter` / X/Y role bindings | 贯通 |
| C08 | 创建气泡图 | `chart.insert.bubble` / X/Y/Size role bindings | 贯通 |
| C09 | 创建树状图 | `chart.insert.treemap` / hierarchy / native codec | 贯通 |
| C10 | 创建旭日图 | `chart.insert.sunburst` / hierarchy / native codec | 贯通 |
| C11 | 创建直方图 | `chart.insert.histogram` / binning / native codec | 贯通 |
| C12 | 创建帕累托图 | `chart.insert.pareto` / aggregation / native codec | 局部 |
| C13 | 创建箱线图 | `chart.insert.box-whisker` / statistics / native codec | 局部 |
| C14 | 创建瀑布图 | `chart.insert.waterfall` / accumulations / native codec | 局部 |
| C15 | 创建漏斗图 | `chart.insert.funnel` / stages / native codec | 局部 |
| C16 | 创建股价图及成交量子类型 | `chart.insert.stock` / role bindings / native codec | 贯通 |
| C17 | 创建曲面/等高线图 | `chart.insert.surface` / geometry / native codec | 贯通 |
| C18 | 创建雷达图 | `chart.insert.radar` / signed value axis | 贯通 |
| C19 | 创建地图图表 | `chart.insert.map` / validated geographic resource | 局部 |
| C20 | 创建组合图及次坐标轴 | `chart.insert.combo` / per-series type/axis | 贯通 |
| C21 | 图表选数据、切换行列、编辑系列 | `chart.selectData` / `chart.series.*` | 局部 |
| C22 | 更改图表类型及子类型 | `chart.setType` / `retargetChartPayload` / editor draft | 贯通 |
| C23 | 修改轴、标题、图例、数据标签/数据表 | `chart.setElements` / chart payload / layout | 局部 |
| C24 | 趋势线、误差线与预测 | `chart.setTrendlines` / `chart.setErrorBars` | 贯通 |
| C25 | 图表系列/类别筛选及隐藏数据 | chart payload / `scalarVector` / projection | 局部 |
| C26 | 选中图表元素、拖动、缩放 | `chartHitTest` / `drawing.transform` | 局部 |
| C27 | 删除图表并撤销恢复 | `chart.remove` / drawing add/remove | 局部 |

每种图表都附带同一纵向验收，不把“类型登记”代替可用性：创建 → 修改源值 → 改系列/类型 → 命中选择 → 行列插删 → undo/redo → server/reload → XLSX/Excel。

## 当前已从源码贯通的公共链

| 环节 | 当前源码证据 | 可得出的结论 |
| --- | --- | --- |
| UI/权限 | `apps/web/src/editor/EditorDialogHost.tsx:130`；`workbook-session.ts:1555,3865,891,2208` | shift dialog 调用 session，dispatch 解析选区，command permission 与 mutation guard 分开检查；远程模式离线写入被拒绝。 |
| 本地结构命令 | `sheet-features/src/index.ts:2504,2536,2577,2609`；`editing/index.ts:1743` | axis 与 cell-shift 都进入 StructuralTransform，但 history 自行构造 inverse intent/cell snapshot。 |
| 模型/引用 | `core-model/src/structural-transform.ts:545,651,671` | TS 仍拥有地址/引用结构算法；内部预计划没有成为完整 wire/history facts。 |
| 计算 | `spreadsheet-app/src/runtime.ts:485,927` | 消费 StructuralTransformResult 的 clear/populate/rule/name effects；普通轴编辑未默认重建 FormulaEngine，但 table sync、audit、history preview 仍有独立成本。 |
| projection | `features/projection/projection-runtime.ts:61,138` | 结构 mutation ID 触发 projection domain 失效；chart source index 另处理 reference owner 变化，不是完整统一 ProjectionImpact。 |
| history | `command-runtime/src/index.ts:1242,1414,1467` | 每次 undo/redo 创建 workbook snapshot 和 preview runtime，preview 与 live 各执行一次 reducer，再覆盖部分公式/name facts。 |
| 协作/服务端 | `collaboration-session.ts:193,209`；`WorkbookOperationService.java:146,153`；`StructuralMutationDescriptor.java:84` | 客户端意图由 Java 再推导；远端客户端 handler 再推导，随后消费 partial patch。exact-base/conflict 已存在，不能误报为没有版本检查。 |
| 协议 | `protocol/src/index.ts:82`；`server/contract/StructuralPatch.java:14` | v3 只含 formula/name deltas，不含完整 cells/metadata/extents；不能精确回放所有结构事实。 |
| 图表 | `features/chart/commands.ts:519`；`core-model/src/domain.ts:703,785` | 已登记 20 家族及子类型，不能推断实现/互操作完整，也不能把这些类型全部算成“缺少”。 |
| 图表投影/绘制/命中 | `features/chart/data.ts:317`；`features/chart/layout.ts:662`；`drawing-renderers.ts:1152,1363,1428` | layout 提供部分事实，但多个 special charts 在绘制和 hit-test 各自重建坐标/统计，存在双重几何所有权。 |
| 图表服务器写入 | `DrawingMutationDescriptor.java:29,40` | affectedRanges 先调用 apply 深拷贝全 workbook；真正 apply 再处理一遍。 |
| 原生文件 | `exchange-excel-ooxml/src/export.ts:22,50,70`；`native-chart.ts:265,303,482` | 未知/无 owner chart 对修改后保存已有 fail-close；已宣称 owned 的 chart codec 仍须核实实际 OOXML schema。 |

路径前缀：TS 包表格未写前缀者均在 `frontend-react/packages/`；`workbook-session.ts` 在 spreadsheet-app/src；drawing renderer 在 `frontend-react/apps/web/src/components/canvas/`；Java 在 `backend/src/main/java/com/xc/luckysheet/server/`。

## 已证实问题与根因分组（尚未全部修复）

### S：结构事实及事务所有权

1. **S1 extent 双重增长**：原 CellMatrix.shiftRows/shiftColumns 的 set 会经 onWrite 扩大 SheetExtent，applyAxis 随后 `+= count` 再扩大。尾行/尾列已有值时可多增长、在边界写后失败。本地改动使用规划前 extent 与明确 nextCount，并移除 storage 的独立 shift 算法；待整批验收。
2. **S2 写入失败出现在清除之后**：cell-shift 原 extractRegion 先删除，再逐 cell normalize/set；后面的非法字体可令前面的写入已生效。本地计划先准备全部幸存者，删除项独立复制，提交只消费目标坐标；未宣称目标 sheet 无 hydration。
3. **S3 history 扫描矩形面积且校验方向错误**：cell-shift snapshot/restore 双层循环枚举 implicit cells；restore 按原 insert intent 预检，合法末端插入后撤销被当成再次插入拒绝。本地已改稀疏枚举及真实逆向检查；Java restore 本来就检查 inverse，不能重复修“同一个 Java bug”。
4. **S4 不完整 patch 迫使全链重复推导**：v3 不含完整 cells/metadata/extents；command inverse、Java reducer、remote handler 和 preview/live reducer 分别解释结构。这是 S1–S3 之外的公共架构缺口，单改局部循环并没有关闭它。迁移必须包含 protocol/history/server/persistence，不能只给 v3 增加一个名义 planner。
5. **S5 undo owner patch 重复复制整本快照（本轮代码已改、CI/本地 Java 验证待完成）**：`StructuralMutationDescriptor.applyWithPatch` 已返回独立候选快照，`WorkbookOperationService.commitInternal` 再以 detached `applyStructuralPatch` 应用 undo owner delta，会为该 undo 再深拷贝整本 JSON。现在只对这个独占候选走 owned reducer，保留原 `next` 供后续权限 owner 检查与数据块引用比较；一般结构 reducer 仍保持 detached 契约，未冒险原地改写变更前状态。

### C：图表语义没有由完整几何事实承载

5. **C1 簇状条形图系列重叠**：`layout.ts:477–491` 算了 `ordinal/band/offset`，但 bar 分支 y 未用 offset、height 用整条带。两个系列同一 category 会重叠；column 分支有 offset。应统一横纵方向的 series slot geometry，绘制和命中消费同一矩形。
6. **C2 自动对数轴无效**：`axisBounds:160` 对正值数据仍默认 `minimum = min(0,dataMinimum)`，随后 logarithmic 分支拒绝 minimum<=0。切换 log 且未手填最小值即失败。需以轴领域选择合适正值边界，不能在 Canvas 忽略错误。
7. **C3 子类型静默折叠**：柱/条 cone/cylinder/pyramid/3D 最终均 `fillRect`；pie-of-pie/bar-of-pie 仍走单饼扇区；surface 的 3D/wireframe/contour 均绘成热图矩形。模型允许这些类型，因此是缺少真实实现而非仅菜单少项。不能以添加标签或关闭报错宣称支持。
8. **C4 层级信息未进入图表几何**：treemap 和 sunburst 在 renderer 将所有 point flatMap，分别画一排矩形/一圈圆环；没有父子层级 owner。需从 source binding 保留层级并在 layout 构造分区事实，hit-test 不能重新展开数据。
9. **C5 雷达图负值被改成正值（本轮已修复）**：原 layout maximum 与 renderer/hit-test 均 `abs(value)`，-10 和 +10 落在相同半径。现在 layout 按 canonical value axis 一次生成 signed vertices，绘制与 hit-test 共用坐标；`radar`、`radar-markers`、`radar-filled` 分别控制线、标记和封闭多边形填充。chart layout 回归覆盖负/零/正/缺失值，Canvas 回归确认命中索引与顶点一致。
10. **C6 趋势线不是对应算法**：`buildTrendline:337` 对所有类型先做原始线性回归；多项式用 `predictor ** 2 * slope * 0.02`，指数/对数/幂也沿用线性系数；forecast 是加在 y 上而不是扩展 x 域。必须用正式回归领域输出曲线、统计及 forecast 范围，覆盖拒绝域和阶数。
11. **C7 直方图 underflow/overflow 丢数据**：`histogram:391` 直接 continue 不生成边界箱，图上总数不等于输入。过小 binWidth 还可产生与输入规模无关的超大 Array.from。需明确边界箱与容量拒绝，不能静默截断或随意改变用户设置。
12. **C8 数据表仍是占位文本**：`drawChartLayoutOnCanvas:1404` 只绘制 `Chart Data Table` 字样，不绘制实际系列/类别/值。数据表必须进入同一布局与命中模型。
13. **C9 大数据按参数展开及 sparkline 空值 O(n²)（本轮已修复）**：layout `axisValuesForSeries` 的 push(...values)，map/stock extrema、structured min/max 和 sparkline group bounds 的 Math.max/min(...array) 会受引擎参数个数上限影响并制造临时数组；same-group 还通过逐项 `find` 形成重复扫描，sparkline `emptyCells=connect` 对每个空槽切片反向查找前值，最坏为二次复杂度。本轮改为逐项统计、sparkline ID Map 和前值状态扫描；130,000 点布局/同组边界回归覆盖引擎参数阈值，空值连接回归校验线段保持同值。

### X：native codec 与图表领域脱节

14. **X1 现代图表写错原生 vocabulary**：`native-chart.ts:294–299` 写 `c:treemapChart` 等 2006 chart namespace 节点；微软 Office2016 使用 ChartEx `cx:chartSpace/chart` 与 layout。需要真正的 ChartEx package/relationships/codec，保留未知部件；自产自读通过不等于 Excel 兼容。
15. **X2 系列与轴语义丢失**：combo exporter 仅按 chartType 分组，将所有组绑定 201/202，忽略 series.axis；scatter/bubble 的 axisXml 仍默认分类轴；stock 把 roles 写为单个 ser 内 open/high/low/close/vol 元素。需 type-specific 原生系列/轴编译，不靠通用 XML 字符串拼接重猜模型。
16. **X3 原生类型识别丢子类型**：classifier 对 line3DChart 不保留 3D subtype，surfaceChart/3DChart 仅凭 wireframe 区分，ofPieChart family=pie 后 subtype 落成 pie。读写 codec 需共享正式 family/subtype 映射和拒绝边界。
17. **P1 服务端图表修改重复全量复制**：DrawingMutationDescriptor 的 affectedRanges 用 apply(snapshot) 预检，apply 再 deepCopy/reduce。重绘/拖动阶段之外的服务端计算和 heap 同样应计入性能。需事务拥有一个已验证的写入计划，而非跳过预检。

### B：基础操作追加证据

18. **B1 格式化隐式空白区域造成体积膨胀**：`sheet-features/src/index.ts:1818` 对选择矩形的每个地址生成 before 和 cell.restore；`style.set:1748` 为每个空白地址创建 `{value:null}` 再写样式。整列/整表格式操作的时间、history 和存储都与矩形面积相关，而不只是 occupied cells。不能简单跳过空白（会丢失 Excel 格式语义），应将范围/行列格式纳入 canonical style owner，统一 cell resolution、render、undo、server 与 OOXML。
19. **B2 Table 重命名 command/replay 不同语义**：`sheet-table-commands.ts:175` 的 command 执行 `planSheetTableRename(...).apply()`，但 `sheetTable.update` mutation handler（124）只替换 table model。CommandRuntime 把公式 delta 附到 inverse，undo 可还原引用；redo forward 无相应公式 facts，handler 也不重新改引用。最小链为 `Table1→Table2 → undo → redo`，表名和公式可能分离。解决点是完整事实 replay，不是在 redo handler 再补一套 rename。

20. **B3 Sheet 删除撤销快照遍历整个 Workbook**：`getSheetSnapshot` 从 `WorkbookModel.snapshot()` 生成所有工作表的完整快照后才选中目标 Sheet。成本与全簿所有单元格数和对象数相关；多 Sheet、大数据文件仅撤销删除一个 Sheet 就复制无关数据。本轮改为目标 `WorksheetModel.snapshot()`，并让恢复直接用 `WorksheetModel.fromSnapshot()`，保留延迟单元格 hydration。
21. **B4 Sheet 恢复依赖临时 Workbook 且可能部分提交**：旧恢复路径把当前 Workbook 其余工作表移除后用 `WorkbookModel.fromSnapshot()` 解析单 Sheet，跨 Sheet anchor 的名称因此无法通过所有权校验；之后逐个调用 `setDefinedName`，重复名字可能覆盖。打印文档也在工作表插入之后才写入。现先构造 Sheet、组合校验名称 identity/anchor、校验打印文档所有者并标准化，再一次性更新名称并插入；拒绝测试确认失败时 Workbook 快照不变。

22. **C10 簇状横条系列几何重叠**：`layout.ts` 的 bar 分支读取系列 offset 只计算 x 坐标（横向条形的数值轴），没有把系列 offset 用在 y 和 height，因此同一类别的系列矩形重合。现让 y 使用系列 slot offset、height 使用每个系列的 band；Canvas 绘制和 hit-test 共用这组矩形。
23. **C11 正值对数轴默认最小值为零**：`axisBounds` 对非百分比轴统一从 0 起算，之后 logarithmic 检查又拒绝 `minimum <= 0`。即使数据全为正，未手填最小值的对数轴也会被判无效。本轮对数轴默认采用最小正数据值，并只在最大值为自动值时扩展；显式边界仍按有效域校验。
24. **B5 worksheet identity 覆盖与标签歧义**：`WorkbookModel.fromSnapshot` 对重复 `sheet.id` 连续执行 `Map.set`，后一个 Sheet 覆盖前一个，而 `sheetOrder` 保留重复 ID；`getSheetByName` 按不区分大小写查找，重复标签会令公式/命令落到第一个匹配项。现在在 canonical snapshot、model hydration、创建、恢复、重命名和复制之前校验唯一 ID 与 case-insensitive name；冲突拒绝路径不改变模型。
25. **B6 其他 workbook Map owner 重复时静默覆盖**：`fromSnapshot` 对 table、relationship、view 直接 `Map.set`，query definition 和 print document setter 也按 id/sheetId 覆盖已有 owner；style template 的直接 hydration 同样可能覆盖。现在 canonical 与直接 hydration 共用 Map owner identity 预检，涵盖 data source、table、relationship、view、query、style template 和每 Sheet 唯一打印文档；来源 Map 原有 `addDataSource` 拒绝语义保持一致。
26. **C12 瀑布图连接线累计/端点错误且零值条不可命中（本轮已修复）**：布局把 from/to 排序成 min/max 后丢失累计方向；负增量连接线取错边，总计柱入线误落到零基线，连接线也没有从前柱边缘接到当前柱。绘制还强制 1px 最小高度，hit-test 却只用原始零高边界。现在 layout 为 connector 保留进入当前点前的累计值并按相邻 bar geometry 计算两端；draw/hit 共用规范几何。
27. **C13 箱线图两个声明选项完全无效（本轮已修复）**：`showInnerPoints`、`showMeanMarkers` 只有 domain 字段，仓库无任何消费者；layout 现在根据须线内数据与 mean marker 开关生成事实，Canvas 按事实绘制。拒绝/无效选项校验仍需补齐。
28. **C14 直方图 `by-category` 被当作自动数值分箱（本轮已修复）**：旧实现只提取数值并用 Scott width，重复文本类别没有分组。现在按有类型身份的类别聚合，并把数值列求和；layout 同时提供分类 bin 几何，Canvas 命中返回 category。Pareto 对负聚合值 fail-close。
29. **C15 饼图/圆环图数据标签配置不进入渲染（本轮已修复）**：`drawChartDataLabels` 只在 cartesian 分支调用，pie slices 从未显示类别/值/百分比；且全部字段显式关闭时旧逻辑仍强行生成值标签。现在 slice layout 按 series/chart 选项生成标签文本、放置坐标和 hit bounds；无启用字段时不造标签，renderer 与 hit-test 使用同一 slice label facts。手动格式/全部 label positioning 的 Excel 细节仍待实测。

另记 **T1 测试契约漂移**：缺失 Pivot 图表代码已返回 `chart` drawable 并带 typed error/hit area，旧测试仍断言 `shape`，导致套件本身失败。现断言实际错误文本和 chart-area 命中。当前根因登记共 30 条（29 个运行时根因 + 1 条测试契约问题），但不是 30 条全已整改，也不是 60 个操作都已审完。

能力缺口单列，不冒充 silent corruption：跨 Sheet cut/paste 在 `editing/index.ts:1461` 明确 UNSUPPORTED；带外部引用的 Sheet 删除在 `sheet-identity-transform.ts:856` 明确拒绝。要达到目标 Excel 语义，仍须在 Java authority 中建立可逆跨表 move/delete-reference facts；不能只删除这些 guard。

以上按根因分组，不把每个子类型或引用位置再拆成数量。其余未深入链路保持待审，不默认无缺陷。

### 已排除的误判

- 负值柱形/条形的 start/end 已先取 min/max 再 scale，因此“负柱被 Math.max 压成 1px”不是当前源码问题；真正证据是 bar 多系列 slot 忽略 offset。
- 20 家族已经在 domain、命令注册和编辑面板存在，不能把“没有看到菜单”直接记为缺失类型。
- 服务端已有 exact-base 与权限检查，不能用“重复 reducer”推出“没有事务隔离/版本控制”。
- 未知原生部件的 unchanged-save/fail-close 已存在；X1–X3 针对被声明为可编辑 owned 的格式，不是要求删除未知内容。
- Sheet 删除并非“已改写其他表却未保存撤销”：当前 planner 对外部引用先拒绝。应登记能力缺口，不误报为成功删除后的引用数据丢失。
- 合并非 anchor 内容由 `home-commands.ts:1107` 的 clear + anchor restore 处理，不能只看到 `merge.set` 加 span 就断言 UI 合并保留了所有非 anchor 值。

## 本轮六轮真实问题复核

1. **调用输入与 owner**：Sheet 恢复名称必须归属于被恢复的 Sheet；全簿名称投影仍只含 workbook owner。拒绝后快照比较确认没有部分写入。
2. **事务边界**：无效插入位置、错误名称 owner、重复名称 identity、错误打印文档 owner 都在写入前被拒绝；Sheet map、顺序、名称 map 和打印文档没有分步泄漏。
3. **内存与稀疏数据**：恢复只 hydration 目标 Sheet；稀疏单元格仍 deferred。撤销快照只深拷贝目标 Sheet，普通 `CellMatrix.toJSON` 和全簿快照维持原有 cell clone 成本。
4. **图表几何与命中**：簇状横条系列使用分开的 y slot 与 band 高度；Canvas 绘制和 `chartHitTest` 都消费 `ChartLayoutBar`，几何一致。
5. **坐标轴边界**：全正值 log 数据默认从最小正数开始；最大值仅在未显式指定时自动扩展。非正数据与用户设置的冲突边界仍被判为无效。
6. **服务端权威**：Java 事实 DTO 的冲突预检和精确回放有测试，但在线 `WorkbookOperationService` 仍调用结构 reducer，前端仍按 intent 执行本地/远端结构变换；Java 唯一规划权尚未实现，本项保持未通过。

本轮定向证据：core-model 59/59、Java structural facts 10/10、chart layout 9/9。项目级 TypeScript 检查未通过，因为已安装依赖缺少根 package 声明的 `@types/react`，产生 3243 条连锁类型诊断；修改文件不在诊断路径内，但因此不能记作 typecheck 通过。

## 本轮雷达图六轮真实问题复核

1. **数据/类别绑定**：series 值与类别按同一索引输入顶点数组；layout 固定 `count`，空值保留槽位且标为不可见，不压缩后续类别索引。
2. **有符号轴语义**：正负数据进入同一 `valueAxis` 的 min/max 与 `scale`；不再对值取绝对值。回归样本 `[-10, 0, 10]` 的半径严格递增，缺失值不伪装为零。
3. **绘制几何**：每个顶点的坐标在 layout 中生成一次；Canvas rings、series edges、filled polygon 和 markers 都引用同一 center/radius/vertices，不在 renderer 再推导数值半径。
4. **交互几何**：`chartHitTest` 直接命中对应 vertex，并返回相同类别的 point index；Canvas 集成回归分别命中 -10 与 +10，确认两位置不重叠。
5. **子类型与缺口**：只有 `radar-filled` 填充完整闭合系列；`radar-markers` 画点标记；缺失值两侧断线且不会形成闭合填充，避免用不存在的数据跨越缺口。
6. **拒绝与验收边界**：无数值系列仍返回 invalid；缺失 Pivot 的既有 renderer 测试曾错误要求 `shape`，现改为验证 `chart` 错误视图和 chart-area 命中。chart 10/10、Canvas 8/8；项目 `tsc` 退出码 2，原因为缺少 React typings，改动文件无诊断。浏览器、性能及真实 Excel 仍未验收。

该复核只关闭 C5 的一个根因组，不代表用户要求的 30 个新增问题或 60 个操作已审完；未完成项继续留在本 PR 的同一审计范围内。

## 本轮图表基础操作六轮真实问题复核

1. **瀑布累计状态**：入线保存进入当前点前的累计值；负增量不再因矩形上下边排序而丢失方向；total 柱从零绘制，但入线仍连接前一累计值，后续点再以 total 值继续累计。
2. **瀑布绘制/命中等价**：layout 只计算 bar rectangle 和连接线；连接线两端与相邻柱边一致；renderer 与 hit-test 共用 rectangle，零值条的 1px 可见范围也可点击。
3. **箱线统计**：inner points 使用 lower/upper whisker 范围内的原始观测；离群点仍独立受 `showOutlierPoints` 控制；mean marker 不混进 whisker 或异常值统计。
4. **分类直方图**：按 PivotScalar 类型区分 `1`、`"1"`、空白与错误类别；重复类别按 Excel 支持说明求和，Pareto 的负总额明确拒绝，所有 bin 的绘制与命中几何由 layout 生成。
5. **饼/圆环标签**：series label overrides chart-level labels；百分比在对应 pie/doughnut ring 内计算；爆炸切片偏移同步应用于 label 位置和 hit bounds；所有字段显式关闭时不再强行展示数值。
6. **拒绝与门禁**：chart 15/15、Canvas 13/13；全项目 `tsc` 仍因缺少 React typings 退出 2，修改 TS/TSX 文件目标诊断为 0。Vite 浏览器页已打开但 `GET /api/auth/config` 返回 500，未进入 workbook；不绕过认证，UI/OOXML/native Excel/大数据性能仍阻塞或未验收。

## 本轮图表大数据路径六轮复核

1. **参数边界**：chart axis series 不再 `push(...values)`；130,000 点 line layout 实测构建成功，点数保持完整，无静默截断。
2. **地图分布**：map extrema 在遍历已映射值时累计 min/max，不再额外创建 finite 数组或传入巨型参数列表；无有效值仍保持既有 0/1 空域。
3. **股价范围**：draw 与 hit-test 均移除 `flatMap + Math.min/max spread`，仅用常量数量的标量更新界限；volume 最大值同样单遍扫描。
4. **结构化聚合**：table/report series 的 min/max 通过 reduce 扫描已有数组，不复制元素到参数栈；空输入仍在既有分支中先返回。
5. **Sparkline 内存与复杂度**：same-group 从每个 id 搜索全体 sparkline 改为一次 Map 建索引，跨 series 的 group bounds 逐项累计、不再拼接总 values；`connect` 空槽沿用上一有效值状态，避免 per-gap slice/reverse/find。
6. **实证边界**：chart 15/15、Canvas 13/13；全项目 typecheck 退出码 2 的原因仍为 React typings 缺失，5 个改动 TS/TSX 文件无诊断。in-app browser 的 API auth config 500 阻止 UI 和网络保存链实测；百万级性能、OOXML/Excel 仍未验收。

## 本轮第二组六轮身份边界复核

1. **Map identity**：重复 Sheet ID、table/relationship/view id、query/style id 和 print sheet owner 原先会覆盖旧 owner；所有 Map 写入前统一拒绝。
2. **名称解析**：Sheet 查找使用 `toLowerCase()`；重复名称验证使用相同折叠规则，不改变既有 lookup 语义。
3. **加载入口**：canonical validator 与直接 `WorkbookModel.fromSnapshot` 调用同一 identity contracts，防止内部 snapshot/hydration 旁路。
4. **编辑入口**：Sheet create/restore/rename/duplicate 都预检 ID/name；冲突发生在引用重写或 Sheet 插入之前。
5. **回归/原子性**：重复 Sheet/map owner identity 及大小写重名均由 59 项 core-model suite 覆盖；拒绝分支保留现有工作簿状态。
6. **协议/迁移**：未改 snapshot 版本或序列化字段；Java 在线规划权仍未实现，identity guard 不能替代用户选择的服务端结构 planner。

## 有界统一实施方案

1. **先完成 60 项静态证据闭环**，将上述局部/待审补齐。基础编辑、chart source/reference/geometry/native codec、history/server transaction 各有唯一 owner；不要在 renderer、UI 或 export 修上游语义。
2. **结构事务**：Java 是唯一结构及引用规划 authority；以完整、可逆、版本化的 facts 替代 inverse intent 与 partial v3 patch；before 前置身份/版本及 after extents/cells/metadata/reference/impact 一起规划、校验、提交。前端发 intent + revision，服务器准备/权限/提交，前端应用已确认 facts 后才发布计算与显示影响。History/remote/server replay 只执行 facts。同步迁移消费者、删 TS 结构推导与旧反方向 history reducer；local-only 的结构操作必须显式要求连接服务，不引入新运行栈或隐式兼容执行。
3. **图表领域**：沿既有 ChartDrawingPayload → resolved data → ChartLayout 链收敛；扩展明确的数据角色、层级、统计及 type/subtype capability，不新增第二份图表状态。布局一次产生用于 draw、hit、labels、data table 的全部几何；删除 renderer/hit-test 的重复统计算法和降级形状。能力声明必须覆盖创建/编辑/渲染/导入/导出，缺少实现保留真实错误及待办，不算整改完成。
4. **原生编译**：经典 chart 与 ChartEx 各按真实 schema 编译，共享 canonical 角色/轴身份；分类与 subtype 读写对称。format conversion 只能在显式边界做，不修改 unknown/macro preservation 权属。
5. **复杂度/事务**：按实际 affected owners/cells 处理；统计与几何不展开巨型参数；服务端校验不以两次全 workbook apply 为代价。新增事实不能退化为复制整个 workbook 的补丁。
6. **集中完成实现后六轮复核**：产品语义；输入/权限/拒绝；owner/reference 完整性；事务/history/remote；复杂度/lazy/内存；OOXML/真实验收证据。每轮记录新增证据与被推翻假设，不能把同一段 diff 看六次算六轮。

当前 localOnly 默认规则位于 `runtime.ts:228`，浏览器直连/本地模式确实存在；用户已授权调整结构操作在线前置条件。**决策已定，实现未完成**，不能把当前 partial v3 patch 称为 Java-only 架构。切换必须包括所有绕过 dispatch 直接调用 runCommand 的结构入口、未确认操作的 UI 状态、撤销记录、事件重放与检查点恢复。

## 本轮服务端 undo owner-patch 六轮静态复核

1. **独占所有权**：公共 `applyWithPatch` 契约不修改输入并返回独立快照；结构 descriptor 确实先 `deepCopy()`。只有该返回候选适合 owned patch，不把工作簿当前基线交给可变 reducer。
2. **变更前状态消费者**：commit 仍需用原 `next` 执行 `committedRanges` 的保护 owner 检查，并与候选比较 data-block refs。直接原地改写 `next` 会破坏两项基准，所以本轮没有采用全结构 reducer in-place。
3. **提交边界**：operation row、revision event、workbook revision/checkpoint 的写入均发生在 mutation loop 后。inverse patch 预检失败时，受影响的只是未发布候选；基础快照与持久化行仍在原状态。
4. **拒绝语义**：owner delta 按序应用，后续 precondition 冲突可能发生在前一 delta 已写入候选之后。owned API 明确要求异常时丢弃候选；新增拒绝用例覆盖“候选部分变化、基础快照不变”。
5. **复杂度边界**：undo 的每个非空 inverse patch 不再额外复制完整 JSON 树，patch 仍按 owner delta 处理。本轮不声称已消除正常结构 reducer 自身的快照 clone，也没有大数据基准证据。
6. **契约与验收**：detached `applyStructuralPatch` 保留给需要隔离的调用方，仅 commit undo 路径转用 owned 版本；新增成功路径核对返回对象身份，拒绝路径核对 precondition 和基础快照。测试未执行：本机只有 JDK 17、无 Maven，CI Java 21 结果待本次 push 后读取。

这是一项已证实并已修改的独立性能问题，不等同于每轮至少 30 项根因完成整改；当前结构规划权、history facts、客户端重复推导等主目标仍未解决，本审计和 PR 保持进行中。

## 最终实测门禁（当前未执行）

- 行为回归：每个根因至少成功与拒绝路径；矩阵中每个操作有实际命令/输入与观测结果。测完整链和状态等价，不只测 isolated helper。
- 前端：受影响 unit suites、`npm run typecheck`、`npm run check:boundaries`、`npm run build`；真实脚本见 frontend-react/package.json。保留退出码，长输出写仓库外日志。
- 后端：按 backend/pom.xml 的实际 Java21/H2 构建/测试入口执行；commit、undo、remote replay、reload 同一 corpus，记录 PR checks。
- 浏览器：真实应用操作、Console/Network、权限拒绝、保存后重开、拖动/命中、图表每家族及子类型、freeze/跨 Sheet；使用真实 server/data，不以 mock handler 或只调 session 方法代替 UI。
- 数据规模：至少 100k formulas、500k/1m **occupied cells** 和多 Sheet/chart；分别采集 latency、heap、scanned/transformed owners、worker bytes、payload、dirty ranges。稀疏末行坐标不算百万数据量。
- OOXML：真实 Excel 文件导入→编辑→导出→重导入；native package schema validation；桌面 Excel 打开/保存校验。缺失 Excel/环境时只将该验收项标为 Blocked，不能整批宣称完成。
- 集中改动及证据交付同一个 draft PR #345。未完成上述 gates 不合并、不标完成；回退为整个事实契约/消费者/迁移的配套回退，不恢复一侧旧语义。

## 微软官方基准

- [Available chart types in Office](https://support.microsoft.com/en-us/excel/available-chart-types-in-office)：产品家族与子类型依据。
- [Create a histogram](https://support.microsoft.com/en-us/excel/create-a-histogram)：by-category 聚合、underflow/overflow 与 bins 边界依据。
- [Create a box and whisker chart](https://support.microsoft.com/en-us/excel/create-a-box-and-whisker-chart)：inner points、outliers、mean markers 与 quartile 设置依据。
- [Trendline options in Office](https://support.microsoft.com/en-us/office/trendline-options-in-office)：趋势线拟合类型、阶数与 moving average 点数依据。
- [XlChartType](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype)：原生类型身份依据。
- [Office2016 ChartDrawing](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2016.drawing.chartdrawing?view=openxml-3.0.1)：ChartEx vocabulary 依据；具体 codec 实现仍需逐项 schema 核对与真实文件验收。

性能设计和代码根因是本仓库的工程分析，不声称微软内部采用了本仓库的实现。
