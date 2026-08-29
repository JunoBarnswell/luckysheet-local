# Univer 设计吸收与工作簿引擎收敛 PRD

状态：Frozen v1  
日期：2026-08-29  
目标分支：`codex/univer-engine-convergence`  
本项目基线：`d3be802bbd0f0c5ff4485f32f76ed663fa16cbfc`  
Univer 源码基线：`dev@524a9aa357ac8b0266500d3a8cbfb6fb842a9735`（`1.0.0-beta.2`）  
Univer Pro 可读 npm 样本：`0.25.1`（仅作为架构与类型契约证据，不作为运行时兼容版本）

## 1. 产品目标

本计划不是把 Univer 作为第二套 spreadsheet runtime 嵌入项目，也不是复制一组 Ribbon 按钮。目标是把 Univer 已验证的优秀机制吸收到本项目唯一语义链中，并修复当前 `main` 中仍然存在的真实断链：

```text
Home / Insert / shortcut / context menu
  -> typed intent
  -> canonical command
  -> synchronous mutation / transient operation
  -> model + domain revision
  -> formula / rule / analytics derived state
  -> PaneMap + render projection
  -> history / collaboration / persistence
  -> native document import/export
  -> backend authoritative reducer
```

完成后的产品必须同时提升：

1. 综合功能：用户列出的工作簿、工作表、区域、选择、公式、数字格式、筛选、排序、数据验证、条件格式、超链接、评论、查找替换、批注、表格、绘图、导入导出、打印、图表、数据透视表、迷你图、分级显示、形状、单元格内图形、数据连接器和区域预处理全部具有真实语义。
2. 性能：局部动作的成本由受影响范围、依赖闭包或对象集合决定，不再与整个工作簿、整个历史或整个数据源线性绑定。
3. 一致性：Canvas、编辑器、命令、协作、后端、OOXML 和打印使用相同地址、值、格式、隐藏行与对象所有权。
4. 可观察失败：不支持、版本不匹配、资源损坏、Worker 失步和后端历史缺口必须 fail-close。

许可证不在本 PRD 的问题范围内。

## 2. 非目标

- 不把 `@univerjs/*` 或 `@univerjs-pro/*` 直接安装到生产运行时。
- 不增加 Univer/React Sheets 双模型、双写、别名、兼容读取器或 UI-only repair。
- 不把 `0.25.1` Pro API 当作本项目协议版本。
- 不用 notice、空 mutation、空表、空图或 cached value 冒充成功。
- 不在页面组件中重新解释命令、权限、选择、值或资源状态。
- 不以单元测试、构建通过或既有 Issue 已关闭代替真实行为和互操作验收。

## 3. 用户场景

### 3.1 Home

- 用户在任意合法选择上执行剪贴板、字体、对齐、数字格式、样式、行列/单元格、清除、填充、筛选排序、查找替换。
- mixed selection 的按钮状态来自同一 selection aggregate，不允许页面局部猜测。
- 任何命令必须支持权限、保护、合并、spill、隐藏投影、撤销、协作和持久化。

### 3.2 Insert

- 用户插入 Table、Pivot、Chart、Sparkline、Shape、Connector、图片、单元格内图形、Hyperlink、Comment、Note、Text Box 等对象。
- 每个可见入口必须产生真实 command/dialog；宿主能力不可用时显示明确 disabled reason，并返回 `UNSUPPORTED_FEATURE`。
- 对象插入后可以选择、编辑、复制、撤销、协作、保存、重新加载和 OOXML round-trip。

### 3.3 大工作簿与数据源

- 100k×20 dense workbook 和 1M logical-row sparse workbook 中，滚动、选择、局部格式、单格编辑和依赖重算保持可交互。
- JDBC/REST/CSV/TSV/XLSX 数据通过分页/列式块进入数据源域；查询、筛选、Pivot、图表和公式不需要把整表物化到主线程或 JVM heap。

### 3.4 互操作与打印

- Excel 文件导入后，用户可识别的对象进入 canonical model；未知对象保留原始 part ownership。
- 编辑后导出不得静默删除未知 worksheet node、评论、验证、条件格式、DrawingML 或连接。
- PDF/打印预览必须绘制真实单元格、样式、隐藏投影、重复标题、图表和绘图对象。

## 4. 审计基线与已确认根因

### 4.1 运行时和功能装配

- `frontend-react/packages/spreadsheet-app/src/feature-registry.ts:44-80` 启动时一次性注册所有功能；没有 Starting/Ready/Rendered/Steady 生命周期，也没有按环境或文档类型拆分。
- `frontend-react/packages/spreadsheet-app/src/workbook-session.ts` 约 6500 行，既负责 UI intent，也负责 Pivot、Drawing、Print、Query、Native Document 和 Formula orchestration。
- `frontend-react/packages/spreadsheet-app/src/ui-command-catalog/index.ts` 约 2080 行，菜单可见性与真实 capability 没有编译期闭包验证。
- `frontend-react/apps/web/src/components/InsertRibbon.tsx:70-80` 存在可见但无行为的 `More Charts...`。
- `frontend-react/packages/spreadsheet-app/src/ui-command-catalog/index.ts:1789-1800` 的 Forms/Screenshot 仅 notice，不是功能事务。

Univer 的可吸收机制：

- `packages/ui/src/services/menu/types.ts:29-73` 与 `menu-manager.service.ts:95-181` 把 Ribbon 作为插件菜单 schema。
- `packages/sheets-formula-ui/src/plugin.ts:91-138`、`sheets-data-validation-ui/src/plugin.ts:84-135`、`sheets-drawing-ui/src/plugin.ts:95-150` 按生命周期装配 model/controller/render/UI，并把非首屏控制器推迟到 Steady。

### 4.2 公式与数字格式

- `frontend-react/packages/spreadsheet-app/src/runtime.ts:375-405` 每次调度取消/reset 后重新扫描工作簿并注册公式和值。
- `frontend-react/packages/formula-engine/src/formula-engine.ts:421-473` 每个 Worker task 克隆、排序并发送全部 cells、names、tables 和 spill occupancy。
- `frontend-react/packages/formula-engine/src/formula-engine.ts:869-918` 重算重新构建完整依赖图；`range-index.ts:48-115` 更新时重建范围树。
- `dependencies.ts:75-77` 只记录静态 AST 引用；`evaluator.ts:447-495` 的 INDIRECT/OFFSET 动态目标不进入依赖图。
- `evaluator.ts:322-355` 在 IF/IFERROR 前 eager evaluation 全部参数。
- `frontend-react/packages/number-format/src/index.ts:37-39,536-539,572-599` 固定 1900 epoch、General `toFixed(6)`，且没有完整 section/condition/color/placeholder 语义。

Univer Pro 可吸收机制：数字化 CellId、Point/Range index、dirty arrays、shared formula compression、持久 dependency engine 和计算计划；这些是算法设计证据，不是依赖引入计划。

### 4.3 规则、筛选、排序和区域

- `DataValidationPanel.tsx:5-10,24-41` 不接收真实选择，固定创建 `B2:B21`，不同 validation type 仍写列表文本。
- `sheet-features/src/data-features.ts:1225-1230` 重叠 validation 取第一条；`:1275-1308` custom validation 绕开 FormulaEngine。
- `data-features.ts:406-495,599-633` 条件格式读取存储值并自建公式求值；错误会退化成“不匹配”。
- `core-model/src/data-transform.ts:204-214` 稀疏排序为每一行遍历整张 cell store，再对矩形逐格删除。
- Table AutoFilter、Worksheet AutoFilter、hidden row、formula visibility 和 export 必须继续保持一个 typed resolver；不得复制 Univer 已暴露出的双所有权缺陷。

### 4.4 Review、链接和查找

- `FeatureSidebar.tsx:730,781-800` 编辑现有评论仍调用 add；`workbook-session.ts:5155-5174` 总是生成新 thread id。
- Hyperlink 只进入 snapshot display，未进入 Canvas hit/activation；双击链接单元格会进入编辑。
- `sheet-features/src/find-replace.ts:164-181,240-299` 每次查询全量扫描、排序并反复编译正则，且只读取 materialized cells。
- `exchange-excel-ooxml/src/ooxml.ts:2385-2401` 只读部分 legacy comments，保存不会根据 canonical review 重写 comments/threadedComments/persons。

### 4.5 Rendering、Drawing 和单元格内图形

- `SheetCanvas.tsx:517-568` provider 在绘制热路径重复解析 merge、validation、review、hyperlink 和 style。
- `cell-renderer.ts:118-243` grid/content 阶段重复遍历；`drawing-renderers.ts:1517-1542` 浮动图表重读 source 并重算 layout。
- 本项目 Shape gallery 约 47 种；Univer Pro Shape 类型契约覆盖完整 preset shape/connector/route model。
- `exchange-excel-ooxml/src/ooxml.ts:700-714` 导入 sheet 时返回空 drawings；`:498-566` 只输出图片 anchor，没有完整 Shape/TextBox/Connector writer。
- 后端 `DrawingMutationDescriptor.java:207-214` 的允许类型与前端 connector mutation 不一致。
- In-cell image 资源未 ready/error 时当前 render 会直接不画；打印和导出没有统一 object renderer。

### 4.6 Chart、Pivot、Sparkline、Outline

- block-backed data source 创建 Pivot 时会构造 `data-source` Pivot，随后调用明确拒绝该 source 的同步 input builder。
- Chart/Pivot/Sparkline 已有 typed domain，但 OOXML capability 仍为 partial；source revision、render cache、worker result 和 target collision 需要统一事务。
- Univer Pro 的优势是 engine/model、sheet integration、UI、RPC/worker 分层；本项目不得复制其全量 records/tuple maps 或全局可变配置。

### 4.7 Native document

- 后端 `WorkbookImportController.java:25-37` 和 `WorkbookCatalogService.java:295-339` 信任浏览器上传的 snapshot/format/nativeMetadata，没有证明它们与原文件一致。
- `spreadsheet-app/src/features/native-document/index.ts:12-35` 的 `document.import/export` 是零 mutation 空处理器，真实交换走另一条直接调用链。
- capability manifest 已明确 formulas、CF、validation、filters、print、pivot、slicer、timeline、XLSB/BIFF 等为 partial/none，但 balanced/best-effort 仍可能返回 buffer。
- OOXML 主链声明了 `maxCells/maxXmlDepth/maxXmlBytes`，实际只执行 ZIP 预算。
- unknown worksheet nodes 只检测，不具有通用 ownership-preserving writer。

### 4.8 Print

- `spreadsheet-app/src/features/print/pdf-export.ts:75-145` 默认只写标题、页码、sheet id 和行列范围。
- `workbook-session.ts:6106-6128` 未传真实 page renderer/pageText，仍通知 PDF exported。
- Unicode literal、fit-to-page 缩放、重复标题、Drawing 与资源失败没有完整语义。

### 4.9 Data connector 与区域预处理

- QueryPanel 对除 JSON 外的 connector 都发送同一种 `{url: jsonData}` 配置，与 CSV/TSV/XLSX/REST/JDBC contract 不一致。
- `DataSourceContentQuery.prefetchRows/getRows` 按每一逻辑行查找 block，没有 block interval index。
- 后端 DataSource mutation 只校验 manifest 字段，不验证 block bytes/checksum/length 实际存在。
- `QueryExecutionService.java:166-260` 把 JDBC/REST 完整物化为 `List<List<JsonNode>>` 再执行 filter/join/group/pivot；取消只取消 Future，SQL read-only 主要依赖字符串检查。

### 4.10 Backend operation、协作和持久化

- `WorkbookOperationService.java:113-134,280-310` 在每次提交/读取时从 checkpoint 重放 operation。
- `WorkbookOperationService.java:317-321` 为一次 checkpoint 判断两次调用 `store.listOperations`，而 `WorkbookStore.java:208-215` 会加载该 workbook 全部历史。
- 客户端历史请求异常会继续 replay pending operations；历史缺口时无法证明结构坐标已 rebase。
- 服务不可用时会转入页面 memory workspace；该边界必须变成用户明确选择和可恢复状态，不能把 remote workbook 的失败当作成功 ready。

## 5. 目标架构

### 5.1 Feature Runtime

```text
FeatureManifest
  id / version / documentType / environment
  dependencies
  commands / mutations / operations
  model services
  render extensions
  ribbon surfaces
  permissions
  persistence resources
  native ownership

Lifecycle
  Starting  -> 注册类型、命令、服务，不访问 DOM
  Ready     -> workbook/sheet 实例存在后绑定模型
  Rendered  -> 注册 PaneMap/render extension/interaction
  Steady    -> 非首屏 panel、find index、gallery、context menu
```

强制规则：

- Runtime 只装载目标环境和文档类型需要的 feature。
- Home/Insert 由 manifest 编译成唯一 Ribbon schema。
- 可见 surface 的 command、permission、model owner、backend availability 和 failure code 必须闭包完整。
- feature load 失败使对应 surface unavailable；不得保留可点击旧状态。

### 5.2 Canonical Calculation Runtime

```text
mutation delta
  -> formula source/input revision
  -> persistent worker session
  -> Cell/Point/Range/Name/Table indexes
  -> dirty dependency closure
  -> lazy argument evaluation
  -> batched result/spill patch
  -> generation check
  -> projection invalidation
```

- 初始 snapshot 只注册一次；后续任务传 delta/dirty roots。
- dynamic reference 在运行时解析后注册依赖。
- Number Format 独立为 AST + locale/date-system context，由 cell、formula、filter、chart、print、OOXML 共同消费。

### 5.3 Canonical Resolved Cell 与 Rule Runtime

```text
ResolvedCell
  authored value/formula/style
  calculated value/spill
  data-block value + sparse overlay
  visibility projection
  number-format display

RuleRuntime
  AutoFilter/Table Filter owner resolver
  DataValidation resolver
  ConditionalFormatting resolver + range statistics cache
  Table/DataRegion context
```

- 所有 rule 使用同一 resolved-cell 和 FormulaEngine。
- overlapping owner 无法唯一解析时 fail-close。
- selection、sort、find、chart、pivot、print、export 不得各自实现值读取。

### 5.4 Render Runtime

```text
PaneMap visible ranges
  -> sheet/domain revisions
  -> CellRenderData cache
  -> one cell pass
  -> layer dirty rectangles
  -> drawing spatial index
  -> chart/object layout cache
```

- PaneMap 是坐标唯一权威。
- hit、selection、editing、commit、hyperlink 和 drawing handle 地址相同。
- 不支持或加载失败对象绘制明确错误 surface；不得消失。

### 5.5 Native Document Transaction

```text
source bytes
  -> format/resource budget
  -> native package graph
  -> typed feature ownership
  -> canonical snapshot
  -> trusted fileHash + snapshotHash binding
  -> edits with part ownership ledger
  -> patch writer
  -> compatibility proof
```

- import/export command 是唯一入口。
- unknown part/node/extensions 保留；无法安全 patch 时阻止保存。
- backend 不信任浏览器声明的 snapshot/file 绑定。

### 5.6 Backend State 与 Query Runtime

```text
Operation commit
  -> exact workbook lock/revision
  -> bounded checkpoint tail query
  -> typed reducer
  -> operation + outbox + state revision transaction
  -> async checkpoint compaction

Connector query
  -> configured source
  -> read-only/cancellable execution
  -> columnar block stream
  -> pushdown / hash aggregate / join
  -> checksum-bound manifest commit
```

## 6. 分类 Issue 计划

每个类别独立 Issue，禁止再次创建一个混合“95% parity”总 Issue。

| 编号 | 优先级 | 类别 | 核心交付 | 依赖 |
|---|---|---|---|---|
| [#320](https://github.com/JunoBarnswell/luckysheet-local/issues/320) UCP-01 | P0 | Feature Runtime / Home & Insert | 生命周期、依赖闭包、唯一 Ribbon schema、删除 notice/no-op surfaces | 无 |
| [#321](https://github.com/JunoBarnswell/luckysheet-local/issues/321) UCP-02 | P0 | Workbook/Sheet/Range/Selection/Render | 稀疏索引、PaneMap 地址一致、CellRenderData/对象增量缓存 | UCP-01 |
| [#322](https://github.com/JunoBarnswell/luckysheet-local/issues/322) UCP-03 | P0 | Formula Engine | 持久 Worker、delta/dirty dependency、dynamic ref、lazy args | UCP-01 |
| [#323](https://github.com/JunoBarnswell/luckysheet-local/issues/323) UCP-04 | P1 | Number Format | AST、1900/1904、General/section/locale、全域 display contract | UCP-03 |
| [#324](https://github.com/JunoBarnswell/luckysheet-local/issues/324) UCP-05 | P0 | Filter/Sort/Validation/CF/Table/DataRegion | 单一 resolved-cell、owner resolver、规则索引与拒绝路径 | UCP-02,UCP-03 |
| [#325](https://github.com/JunoBarnswell/luckysheet-local/issues/325) UCP-06 | P1 | Hyperlink/Comment/Note/Find Replace | thread update、链接激活、索引查找、OOXML review | UCP-02,UCP-08 |
| [#326](https://github.com/JunoBarnswell/luckysheet-local/issues/326) UCP-07 | P0 | Drawing/Shape/In-cell Graphics | DrawingML、shape/connector、资源、hit/print/export 闭环 | UCP-02,UCP-08 |
| [#327](https://github.com/JunoBarnswell/luckysheet-local/issues/327) UCP-08 | P0 | Native Import/Export | 唯一事务、可信绑定、资源预算、未知节点 preservation | UCP-01 |
| [#328](https://github.com/JunoBarnswell/luckysheet-local/issues/328) UCP-09 | P0 | Print/PDF | 真实页面渲染、Unicode、fit/repeat/object/resource | UCP-02,UCP-07 |
| [#329](https://github.com/JunoBarnswell/luckysheet-local/issues/329) UCP-10 | P0 | Chart/Pivot/Sparkline/Outline | source revision、Worker、layout/cache、OOXML | UCP-02,UCP-03,UCP-05,UCP-07 |
| [#330](https://github.com/JunoBarnswell/luckysheet-local/issues/330) UCP-11 | P0 | Data Connector/Range Preprocess | 动态配置、block interval、列式流、checksum commit | UCP-02,UCP-05 |
| [#331](https://github.com/JunoBarnswell/luckysheet-local/issues/331) UCP-12 | P0 | Backend Import/Collaboration Integrity | import authority、history gap fail-close、remote/offline 状态 | UCP-08 |
| [#332](https://github.com/JunoBarnswell/luckysheet-local/issues/332) UCP-13 | P0 | Backend Operation Tail Scalability | 删除全历史扫描，bounded replay/checkpoint tail | 无；与全链 clean-break 同批实施 |
| [#333](https://github.com/JunoBarnswell/luckysheet-local/issues/333) UCP-14 | P1 | Performance/Observability Corpus | 浏览器/JVM/Worker 指标、真实文件 corpus、回归门禁 | 横切 |

## 7. 性能合同

### 7.1 Frontend

| 场景 | 目标 |
|---|---|
| 100k×20 dense sheet 连续滚动 | frame p95 ≤ 16.7 ms；单个 long task < 50 ms |
| 1M logical-row、50k occupied sparse sheet 单格选择/编辑 | 不遍历 1M 行；同步阶段 p95 ≤ 20 ms |
| 单格公式编辑 | Worker payload 与 dirty closure 成正比；主线程 prepare p95 ≤ 16.7 ms |
| 100k 公式局部依赖链 | 不重新传输全部普通值单元格；可取消且 generation-safe |
| Home mixed selection | 已索引后 p95 ≤ 8 ms；不随 workbook 总 cells 线性增长 |
| Find next | 已构建索引 p95 ≤ 10 ms；索引按 content revision 增量维护 |
| Drawing move/resize | pointer frame p95 ≤ 16.7 ms；pointerup 仅一条 canonical mutation |
| Chart source 单格更新 | 只更新命中 series/layout；不重读所有 sheet drawings |

### 7.2 Backend

| 场景 | 目标 |
|---|---|
| 单格 commit，历史 1M operations | operation query 只读取 checkpoint tail；无全历史 scan |
| checkpoint 判断 | 一次 tail stats/read；不重复读取 operation log |
| snapshot at revision | 从最近 checkpoint 读取明确 revision 区间，顺序流式重放 |
| JDBC/REST query | 首个 block ≤ 500 ms（源允许时）；heap 与 block size 有界 |
| query cancel | 能调用底层 request/statement cancel；取消后不提交 data region |
| manifest commit | 每个 block 的存在、checksum、byteLength 在 workbook lock 内验证 |

### 7.3 Import/Export/Print

- ZIP/XML/cell/object 预算在 materialize 前执行。
- 100 MB archive 或超预算 XML 必须在显式错误中停止，不能 OOM 后由 UI 恢复。
- print page 只构建可见 page projection；不把完整 workbook 绘制到一个 Canvas。
- 未修改原文件保存保持未知 part；修改已拥有字段只 patch 对应 owner part。

## 8. Fail-close 与 Clean-break

每个 Issue 必须同时交付成功路径和拒绝路径：

- Feature dependency 缺失：`FEATURE_DEPENDENCY_UNAVAILABLE`。
- 可见入口无实现：surface disabled；执行返回 `UNSUPPORTED_FEATURE`。
- Formula Worker version/generation/desync：丢弃结果并要求 resync，禁止旧结果回写。
- Rule overlap/source unavailable：事务不产生任何 mutation。
- Native part 无法安全保存：返回 part/node/location/recovery，禁止输出有损文件。
- History gap/rebase failure：pending operations 保持未应用，base revision 不推进。
- Query/block mismatch：manifest 不落库，已上传孤儿块进入可回收状态。

Clean-break 要求：

- 删除 `document.import/export` 空处理器或把所有消费者迁入唯一 transaction。
- 删除 notice-only Forms/Screenshot 执行路径；不保留旧按钮 fallback。
- 删除 Formula full-snapshot-per-task 协议，运行时只接受新版本。
- 删除后端无界 `listOperations(unitId)` 业务读取路径。
- 删除页面内固定 B2:B21 validation builder 和自建规则公式 evaluator。
- 删除 Drawing/Chart/Print 的第二套独立值/布局读取链。

## 9. 实施顺序与 TODO

### Phase 0：计划和 Issue

- [x] 审计当前 `main`、Univer OSS、可读 Pro 包和既有 #295/#304-#317。
- [x] 记录功能矩阵、真实缺陷和目标架构。
- [x] 为 UCP-01..UCP-14 创建 GitHub Issue，并回填链接。
- [x] Issue body 包含代码证据、整改方案、依赖、成功/拒绝验收和 Blocked。

### Phase 1：统一架构开发批次（完成前不执行全量编译/测试）

本阶段是一个完整 clean-break，不交付局部补丁，不允许新旧协议并存。实现可以按文件所有权并行，但合入的是同一个 canonical runtime：

- [x] UCP-01 FeatureRuntime、生命周期、manifest compiler 与唯一 Home/Insert schema。
- [x] UCP-02 sparse range/selection/PaneMap/render indexes。
- [x] UCP-03 stateful formula Worker、delta protocol、dependency indexes、dynamic refs 与 lazy arguments。
- [x] UCP-04 number-format AST 与统一 display context。
- [x] UCP-05 canonical resolved-cell/rule/filter/table/data-region runtime。
- [x] UCP-06 review/link/find index 与 OOXML review transaction。
- [x] UCP-07 DrawingML、Shape/Connector/In-cell Graphics 与统一 object renderer。
- [x] UCP-08 native transaction、资源预算、未知节点 ownership 与 backend import authority。
- [x] UCP-09 print renderer、Unicode PDF、fit/repeat/object/resource。
- [x] UCP-10 analytics source/worker/layout/cache/OOXML。
- [x] UCP-11 connector config、range preprocess、block interval、列式 query/manifest commit。
- [x] UCP-12 history gap/rebase/import/offline fail-close。
- [x] UCP-13 bounded operation tail：repository 增加 `(unitId, fromExclusive, toInclusive)` 有序查询；commit/read/checkpoint/revision replay 只消费明确区间；删除无界 `listOperations(unitId)` 业务 API。
- [x] UCP-14 性能指标、真实文件 corpus 和浏览器观测点随实现一起埋入，但不在本阶段频繁运行。

### Phase 2：一次性验证与缺陷闭环

- [x] 架构审查：证明旧命令、旧协议、旧读取器、旧 fallback 已删除。
- [x] 一次运行 frontend typecheck、unit、boundaries、contracts、mutation registry、acceptance matrix、build。
- [x] 一次运行 backend Maven tests 与数据库迁移门禁。
- [x] 一次运行性能 corpus；失败只修复本批次真实根因，不回退旧链。
- [x] 一次完成内置浏览器交互、console、network、截图和性能验收。
- [ ] 一次完成真实 Excel/WPS corpus；桌面 Excel 不可用项明确标为 `Blocked`。

### Phase 3：统一提交

- [x] 更新全部 UCP Issue 的实现证据与剩余 Blocked。
- [ ] 提交、推送同一 `codex/univer-engine-convergence` 分支。
- [ ] 创建一个记录完整 clean-break、验证证据和回滚方式的 GitHub PR。

本轮浏览器回归补充：选中 PivotTable 后，Analyze/Design contextual tabs 现在由 `pivot` manifest 提供真实命令面，不再出现空白工具带；字段列表首次进入默认使用四区 2×2，Filters、Columns、Rows、Values 同时可见。此前打印公式值在浏览器预览中为空的缺陷也已通过统一 calculated-cell reader 修复，预览显示计算结果而不是空字符串。

## 10. 验证矩阵

### 10.1 静态与单元门禁

- frontend typecheck、unit、boundaries、contracts、mutation registry、acceptance matrix。
- backend Maven tests、数据库三方言 migration/profile tests。
- 每个 clean-break 验证旧协议/旧字段被拒绝。

### 10.2 内置浏览器

- 登录/打开真实 workbook。
- Home/Insert 每个入口点击或 disabled reason。
- selection、editing、filter、validation、CF、comment、hyperlink、find/replace。
- chart/pivot/sparkline/shape/in-cell image/print。
- console 无未处理异常；network 检查 request、status、payload、取消和错误响应。
- 100k/1M corpus 记录 frame、long task、Worker task、projection rebuild。

### 10.3 真实文件 corpus

- Excel/WPS 保存的 XLSX/XLSM/XLSB/BIFF 样本。
- formulas、number format、validation、CF、filter/table、comments/notes、drawings、chart/pivot/sparkline/outline、print setup、external connection、unknown ext。
- import -> edit -> export -> reopen；未拥有对象必须保持或显式 Blocked。

### 10.4 Blocked 定义

- 当前机器无法调用桌面 Excel 时，Excel 最终 reopen/visual comparison 标记 `Blocked`，不得用库自导自验代替。
- Pro exchange server、Rust native/WASM 或内部后端源码不可获得时，只能吸收公开 contract；不得宣称复刻内部算法。
- `0.25.1` Pro 包与 `1.0.0-beta.2` checkout 不同版本，不允许混装运行验证。

## 11. Definition of Done

整个计划完成需要：

```text
功能入口真实可用
  + 单一模型/命令/渲染/持久化/OOXML/后端链
  + 局部成本由 dirty range/closure/object/block 决定
  + 成功和拒绝路径测试
  + 内置浏览器 console/network/interaction
  + 真实文件互操作或明确 Blocked
  + GitHub Issue 与 PR 证据完整
```

单个 Issue 只有在其完整纵向切片达到上述适用门禁后才能关闭；部分 groundwork 必须保持 Issue open，不能以绿色构建提前关闭。
