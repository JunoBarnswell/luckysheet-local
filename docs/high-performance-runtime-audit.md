# High Performance Runtime Audit

审计基于分支起点 `48a64dfda6f657c873ab0df914a2de8800485d1d`（当时与 `origin/main` 一致）的可执行代码；规划文档未用作现状证据。

## Runtime Architecture Audit

### Calculation Flow

`WorkbookSession` 将编辑转换为命令；`CommandRuntime` 校验并提交规范化 Mutation，`spreadsheet-app/src/runtime.ts` 通过 `synchronizeCellMutation` 把受影响单元格增量同步至 `FormulaEngine`。`FormulaEngine` 根据 `RangeIndex` 收集脏公式并求值。浏览器默认通过 `CalculationTaskPort` 使用持久 Worker；Worker 首次/上下文代际变化接收计算快照，普通编辑接收输入 delta。UI 快照消费已算出的公式值。

计算所有权仍主要是单元格公式。Workbook 名称、表公式字段、Lookup、Rollup、外部绑定、报表表达式和排程尚未汇入同一计算图。

### Formula Dependency Flow

`parser.ts` 生成正式 AST；`dependencies.ts` 收集 cell、range、name 和 reference 依赖；`FormulaEngine` 将依赖放入 `RangeIndex` / `ReferenceIndex`，通过索引查找受影响公式。动态数组以 anchor/spill 记录，已有增量输入更新和 generation 检查。

普通重算调用 `getCircularComponents().filter(...)` 遍历当前 SCC 列表；SCC 缓存按拓扑 generation 建立，拓扑变化时从公式 Cell 重建图。前者在每次重算仍与 SCC 数量相关，后者在拓扑变化是全公式成本。

### Sheet Reference Flow

单表引用用解析后的 sheet identity 进入 `resolveFormulaSheetId`；范围由单 sheet geometry 表达。Formula AST 已包含 `SheetRangeReferenceNode`。ReferenceIndex 的 3D 几何当前展开为每张工作表的空间 posting；`FormulaEngine.resolveReference` 尚未执行该节点。

### Structural Reference Flow

工作表坐标权威为 `core-model/src/structural-transform.ts` 的 `StructuralTransform`，以 `StructuralPatch` 描述公式与范围 owner 的前后值。命令预检、协同 replay 和 workbook structural APIs 消费该事实；结构 owner 索引由 `RangeIndex` 支持。工作表 identity/order 由 `WorkbookModel` 持有，但 FormulaEngine 的 `updateSheetNames` 明确拒绝 identity 顺序变化。

### External Reference Flow

`parser.ts` 和 AST 可保留 ExternalReference；结构改写与序列化识别此类型。`FormulaEngine.resolveReference` 对它返回 `#REF!`，仓库没有 ExternalLinkDomain、range binding cache、refresh 或源工作簿 ACL 读取链。

### Command / Mutation Flow

UI action → `WorkbookSession` / feature command → `CommandRuntime` 预检和执行 → WorkbookModel canonical mutation → `MutationInfo` → undo/history/collaboration/persistence。命令 registry 与 protocol mutation capability 对可传输操作做约束；禁止 renderer 直接修复业务状态。

### Collaboration Flow

`CollaborationSession.enqueueLocalMutations` 将命令产生的 Mutation delta 包装成单一 `OperationEnvelope`，进入持久 `OfflineQueue` 并通过认证 REST 传输；提交后以 revision ACK 确认，远端重放使用 `CommittedOperationEnvelope` 和 StructuralPatch，OT 按 Mutation 类型 rebase。WebSocket handler 只接受 presence/cursor 临时消息，不接受 workbook 操作。

`assertNextRevision` 对 revision gap 明确抛错；session 本身没有自动取得权威快照并恢复的闭环。协同分类目前主要基于 cell/range/worksheet/pivot/name 等 Mutation，尚无 Table+Record+Field、Task+Field 等逻辑对象冲突身份。

### Revision Flow

协同 `baseRevision`、FormulaEngine 的 `calculationGeneration`、`calculationContextGeneration` 与 Worker `inputRevision` 分别驱动 replay、过期结果拒收、上下文重建和输入 delta ACK。公式拓扑代际是内部缓存字段，尚无统一可观测的 `modelRevision` / `formulaTopologyRevision` 合约。

### DataModel Flow

`core-model/src/data-model.ts` 定义 `WorkbookDataModel`、Table/Field/Block、Relationship、View、Gantt 和 Report 定义。`WorkbookModel` 保存其规范定义；Snapshot 序列化元数据。`DataBlockStore` 和 `DataBlockSynchronizer` 是大块数据持久化入口，但当前 Table View/Gantt/Report/Analysis 部分投影仍依赖 worksheet-backed `sourceRange`。

### Table / Record Flow

普通 Table 定义稳定 `tableId` 与 `fieldId`，Block 有 storage key，但没有规范 Record API / `recordId` revision 写入链；关系定义目前是模型元数据。Gantt task ID 从映射字段读取，投影同时保留源 row 地址。关系查找和 Rollup 没有 workbook data-model runtime。

### Gantt Flow

`features/gantt/projection.ts:buildGanttProjection` 从表绑定和 Canvas sheet 一次读取全体任务，生成完整 `tasks` 数组；层级与依赖 cycle validation 在投影时遍历任务图。依赖仍是文本拆出的 `string[]`，没有 FS/SS/FF/SF、lag 或增量 ScheduleEngine。Renderer 接收已物化任务投影；当前投影 API 没有 viewport 参数。

### Report Flow

`features/report/projection.ts:buildReportProjection` 校验 `ReportBinding`，从 sheet/table 读取值并物化全部重复行的 `cells`；页数按固定 `rowsPerPage` 推算。尚无 band layout、共享 Preview/PDF/Print PageLayout 或按页懒生成接口；formula binding 目前保留表达式文本。

### Query Flow

`features/query` 有 JSON/CSV/TSV/OOXML/REST connector 与数据源命令；`data-source/content-query.ts` 对规范区域做内容投影。查询能力与 WorkbookTable 的关系/公式字段没有统一 calculation graph。

### Pivot Flow

`features/pivot/engine.ts` 从 sheet/table/source 读取行、连接关系、计算字段并建立 Pivot 树；异步任务 port/Worker 已存在，输入以 source descriptors 和 revision/generation 标识。部分 source table 和 join 会构建完整行集合及 lookup Map，尚未成为通用 Projection Graph。

### Snapshot / Persistence Flow

WorkbookSnapshot 保存工作表与 canonical 元数据；WorkbookCatalogService 调用服务器 Workbook API；协同 operation journal 持有未 ACK Mutation。数据块和图片资产各走独立内容/块存储，Snapshot 引用它们。报表页、Gantt geometry、Analysis chart points 当前不作为 canonical Snapshot 字段。

### Performance Hot Paths

- Formula 重算的循环组件筛选扫描 SCC 列表；拓扑变更重建 SCC 时扫描公式 Cell。
- 3D ReferenceIndex 每个跨表引用为每个 Sheet 建立普通二维几何 posting，reorder 无专属 3D owner invalidation。
- Gantt validation 对任务依赖递归遍历且每层复制 path Set；投影读取全任务。
- Analysis projection 先复制全部行/字段，再过滤并为每个 chart 再生成点集。
- Report projection 重复扩展所有数据行并保留全部 cells；翻页没有页级 seek。
- Pivot 在关系连接、分组和计算项路径创建完整中间行/树结构；有 Worker 隔离但不代表增量计算。

### Duplicate Ownership

- Worksheet Cell Matrix 和 WorkbookTable Block 是两种数据表示；worksheet-backed projection 通过 `sourceRange` 读取，block-backed 表尚无相同 projection reader，因此不能声称统一 Record canonical store。
- Gantt Definition 持有字段映射而任务仍从 worksheet 值解释；没有独立 Task Record identity API。
- Analysis/Report/Gantt 各有专用全量投影代码；没有共用 Projection Graph 或 revision cache。
- `FormulaEngine` 在主 workbook runtime 之外还被 Pivot、What-if、History preview 等调用方单独创建，需按调用场景审查是否产生额外图与大数据副本。

### Full-Scan Risks

公式 topology 重建、worksheet-backed Gantt/Analysis/Report 初次投影、Pivot full source table 与 Data View filter/sort 都随输入总体大小线性增长；除公式稀疏 range index / Pivot source index 之外，关系 lookup、rollup、排序视图没有增量索引。

### Memory Duplication Risks

Analysis rows + filtered intermediate + chart points 同时物化；Report cells 扩展成模板投影副本；Gantt tasks 从工作表字段再次形成对象；Pivot source rows/group tree 有全量工作集。大表 block 设计已存在，但主 projection 尚不能按 viewport 从 block reader 取数。

### Network Amplification Risks

普通 Cell collaboration 发送 Mutation params，而不是 Workbook Snapshot；Worker 常规输入使用 delta。外部 Workbook refresh 不存在。大表 Record Mutation 协议不存在；Gantt task 只能由表格/单元格 mutation 间接表达。REST revision replay 和 presence WebSocket 是分开的。

### Target Ownership

统一由 Workbook identity、canonical Record/Field、ReferenceIndex/Calculation Graph、StructuralTransform、Mutation/Revision、AccessControl 与 Projection Graph 承担跨域事实。Worker 接收 revision + delta；外部绑定、关系字段、排程、报表表达式注册为计算节点；UI 只保存 revision 和 viewport selector，Renderer 只消费窗口化 Projection。

## High Performance Runtime Audit

### 1. Current Calculation Architecture

有增量 FormulaEngine 和 Worker 传输；公式与表/关系/排程/报表计算仍分域。

### 2. Current Formula Reference Architecture

Cell/range/name/table 引用使用 AST + RangeIndex/ReferenceIndex；sheet-range/external AST 已存在。

### 3. 3D Reference Gaps

Evaluator 未实现；posting 按当下 sheet order 扩成多个 sheet geometry；sheet reorder 缺 3D owner 定向失效。

### 4. External Workbook Gaps

Parser/preserve 完成，运行时返回 `#REF!`；无 source identity/binding/cache/refresh/permission。

### 5. Current Collaboration Architecture

单一 OperationEnvelope + REST mutation journal/ACK + OT/StructuralPatch；presence/cursor 经 WebSocket；gap 抛错但无 snapshot recovery。

### 6. Current DataModel Architecture

Snapshot 中有 Table/Field/Block/Relationship/View 元数据，数据块独立存储；sheet-backed projection 仍是多数 read path。

### 7. Record Identity Gaps

Record ID/revision/typed field value 更新 API 与协同协议不存在；Gantt row projection 暴露行号。

### 8. Relation / Lookup / Rollup Gaps

只有关系定义，没有 canonical RecordId 引用索引或 Lookup/Rollup 增量求值。

### 9. Current Gantt Architecture

Definition + 全量 worksheet read + task object projection + cycle validation；无独立 schedule domain。

### 10. Gantt Scheduling Gaps

依赖为 string ID 列表；无边类型/lag/calendar scheduling/critical-path/增量 dirty propagation/viewport clipping。

### 11. Current Report Architecture

Worksheet template + binding projection + 固定行数分页统计。

### 12. Report Engine Gaps

没有 bands、正式表达式、scope aggregate、Master/Detail、lazy PageLayout 或共享打印布局。

### 13. Current Analysis Architecture

Persisted view/filter/chart bindings 从 sourceRange 读取并生成内存 rows + chart points。

### 14. Full Scan Hotspots

Gantt/Analysis/Report source rows、Pivot joins、view sorts 与公式 SCC 的部分查询按总数据集扫描。

### 15. Memory Duplication Hotspots

各业务 projection 复制 source rows/tasks/cells/chart points，block-backed canonical source 尚未被这些 projection 共用。

### 16. Network Amplification Hotspots

已有 cell delta 与 Worker delta；逻辑 Record/Task/View mutation 和外部 fetch dedup 还未建立。

### 17. Main-thread Hotspots

同步 Gantt/Report/Analysis projection 和部分公式内联 fallback 在调用方未使用 Worker 时会阻塞主线程；Pivot/计算 Worker 已部分承载重任务。

### 18. Duplicate Ownership

公式辅助 runtime、sheet vs block row source、业务各自投影的 ownership 缺少统一层；没有发现第二套外部公式解析器。

### 19. Target Architecture

本审计上文 `Target Ownership`；按需求阶段顺序渐进，不引入并行模型或兼容路径。

### 20. Destructive Refactors

移除 3D per-sheet 扁平依赖，改为 interval + range geometry；新增 Canonical Record API 后让 Gantt/Table/Analysis/Report 消费同一数据访问面；删去依赖全量投影的旧消费者时同 PR 更新全部调用点。

### 21. Performance Index Plan

3D sheet interval owners、formula SCC member reverse map、table record/field/relationship index、view ordered/filter index、Gantt forward/reverse adjacency、projection revision cache 与 viewport row index。

### 22. Benchmark Plan

公式：稀疏 1M cell/100k formula 单点编辑与 3D interval；数据：100k Record 单字段更新/Lookup/Rollup/filter/sort；甘特：10k/100k task 增量 schedule 与 viewport；报表：10k/100k 首屏/下一页；协同：2–100 客户端 payload/replay。每类记录基线与预算。

### 23. Implementation Order

本分支从 Phase 1 Calculation Graph ownership、Phase 2 3D Reference 开始；Phase 3 External Workbook、Phase 4–6 Record/Relation/Collaboration identity、Phase 7–9 Gantt/Report/Analysis、Phase 10 ACL integration、Phase 11 benchmarks/cleanup 仍需依序实施并逐项验证。

## Implementation Status (2026-10-02)

This status records changes made after the baseline audit above.

- **Phase 1, calculation ownership groundwork:** added explicit formula-topology and calculation-input revision accessors; cached cyclic formula members by address so affected-subgraph recalculation does not filter the full SCC list on each edit; made SCC traversal iterative to avoid call-stack failure on deep dependency chains.
- **Phase 2, 3D references:** `sheet-range-reference` now evaluates against live worksheet order. A 3D dependency is stored as one sheet-interval posting plus row/column geometry, rather than one spatial posting per worksheet. Point and range invalidation query the interval index; sheet reorder reindexes only formulas that own 3D references and refreshes dependent named formulas and formula rules. The existing `FormulaEngine` instance is retained across reorder.
- **Backend write serialization:** the per-workbook lock now remains held through transaction completion for commit, checkpoint, and restore. This closes a real race where the synchronized method body released its lock before Spring committed the JPA entity version update. Tests cover both commit completion and rejected-write rollback completion.
- **Real browser/backend evidence:** Chromium authenticated against the local Spring Boot service backed by file-based H2 (`jdbc:h2:file:/proof-data/luckysheet_canonical`). Each run created a remote workbook and committed sheet, cell, and formula mutations through the server API. `=SUM(Jan:Dec!B2)` rendered `60`; moving `Feb` outside the boundary interval rendered `40`; reloading the server workbook preserved the formula and result at revision 13. Four consecutive full browser/backend runs after the lock fix, including one after the final formula-engine edit, completed with zero failed requests, page errors, or relevant console issues. Three screenshots are retained in the task workspace.
- **Verified checks:** formula-engine index tests passed (29/29), focused application reorder tests passed (2/2), TypeScript typecheck, production build, and repository boundary/contract/acceptance gates passed, and the Java backend Maven suite passed (308 tests, 0 failures/errors/skips). The full `npm run test:unit` command remains red on the branch and at baseline: 150 reported failure markers/108 distinct failing names here versus 152/109 at the exact branch-start commit, with no new failing name introduced and the worker protocol test now passing. This includes `recalculates spills when sheet-table and merge blockers change geometry` (`Sheet Table identities must be unique within a workbook`), reproduced on the exact baseline commit.

Phases 3–11 and DoD items outside calculation, 3D references, and the stated runtime groundwork are not complete. The broad cross-workbook, Record/Relation/Rollup, collaboration recovery/logical identity, ScheduleEngine, band report, dashboard projection, unified domain access control, and benchmark CI requirements remain outstanding.
