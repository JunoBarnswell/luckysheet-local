# Calculation domain: ownership and acceptance

## Canonical flow

Authored input belongs to WorkbookModel (worksheet cells, scoped names and data
definitions). Commands produce the only mutation delta. Java authorizes and
commits it, and the client replays the same committed delta before calculating.
FormulaEngine owns AST evaluation, dependency indexes, dirty roots and spill
results. The Worker carries those inputs and revisions, never a second persisted
workbook. Rendering and editing must resolve the same source address.

The October 2026 acceptance run exposed failures at these boundaries:

* Java assigns the full restored worksheet to `sheet.restore`; TypeScript
  declares an empty range and rejects the successful server ACK.
* A sparse 12-sheet aggregate materializes 1.2 million blank values even though
  its interval index has only one posting. An edit used 1.7–2.5 seconds and
  increased the heap by approximately 812 MB in the measured Node run.
* Quoted 3D and external qualifiers enter the ordinary worksheet-reference
  parser. Local sheet rename can also rewrite an external sheet qualifier.
* TableSheet render and commit independently translate columns by ordinal,
  ignoring the persisted visible field order and sort definition.
* External references have no authorized value resolver; formula columns,
  Lookup and Rollup have no canonical record calculation owner.

## Reference consumption

A reference remains geometry until its consumer chooses an access mode.
Rectangular functions (FILTER, SUMPRODUCT, INDEX) require shape and blanks.
Sparse aggregates (SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, PRODUCT and variance)
consume occupied input/spill values without allocating blank coordinates.
Both modes use the same cell evaluator and permission-projected inputs. Sparse
reading is an index over canonical inputs, not another value store. Explicit
cell overrides and spill projections must participate in sparse reads.
Authored and permission-projected error constants are AST literals. In
particular, Java's `=#BLOCKED!` projection must evaluate and render as
`#BLOCKED!`, rather than losing its permission identity as a parse failure.
Deleted geometry remains the distinct `invalid-reference` AST node.

LET and LAMBDA belong to the existing AST/evaluator. Local bindings form a
lexical environment; named functions use the existing defined-name owner and
dependency collection. Formula text must never become a string macro.
Defined names return evaluator operands, retaining reference geometry and
lexical closures until the calling function consumes them. The same evaluation
context carries numeric/date semantics, tables, spills and explicit overrides.

## Identity and structure

Local worksheet IDs own references. Names are serialization identities.
3D references retain an ordered interval, with endpoint deletion/reorder owned
by the structural transformation domain. External references are a different
identity space: a local rename must never rename an external source sheet.
Quoted qualifiers are parsed by the existing parser and formatted canonically.

Worksheet restore/remove scopes are full worksheet geometries. They must match
in command declarations, inverse mutations, generated permission contracts,
Java authorization and committed replay. Validation must be possible after an
ACK has already changed the local model.

## External link owner

ExternalLinkDomain must own stable source identity, coalesced range requests,
source/cached revisions and connected/refreshing/stale/denied/unavailable/broken
states. FormulaEngine consumes data-only authorized cache revisions and never
performs HTTP. Access denial clears user-specific cached values and produces
`#BLOCKED!`; unavailable sources may retain an explicitly stale authorized
cache. Changes dirty indexed external owners and synchronize Worker deltas.
Bindings require canonical mutations and an explicit snapshot migration. A
successful snapshot read is insufficient proof of external formula execution.

## Record and field owner

The canonical record address is `(tableId, recordId, fieldId)`. View sort, filter
and grouping change projection positions only. Raw records belong to the
existing block source/store, outside WorkbookSnapshot. Formula fields share
the formula parser, graph and Worker. Relations refer to record IDs; Lookup
and Rollup use forward/reverse indexes, not a scan for every cell. Worksheet
backed tables remain explicitly worksheet backed until an explicit migration
creates stable records; row coordinates must not be advertised as record IDs.

TableSheet projection has one domain resolver for visible column, sorted source
row, render address and commit address. Unsupported calculated/lookup fields
must be rejected observably until the record calculation owner exists; silently
showing a blank or overwriting a source ordinal violates this contract.

## Acceptance

Run parser/evaluator/Worker tests, success and rejection mutation tests, shared
TS/Java vectors, full Java checks, and frontend gates after the implementation
pass. Browser acceptance uses the real Java/H2 service: command → POST → ACK →
replay → dependent values → reload. Capture screenshots, console and network
results. Permission acceptance must prove hidden source values do not enter
the browser, denied mutations retain the revision, and external sources require
independent authorization. Performance results include workload, occupied
count, interval postings, elapsed time and memory; three samples are not p95.
Native desktop Excel interoperability remains Blocked without desktop Excel.

This document is a target ownership contract. It does not certify unimplemented
ExternalLinkDomain or record graph capabilities as complete.

## 下一步实现所需的完整契约

下面是待实现的设计，不是当前运行能力。

| 地址域 | 持久化身份 | 值的所有者 | 视图变化的影响 |
| --- | --- | --- | --- |
| Worksheet | workbookId / sheetId / row / column | Worksheet CellMatrix | 排序等结构变换由结构域重写引用 |
| External | linkId / sourceWorkbookId / sourceSheetId / geometry | 经源工作簿授权的 ExternalLinkDomain 缓存 | 本地工作簿改名不能更改外部身份 |
| Record | tableId / recordId / fieldId | BlockStore 原始字段与 FormulaEngine 派生字段 | 排序、筛选、分组只改变位置 |
| Relation | relationId / sourceRecordId / targetRecordId | Record 域的正向、反向关系索引 | 删除记录由关系域原子处理 |

### ExternalLinkDomain

1. 在显式快照迁移中加入外部链接定义。公式中展示名称到 linkId 的
   绑定必须由规范命令完成，不能用同名工作簿猜测身份。
2. Java 以当前用户分别检查目标工作簿和源工作簿访问权限。批量读取
   按源工作簿、工作表合并区间；返回 sourceRevision、accessRevision
   和授权投影后的数据。公式求值器不发送网络请求。
3. 缓存键包含主体与访问版本。授权撤销先清除数据，再发送 Worker
   invalidation；旧请求在 accessRevision 变更后不得回填。
4. 请求中的 sourceRevision 必须与缓存值属于同一版本。一个刷新批次
   在单一提交点发布，不能让公式读取一半新值、一半旧值。
5. connected → refreshing → connected/stale/denied/unavailable/broken
   由链接域发布。denied 产生 #BLOCKED! 并删除缓存；离线 stale 只保留
   已授权缓存并显示来源版本；首次无缓存不能返回空值或零。
6. 对每个 linkId 建立外部区间依赖索引。只使相交公式和下游变脏；
   Worker 消息携带链接版本与变化区间，过时返回值不得覆盖新版本。
7. OOXML externalLinks 的导入/导出保留工作簿、工作表与引用身份；
   未能执行的外部格式可以保留，但必须明确报告能力未支持。

### Record / Field / Relation

1. 表字段的公式属于字段定义，不属于 TableSheet 的列定义。视图列
   只引用 fieldId。记录创建时分配稳定 recordId，不能用行号代替。
2. 共享 AST 增加带 table/field 身份的字段引用。只有 Record 求值上下文
   能解析 `[字段]`；字段改名更新展示文本但保留 fieldId。
3. 依赖键同时包含记录与字段。字段公式修改发布字段版本，原始值修改
   发布记录字段 delta。强连通分量负责跨字段、跨关系循环错误。
4. Lookup 通过关系正向索引读取目标字段；目标字段变更通过反向索引
   传播到受影响记录。Rollup 复用同一聚合函数与错误规则。
5. Java 原子提交记录、关系、字段定义变更，并验证 Record/Field ACL。
   关系目标不可见时不能暴露其 ID、计数或旧缓存值。
6. Worksheet-backed table 转成 Record table 必须显式迁移：一次分配
   recordId、校验区块摘要、重写视图绑定，并提升规范版本。普通读取
   不进行隐式升级。撤销、远程回放与刷新都使用同一地址域。

### 验收矩阵的扩展

当前执行的语义探针包含 50 项；真实浏览器覆盖跨 Sheet 公式、源值
修改、保存重开、TableSheet 排序/隐藏列/回写以及权限拒绝。后续必须
增加以下能力验收，当前不能把它们记为通过：

| 能力 | 必需成功路径 | 必需拒绝或失效路径 |
| --- | --- | --- |
| 外部链接 | 单元格/区间/名称/Table、刷新、重开、源改名 | 无绑定、源删除、离线首次加载、权限撤销、过时请求 |
| 外部传播 | 源值与源公式变化只重算相交依赖 | 外部循环、两源版本混用、Worker 返回版本过旧 |
| 3D 生命周期 | 两端删除、端点跨越、插入和内部移出 | 未支持结构变换须在提交前拒绝 |
| Record 公式 | 数值/日期/布尔、字段改名、排序后编辑 | 错误类型、循环依赖、字段删除、不可见字段 |
| Lookup / Rollup | 一对多、多对多、链式关系、增删目标 | 关系环、无权限目标、断裂关系、未授权计数 |
| 多维协同 | 两用户编辑不同 recordId，排序不改变编辑对象 | 记录删除与编辑并发、基版本过旧、权限版本过旧 |
| 大数据 | 百万记录区块读取、可视范围计算、索引内存 | 每次全表扫描、全量 Worker 快照、缓存越权 |

实测与期望不同的案例必须保留原始公式、输入、源身份、版本、错误码
和截图。当前 3D 端点删除被规范结构域拒绝，不能把这种拒绝写成 Excel
的 #REF! 或端点收缩行为已经实现。
