# Pivot Field List 与高性能 Pivot Runtime PRD

## 产品目标

Pivot 的字段配置必须复制 Excel 官方 Field List 的交互语义，同时保持浏览器主线程可交互：

- 上方 Field section：搜索、勾选字段、拖拽字段、字段菜单指定目标区域。
- 下方 Areas section：固定 Filters / Columns / Rows / Values 四区，支持拖拽重排和字段菜单。
- 文本字段默认进入 Rows，数字字段进入 Values，日期字段进入 Columns。
- 大数据源默认启用延迟布局更新；字段编排只修改本地 draft，用户点击 Update 后执行一次 canonical `pivot.update`。
- 格式和显示选项不是字段编排主流程，默认折叠，不能挤占 Field/Areas 工作区。

Microsoft 官方依据：

- https://support.microsoft.com/en-us/excel/get-started/use-the-field-list-to-arrange-fields-in-a-pivottable
- https://support.microsoft.com/en-us/excel/design-the-layout-and-format-of-a-pivottable

## 当前根因

当前 Pivot 计算仍以 `SourceRow[]` 承载行对象，每行再用 `Record<fieldId, PivotScalar>` 保存字段值。布局变化会重复执行：

```text
worksheet cells
 -> SourceRow objects
 -> field member catalogs
 -> filters
 -> row/column grouping
 -> aggregate arrays
 -> dense result cells
 -> dense PivotGridProjection cells
```

这会产生大量短命对象、重复字符串键、重复 member key 编码和完整笛卡尔矩阵。React 侧即使减少 render，也无法消除计算和 GC 长任务。

### 真实附件基线（2026-08-27）

强制验收文件为 `C:\Users\kuo13\Downloads\OCR结果.xlsx`：

- `Sheet1!A1:W4059`，23 列、4,059 行、93,357 个 OOXML cell records。
- 文件内全部单元格均为 shared-string `t="s"`，style `0`，number format `General`；原生 XLSX 导入不得把 `4.60`、`260` 或 `22/07/2026` 按内容猜测为 number/date。
- 内置浏览器导入完成约 1.36 秒，边界单元格 `W4059` 可达且值正确。
- 同一范围创建 PivotTable 的现有交互曾超过 55 秒仍未形成字段目录；失败没有 console error，也没有 typed UI error。该行为同时违反交互预算和 fail-close。

OCR 内容类型修复若未来启用，只能是用户显式选择 schema 后运行的一次性数据清洗事务。它不属于 XLSX import、PivotSourceIndex 或 Pivot worker，且不得成为 native OOXML 的 fallback reader。

## 状态与所有权

| 状态 | 唯一所有者 | 生命周期 | 禁止事项 |
| --- | --- | --- | --- |
| `PivotDefinition` | `WorkbookModel` | history / persistence / collaboration | 不保存 worker 状态、source buffers 或派生结果 |
| `PivotSourceIndex` | workbook + canonical source revision cache | source revision | 不复制到 React state，不按 layout 重建 |
| `PivotTaskState` | `WorkbookSession` | task generation | 不持久化，不用 notice 冒充 error state |
| `PivotResultTree` | session result publication | source/layout/filter revision tuple | stale generation 不得发布 |
| `PivotGridProjection` | render projection cache | result revision + viewport | Canvas 不得触发 source scan 或聚合 |

创建/更新数据流必须为：

```text
UI intent
  -> validate source/destination
  -> acquire revision-owned PivotSourceIndex
  -> submit PivotTaskRequest
  -> worker validate + aggregate + physical pivot
  -> revision/generation match
  -> one canonical command transaction
  -> publish result/projection
```

失败发生在 canonical command 之前时，不得创建 worksheet、PivotDefinition 或 history entry；失败发生在已有 Pivot 刷新时，保留 last-valid result，并发布 typed task error。

## 已落地的第一阶段所有权

- Field List 默认使用官方 stacked 布局。
- 大于 50,000 个源范围单元格或 block data source 默认启用 deferred layout update。
- Pivot source table 和 field catalog 由 workbook/source revision cache 单一持有；布局变化不再重复重建源行与 member domain。
- 源 revision 变化会替换缓存，同一 workbook 最多保留 8 个 source entries，避免无界内存增长。

## 目标运行时：`PivotSourceIndex`

现有 `SourceRow[]` 将 clean-break 为一个列式、revision-owned source index：

```text
PivotSourceIndex
  sourceRevision
  rowCount
  fields[]
    fieldId
    type
    validityBitmap
    dictionary[]          // text/error/date member domain
    dictionaryCodes      // Uint16Array / Uint32Array
    numericValues        // Float64Array where applicable
    booleanBits
  rowPathPool
  sourceRowPathIds
```

设计采用 Arrow 的列式、字典编码和连续 buffer 原则，但不引入第二套数据模型；`PivotSourceIndex` 是 worksheet/data-source 到 Pivot worker 的唯一派生边界。

Apache Arrow 依据：https://arrow.apache.org/docs/format/Columnar.html

## 目标算法：Aggregate Once, Physical Pivot Once

DuckDB 的高性能 Pivot 会先对 `(group keys, pivot keys)` 聚合一次，再通过 list aggregate 与 `PhysicalPivot` 完成转置；不会为每个输出列反复扫描源数据。

目标流水线：

```text
Stage 1: source scan
  dictionary ids + validity/numeric vectors
  -> filters and slicer/timeline masks

Stage 2: grouped aggregation
  composite row-key ids + column-key ids
  -> open-addressing aggregate hash table
  -> typed aggregate states

Stage 3: physical pivot
  sparse (rowGroupId, columnGroupId, valueStates)
  -> ordered row/column dictionaries
  -> PivotResultTree

Stage 4: viewport projection
  PivotResultTree sparse cells
  -> row-major coordinate query
  -> visible PivotGridProjection cells only
```

策略根据实际 cardinality 选择：

- 少量 pivot columns：filtered aggregate states。
- 大量 pivot columns：一次 group-by 后生成 aligned lists，再执行 Physical Pivot。
- 高 group cardinality：开放寻址 hash table；payload 与 slot table 分离，达到负载阈值后只重建 slot table。
- 大输入：按 hash 高位做 radix partition，使工作集保持在 cache-friendly 分区中。

DuckDB 依据：

- https://duckdb.org/docs/current/internals/pivot.html
- https://duckdb.org/2022/03/07/aggregate-hashtable.html
- https://github.com/duckdb/duckdb/pull/6961

## Worker 与取消合同

所有 Pivot 计算最终统一进入 `PivotTaskProtocol`：

```text
main thread
  canonical PivotDefinition
  PivotSourceIndex buffers
  source/layout/filter revisions
        |
        | transferable ArrayBuffer / SharedArrayBuffer
        v
Pivot Worker
  grouped aggregation
  physical pivot
  sparse result
        |
        v
main thread
  revision match -> publish
  stale generation -> discard
```

- Worker 计算期间 Field List、scroll、selection 保持响应。
- 新布局提交会取消旧 generation；旧结果不得覆盖新布局。
- 结果通过 transferable buffers 返回，不结构化克隆完整 row objects。
- 不保留 synchronous fallback runtime；测试宿主使用同一个 task evaluator 的 inline port。

### `PivotTaskError` 合同

任务失败必须至少包含：

```ts
interface PivotTaskError {
  code:
    | 'PIVOT_SOURCE_INVALID'
    | 'PIVOT_SOURCE_UNAVAILABLE'
    | 'PIVOT_MEMBER_LIMIT_EXCEEDED'
    | 'PIVOT_TARGET_COLLISION'
    | 'PIVOT_TARGET_BOUNDS_EXCEEDED'
    | 'PIVOT_RESULT_LIMIT_EXCEEDED'
    | 'PIVOT_TASK_CANCELLED'
    | 'PIVOT_TASK_TIMEOUT'
    | 'PIVOT_TASK_PROTOCOL_ERROR'
    | 'PIVOT_PERMISSION_DENIED'
    | 'PIVOT_TASK_FAILED';
  message: string;
  pivotId: string;
  sourceIdentity: string;
  sourceRevision: string;
  recovery: 'fix-source' | 'change-layout' | 'change-target' | 'retry';
}
```

UI 必须显示该失败；console/network 只用于诊断，不是产品错误通道。异常不得被 catch 后转换成 `undefined` 或仅写入瞬时 notice。

DuckDB-Wasm Worker/Arrow 依据：https://www.vldb.org/pvldb/vol15/p3574-kohn.pdf

## 内存合同

- Source index 每个 source revision 只存在一份。
- 字符串只在 field dictionary 中保存一次；行数据保存 integer codes。
- Aggregate state 按实际 group 数分配，不按 `rowCount × columnCount` 分配。
- `sourceRowPaths` 进入池并以 ID 引用，禁止在每个 subtotal/value cell 复制完整数组。
- Pivot presentation 变化不重算 source/filter/aggregate。
- 旧 revision 在新结果发布或任务取消后释放。

## 验收预算

### 4,059 × 23 OCR workbook

- XLSX import：P95 < 2,000ms。
- 首次 source index + field catalog：P95 < 300ms；任何连续主线程 task < 50ms。
- 创建空布局 PivotTable：P95 < 500ms，并形成 worksheet、Field List 和可观察 task state。
- 勾选/移动字段（deferred）：P95 < 50ms。
- 点击 Update：主线程提交 < 50ms；Worker 完成 P95 < 500ms。
- source revision 未变化时，第二次字段目录读取 P95 < 10ms。
- source index 复用后的第二次布局计算不得重新遍历 93,357 个 worksheet cells。

### 100,000 × 20 dense workbook

- Field List 编排保持即时响应。
- Pivot task 不在主线程执行聚合。
- source index 与 aggregate peak memory < 等价 row-object runtime 的 35%。
- 取消旧任务并发布新布局结果 < 2s。
- Worker timeout、message error 或 revision mismatch 必须产生 typed failure，不得同步回退到主线程。

### 高基数与拒绝路径

- 超出 worksheet bounds 在 materialize projection cells 前拒绝。
- 不创建 dense Cartesian cell array 后再报错。
- 失败不修改 PivotDefinition、history 或 last-valid result。

## Clean-break 约束

- `PivotSourceIndex` 落地时同一变更删除 `SourceRow.values` 主计算链；不双写 row/column 两种模型。
- Worker 成为唯一计算所有者后删除 UI/session 同步计算入口。
- Sparse result projection 落地时删除 dense `projection.cells.find`/全量 materialization 路径。
- 迁移只发生在 worksheet/data-source → source index 边界，不能进入 Canvas 或 React 组件。
