# Codex 主任务：基于当前 main 完成 SDK 产品内核收口与 Web 薄壳整改

## 一、任务目标

对 `JunoBarnswell/luckysheet-local` 当前 `main` 分支执行一次基于真实代码的架构复核，并直接完成 SDK 架构整改。

本任务的最终产品定义是：

> `@react-sheets/sdk` 是在线 Excel 唯一产品内核。  
> `apps/web` 只是浏览器 UI 宿主。  
> Web 不拥有任何 Excel 业务语义、工作簿业务状态、认证凭证、权限计算、Command、Mutation、公式、协同、持久化、原生文档或渲染交互领域。

这不是新增一个 SDK 包，也不是在 `WorkbookSession` 外再包一层 API。

最终必须消灭双链路：

```text
SDK Workbook API
        +
WorkbookSession → Web Editor
```

统一成：

```text
Canonical Core
      ↓
SDK Product Kernel
      ↓
Web Thin Shell
```

---

# 二、代码基线

仓库：

```text
JunoBarnswell/luckysheet-local
```

分支：

```text
main
```

上一轮静态审查时 HEAD：

```text
3ff47ca6df8de4af49b809649831ea376179d24c
```

Commit：

```text
feat(sdk): object workbook API, verified identity and authorized multi-workbook calculation
```

执行任务时必须重新读取当前 `main HEAD`。

**当前 `main HEAD` 是唯一事实来源。**

禁止根据旧分析文档、旧 PRD、旧 Issue 或本提示词中的文件大小直接作结论。

本提示词中的数字与路径属于已发现的重点检查项；若当前代码已经变化，以当前真实代码为准。

---

# 三、GitHub 访问规则

优先使用 GitHub Connector 获取：

```text
repository
branch
commit
tree
source
diff
PR
Issue
```

Shell 中的：

```text
git clone
git fetch
```

只用于本地构建、测试、类型检查、静态检查需要。

如果 Shell 出现：

```text
Could not resolve host: github.com
Temporary failure in name resolution
Connection timed out
Failed to connect to github.com
```

不得判定 GitHub 不可访问。

必须切换 GitHub Connector 继续源代码审查。

不得因一次网络错误反复 clone。

只有真实需要本地构建，而当前运行环境持续不能获取代码时，才把对应验证项记录为环境限制。

---

# 四、当前已经确认的正确方向

下面这些方向已经在上一轮 `main` 中成立。

先重新确认当前代码仍保持这些语义，不要推翻正确设计。

## 4.1 `createSpreadsheetSdk()` 已经存在

当前已经形成：

```ts
createSpreadsheetSdk(...)
```

顶层结构接近：

```ts
interface SpreadsheetSdk {
  auth;
  identity;
  workbooks;
  dispose(): Promise<void>;
}
```

这一方向保留。

---

## 4.2 Auth / Credential ownership 已基本进入 SDK

已经存在：

```text
OIDC
external bearer
host session
local auth
CSRF
credential refresh
credential change event
verified auth context
```

公开 `AuthSnapshot` 不应泄漏：

```text
access token
refresh token
CSRF token
```

服务端已经出现：

```text
VerifiedIdentityService
VerifiedAuthContext
IdentityContextFilter
AuthenticatedSocketLifecycle
```

HTTP 和 WebSocket 使用 `contextId` 约束身份上下文。

这条链保持为唯一身份链。

---

## 4.3 WorkbookRole 已统一

此前多套：

```text
WorkbookAclRole
ShareRole
WorkspaceRole
WorkbookRole
```

当前已经开始统一成 generated：

```ts
WorkbookRole =
  | 'owner'
  | 'editor'
  | 'commenter'
  | 'viewer';
```

这条 canonical role contract 保持唯一。

禁止重新生成第二套同义角色。

---

## 4.4 Workbook Center 已迁入 SDK

当前：

```text
packages/sdk/src/workbooks/domain.ts
```

已经拥有：

```text
list
create
open
import
export
rename
copy
move
share
trash
restore
purge
spaces
folders
members
preferences
```

这部分继续收口，不允许迁回 Web。

---

## 4.5 Workbook Object API 已建立

当前已经有：

```text
Workbook
WorksheetCollection
Worksheet
RangeCollection
Range
CellCollection
Cell
DefinedName
WorksheetAxis
WorksheetProtection
WorkbookExternalLinks
```

并且 Object API 没有复制一套 Workbook 数据，而是通过 canonical runtime port 执行。

这一原则保持：

> 对象保存 identity 和 action，不保存第二份业务真相。

---

## 4.6 Dimensions 已迁入 SDK

当前已经存在：

```text
packages/sdk/src/dimensions
```

包括：

```text
domain
contract
autofit-worker
autofit-protocol
```

业务 Worker ownership 已从 Web 移出。

以后所有类似 Worker orchestration 都归 SDK。

---

# 五、当前核心问题

## 5.1 SDK 已建立，但旧链没有退出

上一轮代码中：

```text
@react-sheets/sdk
```

已经存在，但 Web 仍大量直接依赖：

```text
@react-sheets/spreadsheet-app
@react-sheets/core-model
@react-sheets/command-runtime
@react-sheets/render-engine
@react-sheets/protocol
@react-sheets/exchange-excel-ooxml
@react-sheets/sheet-features
@react-sheets/formula-engine
```

上一轮统计非测试 Web 源文件：

```text
spreadsheet-app          47
core-model               34
command-runtime          17
render-engine             8
protocol                  6
exchange-excel-ooxml      6
sheet-features            5
formula-engine            2
```

重新统计当前 HEAD。

最终所有这些 Web 直接依赖必须归零。

---

# 六、最终 Web dependency contract

最终：

```json
apps/web/package.json
```

业务依赖只能面向：

```text
@react-sheets/sdk
@react-sheets/ui-system
react
react-dom
```

Web 不得直接 import：

```text
@react-sheets/core-model
@react-sheets/command-runtime
@react-sheets/protocol
@react-sheets/formula-engine
@react-sheets/render-engine
@react-sheets/sheet-features
@react-sheets/exchange-excel-ooxml
@react-sheets/spreadsheet-app
```

如果 UI 需要某个类型、状态、动作或投影，证明 SDK 公共契约缺失。

正确动作是把它收进 SDK。

禁止让 Web 为了省迁移工作继续引用内部包。

---

# 七、当前最大的迁移缝：`useWorkbook()`

上一轮实现：

```ts
useWorkbook(sdk, resolution)
```

内部仍然创建：

```ts
WorkbookSession
```

并返回：

```ts
{
  session,
  snapshot,
  data,
  dimensions
}
```

这是错误的最终边界。

Web 一旦拿到：

```text
WorkbookSession
```

SDK 就没有成为唯一产品边界。

## 必须整改为

Web 输入：

```text
unitId
```

SDK 内部负责：

```text
resolve
auth
permission
persistence
runtime creation
session lifecycle
```

Web 得到：

```ts
{
  workbook,
  view,
  domains
}
```

或者同等语义。

Web 永远不能看到：

```text
WorkbookResolution
WorkbookSession
```

---

# 八、必须建立正式 `WorkbookView`

这是当前 Web 无法脱离 `WorkbookSession` 的根问题。

SDK 现在有对象式写 API，但是 Interactive Editor 缺少完整 reactive read model。

必须建立 SDK 正式 View Domain。

目标语义：

```ts
interface WorkbookView {
  getSnapshot(): WorkbookViewSnapshot;
  subscribe(listener: () => void): () => void;
}
```

`WorkbookViewSnapshot` 必须承载 UI 真正需要的只读投影。

包括：

```text
phase
workbook identity
workbook name
active sheet
sheet tabs
selection
active cell
formula bar
editing session projection
save state
collaboration state
peer cursors
permissions
range access projection
ribbon state
panel state
dialog state
backstage state
canvas projection
drawing projection
pivot projection
status bar projection
compatibility state
history state
```

Web 只能消费 SDK View Snapshot。

禁止 Web 直接消费：

```text
UiSnapshot
WorkbookSession.getUiSnapshot()
```

---

# 九、WorkbookSession 必须退出架构

上一轮：

```text
packages/spreadsheet-app/src/workbook-session.ts
```

已经约：

```text
412 KB
```

并且继续承担大量领域。

不要把它改名。

不要继续给它加方法。

不要创建：

```text
SdkWorkbookSession
WorkbookFacade
WorkbookSessionV2
```

## 正确整改

把职责拆为 SDK 内部正式领域：

```text
workbook/
  lifecycle
  view
  selection
  editing
  clipboard
  formulas
  names
  tables
  data
  validation
  conditional-formatting
  drawings
  charts
  pivots
  sparklines
  review
  permissions
  collaboration
  history
  document
  print
  viewport
```

每个 Domain：

```text
拥有自己的状态 ownership
拥有自己的公开动作 contract
通过统一 Command Runtime 写 canonical model
通过统一 View Projection 发布只读状态
```

最终删除 `WorkbookSession`。

---

# 十、禁止 SDK 变成 WorkbookSession 转发壳

错误结构：

```text
Workbook.cells.set()
  ↓
WorkbookSession.setCell()

Workbook.chart.add()
  ↓
WorkbookSession.addChart()

Workbook.pivot.create()
  ↓
WorkbookSession.createPivot()
```

如果只是把旧方法原样包进对象，这仍然是 God Object。

正确结构：

```text
Workbook
 ├─ cells → CellDomain
 ├─ selection → SelectionDomain
 ├─ editing → EditingDomain
 ├─ formulas → FormulaDomain
 ├─ tables → TableDomain
 ├─ data → DataDomain
 ├─ charts → ChartDomain
 ├─ pivots → PivotDomain
 ├─ drawings → DrawingDomain
 ├─ permissions → PermissionDomain
 ├─ collaboration → CollaborationDomain
 ├─ history → HistoryDomain
 ├─ document → DocumentDomain
 ├─ print → PrintDomain
 └─ viewport → ViewportDomain
```

Workbook 本身只拥有：

```text
identity
lifecycle
domain composition
transaction context
view publication
resource disposal
```

---

# 十一、统一写链

所有写操作只能通过：

```text
Web Gesture
    ↓
SDK Public Action
    ↓
SDK Domain
    ↓
Command Runtime
    ↓
Canonical Mutation
    ↓
Workbook Model
    ↓
Effects
    ├─ Calculation
    ├─ History
    ├─ Collaboration
    ├─ Persistence
    ├─ Projection
    └─ Native Document Ownership
```

禁止：

```text
Web → CommandDescriptor
Web → Mutation
Web → WorkbookModel
Web → Protocol
Web → FormulaEngine
Web → OOXML
Web → CollabSocket
Web → Persistence
Web → Structural Transform
```

---

# 十二、当前仍缺失的 SDK Domain

逐项重新确认当前 HEAD。

下面这些领域如果仍然不存在 SDK 公共契约，全部完成。

## 12.1 Selection

建立：

```ts
workbook.selection
```

包括：

```text
current selection
active cell
multi-range
row selection
column selection
select all
select address
move
extend
gesture state
```

Web 不直接持有 SelectionService。

---

## 12.2 Editing

建立：

```ts
workbook.editing
```

承载：

```text
begin
draft
commit
cancel
formula reference mode
IME
caret
autocomplete
editor overlay
fixed decimal
enter direction
```

Web 只渲染 editing projection。

---

## 12.3 Clipboard

建立：

```ts
workbook.clipboard
```

承载：

```text
copy
cut
paste
paste special
system clipboard
internal clipboard
transpose
formats
values
formulas
links
```

---

## 12.4 Formula

建立：

```ts
workbook.formulas
```

承载：

```text
calculation mode
recalculate
precedents
dependents
formula audit
evaluate step
defined names integration
external links
function catalog
```

Cell.setFormula 不能代替完整 Formula Domain。

---

## 12.5 Tables

建立：

```ts
worksheet.tables
```

承载：

```text
create
remove
resize
rename
style
filter
total row
header row
convert to range
column operations
structured references
```

---

## 12.6 Validation

建立：

```ts
worksheet.validation
```

---

## 12.7 Conditional Formatting

建立：

```ts
worksheet.conditionalFormatting
```

---

## 12.8 Drawing

建立：

```ts
worksheet.drawings
```

统一：

```text
shape
connector
image
text box
form control
barcode
camera
group
z-order
align
distribute
```

---

## 12.9 Chart

建立：

```ts
worksheet.charts
```

UI 不允许再构造 Chart command。

---

## 12.10 Pivot

建立：

```ts
worksheet.pivots
```

统一：

```text
pivot
pivot chart
slicer
timeline
field catalog
layout
filter
aggregate
calculated field
drill down
refresh
```

---

## 12.11 Sparkline

建立：

```ts
worksheet.sparklines
```

---

## 12.12 Review

建立：

```ts
workbook.review
```

统一：

```text
comment
note
hyperlink
review threads
```

---

## 12.13 Permission

建立：

```ts
workbook.permissions
```

公开 effective capability。

必须结合：

```text
WorkbookRole
Range Access
Protection
Object Protection
```

目录层：

```text
canEdit
canShare
```

不是 opened workbook 的完整 effective permission。

---

## 12.14 Range Access

建立：

```ts
workbook.permissions.rangeAccess
```

Web 不允许直接调用：

```text
createRangeAccessRegion
updateRangeAccessRegion
deleteRangeAccessRegion
```

Session 方法。

---

## 12.15 Collaboration

建立：

```ts
workbook.collaboration
```

公开：

```text
connection state
presence
peer list
cursor
resync
revision
pending operations
```

Web 不知道 Socket。

---

## 12.16 History

当前只有：

```text
undo
redo
```

补全：

```ts
workbook.history
```

包括：

```text
entries
undo
redo
revision list
preview
restore
replay status
invalid history status
```

---

## 12.17 Document

建立：

```ts
workbook.document
```

包括：

```text
save
saveAs
export
compatibility
source format
native capability
source artifact identity
```

---

## 12.18 Print

建立：

```ts
workbook.print
```

包括：

```text
page setup
print area
titles
page break
preview
PDF/export
```

---

## 12.19 Viewport / Render / Interaction

建立：

```ts
workbook.viewport
workbook.render
```

SDK ownership：

```text
coordinate mapping
viewport state
freeze geometry
selection hit test
pointer gesture semantics
keyboard gesture semantics
range drag
fill handle
dimension drag
drawing hit-test
drawing interaction
render scheduling
canvas projection
floating object projection
```

Web 保留：

```text
canvas DOM element
PointerEvent forwarding
KeyboardEvent forwarding
ResizeObserver
browser file picker
browser download
```

---

# 十三、重点整改 Web 大文件

重新审查：

```text
apps/web/src/components/canvas/drawing-renderers.ts
apps/web/src/components/SheetCanvas.tsx
apps/web/src/components/canvas/useCanvasInteraction.ts
apps/web/src/editor/command-controller.ts
apps/web/src/editor/EditorShell.tsx
apps/web/src/editor/RibbonHost.tsx
apps/web/src/editor/FeaturePanelHost.tsx
apps/web/src/containers/WorkbookHubContainer.tsx
```

目标不是单纯缩短文件。

必须逐段判定：

```text
这是 presentation 还是 spreadsheet domain？
```

凡是涉及：

```text
canonical model
command
selection semantics
formula
permission
structural transform
chart/pivot calculation
drawing hit test
business transaction
```

全部进入 SDK。

---

# 十四、Workbook Hub 继续收口

虽然 WorkbooksDomain 已迁 SDK，但 Web Hub 仍然知道部分：

```text
protocol DTO
OOXML utility
storage vocabulary
role interpretation
```

最终 Workbook Hub 只消费 SDK DTO：

```ts
CatalogEntry
Space
Folder
Member
UserPreferences
ImportResult
ExportResult
```

不得直接 import protocol。

不得直接调用 OOXML utility。

---

# 十五、SDK 必须拥有自己的 Public DTO

当前 SDK 还存在：

```ts
interface CatalogEntry extends WorkbookCatalogEntry
```

这种设计。

这是内部 DTO 泄漏。

最终：

```text
SDK Public API
```

不能依赖消费者知道：

```text
WorkbookCatalogEntry
WorkbookResolution
WorkspaceRecord
WorkspaceRecordMetadata
WorkbookCatalogRemoteClient
```

来自内部 package。

定义 SDK 自己稳定的：

```text
CatalogEntry
WorkbookOpenState
WorkbookAccess
Space
Folder
Member
Compatibility
DocumentCapability
```

内部完成转换。

---

# 十六、WorkbooksActions 公共面必须显式定义

当前类似：

```ts
Omit<WorkbooksDomain, 'retire'>
```

这种写法禁止继续使用。

它会让 internal method 自动变成 public API。

必须定义明确 interface：

```ts
interface WorkbooksActions {
  list(...)
  create(...)
  open(...)
  import(...)
  export(...)
  rename(...)
  copy(...)
  move(...)
  trash(...)
  restore(...)
  purge(...)
  share(...)
}
```

禁止暴露内部：

```text
resolve
markOpened
retire
requireRemote
```

---

# 十七、统一 SDK Error Contract

所有 SDK 公共 API 只允许抛：

```ts
SdkError
```

当前需要消灭公共边界的：

```text
WorkbookCatalogError
NativeDocumentError
ApiRequestError
CommandDispatchError
WorkspaceStorageError
```

这些可以作为：

```text
SdkError.cause
```

保留。

消费者只处理：

```ts
SdkError {
  code
  operation
  message
  recovery
  status
  object
}
```

---

# 十八、Guest Share 收进 Auth Domain

当前 Web 如果还在调用：

```text
resolveShareToken()
```

立即整改。

Guest Share capability 必须由 SDK 捕获。

流程：

```text
SDK 初始化
    ↓
读取 route share capability
    ↓
保存到 SDK auth/access context
    ↓
从 URL 中清除 capability
    ↓
Protocol/WS 只读取 SDK 内部 capability source
```

Web 只能看到：

```text
authenticated
guest
anonymous
access denied
```

状态。

Web 不能读原始 share token。

---

# 十九、Identity 领域重新命名和划分

当前如果：

```ts
sdk.identity
```

主要是：

```text
listUsers
createUser
setUserEnabled
resetPassword
```

则领域命名错误。

分成：

```text
sdk.auth
sdk.identity
sdk.users
```

语义：

```text
auth
  authentication lifecycle

identity
  current verified identity

users
  administrative user management
```

Current identity 不应隐藏在 `auth.snapshot.context` 里作为唯一入口。

---

# 二十、React binding 与 SDK Core 分离

当前 `@react-sheets/sdk` 如果仍直接依赖 React，进行拆分。

目标：

```text
@react-sheets/sdk
@react-sheets/sdk/react
```

Core：

```text
framework-neutral
```

React binding：

```text
hooks
providers
useWorkbook
useSdk
```

不得让程序化 Node/Worker consumer 因 SDK Core 强依赖 React。

---

# 二十一、Structural Correctness 仍是 P0

SDK 收口不能掩盖底层语义缺口。

下面这些必须继续完成。

---

## 21.1 Cross-sheet Cut/Paste

如果仍存在：

```text
cross-sheet cut/paste requires a canonical structural move patch
```

完成真正的跨工作表 move transaction。

不能用：

```text
read source
clear source
write destination
```

模拟。

必须单事务覆盖：

```text
cells
formulas
defined names
tables
validation
conditional formatting
charts
pivots
data regions
drawing anchors
history
collaboration
```

---

## 21.2 `Range.moveTo`

SDK `Range.moveTo` 不得永久限制：

```text
same worksheet
```

同一 workbook 跨 worksheet move 属于在线 Excel 基础能力。

它必须进入 canonical structural transaction。

---

## 21.3 Cross-sheet Structural Formula Transform

消除核心场景中的：

```text
cannot rewrite a cross-worksheet range
cannot rewrite a partially qualified range
```

建立统一 Reference Ownership Graph。

---

## 21.4 Table Structural Transaction

消除：

```text
inserting a worksheet column inside Sheet Table
requires a table-column structural patch
```

以及：

```text
structural edit intersects workbook table
and requires a table transaction
```

Table 必须成为 Structural Transaction 正式 participant。

---

## 21.5 History Rebase

逐项搜索：

```text
historyRebase
kind: 'invalidate'
```

核心操作不能继续大面积 invalidate：

```text
range move
cell shift
row permutation
sheet identity
sheet table
```

远程 structural mutation 必须 transform 本地：

```text
undo
redo
pending operation
selection
range ownership
```

---

## 21.6 Data Validation Server Calculation

消灭：

```text
formula entry under data validation
requires the shared calculation evaluator
```

客户端与服务端共享 canonical calculation semantics。

---

# 二十二、Data Domain 补全

当前 SDK DataActions 如果仍只有：

```text
quickSort
toggleFilter
clearFilter
textToColumns
removeDuplicates
subtotal
```

继续拆出完整 Data API。

而且 `subtotal` 不能永久限制：

```text
SUM
COUNT
AVERAGE
```

必须与产品定义的 Excel Subtotal 语义统一。

---

# 二十三、Sort / Outline / Hidden Rows

如果仍存在：

```text
sorting a multi-row outline group together
with separately hidden rows is unsupported
```

统一：

```text
manual hidden
filter hidden
outline collapsed
sort permutation
```

成为 canonical row visibility + row permutation semantics。

---

# 二十四、Boundary Gate 必须重写

当前检查器如果仍然出现：

```text
UI must dispatch CommandDescriptor directly
```

这一规则已经错误。

新的规则必须是：

> UI 不允许知道 CommandDescriptor。

`apps/web/**` 非测试文件出现以下内容立即失败：

```text
@react-sheets/core-model
@react-sheets/command-runtime
@react-sheets/protocol
@react-sheets/formula-engine
@react-sheets/render-engine
@react-sheets/sheet-features
@react-sheets/exchange-excel-ooxml
@react-sheets/spreadsheet-app
WorkbookSession
WorkbookResolution
CommandDescriptor
MutationInfo
WorkbookApiClient
CollabSocketClient
FormulaEngine
StructuralPatch
```

业务网络请求也必须禁止：

```text
fetch('/api/
fetch("/api/
```

---

# 二十五、目标 SDK 顶层结构

最终接近：

```ts
interface SpreadsheetSdk {
  readonly auth: AuthService;
  readonly identity: IdentityService;
  readonly users: UserAdministrationService;
  readonly workbooks: WorkbooksService;
  readonly spaces: SpacesService;

  dispose(): Promise<void>;
}
```

---

# 二十六、目标 Workbook

```ts
interface Workbook {
  readonly id: string;

  readonly view: WorkbookView;

  readonly worksheets: WorksheetCollection;
  readonly names: DefinedNameCollection;

  readonly selection: SelectionService;
  readonly editing: EditingService;
  readonly clipboard: ClipboardService;
  readonly formulas: FormulaService;

  readonly data: DataService;
  readonly tables: TableService;
  readonly validation: ValidationService;
  readonly conditionalFormatting: ConditionalFormattingService;

  readonly drawings: DrawingService;
  readonly charts: ChartService;
  readonly pivots: PivotService;
  readonly sparklines: SparklineService;

  readonly review: ReviewService;
  readonly permissions: PermissionService;

  readonly collaboration: CollaborationService;
  readonly history: HistoryService;

  readonly document: DocumentService;
  readonly print: PrintService;

  readonly viewport: ViewportService;

  save(): Promise<void>;
  flush(): Promise<void>;
  close(): void;
}
```

具体类名可以按当前代码语义调整。

领域 ownership 不能退化。

---

# 二十七、目标目录

最终结构接近：

```text
frontend-react/
  apps/
    web/
      src/
        app/
        routes/
        components/
        hosts/

  packages/
    sdk/
      src/
        core/
        auth/
        identity/
        users/
        workbooks/
        spaces/

        workbook/
          lifecycle/
          view/
          cells/
          ranges/
          worksheets/
          selection/
          editing/
          clipboard/
          formulas/
          names/
          tables/
          data/
          validation/
          conditional-formatting/
          drawings/
          charts/
          pivots/
          sparklines/
          review/
          permissions/
          collaboration/
          history/
          document/
          print/
          viewport/

        internal/

    core-model/
    command-runtime/
    formula-engine/
    protocol/
    sheet-features/
    render-engine/
    exchange-excel-ooxml/
    number-format/
    ui-system/
```

不得新增 compatibility bridge。

---

# 二十八、删除目标

完成迁移后删除：

```text
WorkbookSession 公共入口
useWorkbookSession
Web CommandDescriptor 构造链
Web WorkbookResolution 流程
Web resolveShareToken
Web protocol imports
Web core-model imports
Web spreadsheet-app imports
Web render-engine imports
Web OOXML imports
旧 Auth Web 实现
旧 Dimensions Web 实现
重复 Workbook role 类型
公开 WorkbookCatalogError
```

如果 `spreadsheet-app` 在职责全部迁出后已经没有独立领域价值，直接删除这个 package。

不要保留空壳。

---

# 二十九、SDK 测试要求

已有：

```text
test:sdk
test:sdk-uat
```

继续扩展，使每个正式 Domain 都有：

```text
unit
contract
integration
UAT
```

SDK UAT 必须从真正：

```ts
createSpreadsheetSdk()
```

入口执行。

不得直接拿内部 Session 做 SDK 验收。

---

# 三十、Web 薄壳验收

最终静态检查必须证明：

```text
apps/web
```

业务源码对内部 spreadsheet package import 数：

```text
0
```

Web 源码中：

```text
WorkbookSession = 0
CommandDescriptor = 0
WorkbookResolution = 0
MutationInfo = 0
```

Web 只消费：

```text
SDK objects
SDK snapshots
SDK actions
UI system
browser host primitives
```

---

# 三十一、SDK 完整性验收矩阵

最终输出下面矩阵，并确保不是只写“存在”。

| Domain | Public API | Canonical owner | Command chain | Permission | History | Collaboration | Persistence | Server authority | Web migrated |
|---|---|---|---|---|---|---|---|---|---|

覆盖：

```text
Auth
Identity
Users
Workbook Center
Workbook
Worksheet
Cell
Range
Selection
Editing
Clipboard
Formula
Defined Name
Table
Data
Validation
Conditional Formatting
Drawing
Chart
Pivot
Sparkline
Review
Range Access
Permission
Collaboration
History
Document
Print
Viewport
Render Interaction
```

---

# 三十二、工作方式

先从：

```text
dependency graph
domain ownership
public contract
write chain
view chain
```

判断结构。

逻辑完整后再改代码。

不要因为某一个测试失败就逐点打补丁。

如果当前实现与目标 ownership 冲突，直接重构整个链。

---

# 三十三、禁止事项

禁止：

```text
Bridge
Adapter
Shim
旧 API wrapper
双写
双状态
双 Command 路径
第二套 Workbook Model
第二套 Permission truth
第二套 Formula truth
Web 继续持有 Session
```

---

# 三十四、最终交付内容

完成代码整改后输出：

## A. 当前 HEAD

```text
SHA
Commit
```

## B. 删除的旧链路

列出文件和公共入口。

## C. SDK Public API

按领域列出。

## D. Web Dependency

输出最终 Web 业务依赖。

## E. WorkbookSession 状态

明确：

```text
deleted
```

或者列出仍存在的真实阻塞。

不能模糊描述。

## F. Structural 缺口状态

逐项说明：

```text
cross-sheet move
cross-sheet reference transform
table structural edit
history rebase
server DV calculation
```

## G. 测试结果

列出：

```text
typecheck
boundary
contracts
sdk unit
sdk UAT
frontend unit
backend tests
```

如果本地环境无法执行某项，明确记录为环境限制，不能把整个任务判失败。

---

# 三十五、最终架构判定标准

只有同时满足下面条件，任务才算完成：

```text
Web → SDK 是唯一业务依赖方向

Web 不知道 WorkbookSession

Web 不知道 CommandDescriptor

SDK 拥有完整 Interactive Editor Domain

SDK 拥有完整 Automation Object API

Auth / Identity / Credentials 归 SDK

WorkbookRole 唯一

Workbook Center 归 SDK

Workbook Runtime 领域化

Reactive View 归 SDK

Render / Interaction 业务语义归 SDK

所有业务写统一走 Command → Mutation → Model

SDK 公共边界只抛 SdkError

SDK Public DTO 不泄漏 internal package

Boundary Gate 自动阻断 Web 越层

核心 structural 场景不再依赖 UNSUPPORTED_FEATURE

协同 History 不再依赖大面积 invalidate
```

本轮不是继续给旧 `WorkbookSession` 加能力。

本轮目标是：

> **完成从“拥有 SDK 包”到“整个在线 Excel 产品由 SDK 驱动”的架构切换。**