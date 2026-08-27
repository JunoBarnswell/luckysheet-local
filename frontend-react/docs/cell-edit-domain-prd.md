# CellEditDomain 根治 PRD（Issue #301）

## 1. 文档状态

- 目标 Issue：`JunoBarnswell/luckysheet-local#301`
- 基线：`github/main@aa1536b1e4367a3ca86b1cc0a8dae2d3236410de`
- 交付分支：`codex/issue-301-cell-edit-domain`
- 优先级：P0
- 交付策略：clean-break；不保留旧编辑链、兼容 Bridge、双写或 UI fallback
- 目标产品：Excel / SpreadJS 级单元格编辑域，并把编辑热路径从完整工作簿快照中物理隔离

## 2. 产品问题

当前产品把“正在编辑”压缩为 `editingCell != null`，同时由多个 Surface 自行解释输入：

```text
Canvas keydown / pointer
CellEditor textarea keydown
FormulaBar input keydown
EditorShell callbacks
WorkbookSession begin/commit/cancel
EditSession referenceMode
useCanvasInteraction.editingActiveRef
```

这同时造成语义和性能问题：

1. Enter、Tab、Escape、F2、F4、方向键、IME 和 pointer 在不同 Surface 上有不同所有者。
2. `editingActiveRef` 与 `EditSession` 并存，React render state 与 imperative ref 可以发生分叉。
3. Point Mode 只是 `referenceMode: boolean`，没有公式引用选择、颜色、跨表、移动/缩放和 popup 优先级的完整状态。
4. Formula Bar 与行内编辑器虽然显示同一个字符串，但仍分别决定 commit/cancel。
5. `setFormulaDraft`、caret 和 composition 的每次变化都调用 `WorkbookSession.emit()`；React 订阅完整 `UiSnapshot`，一次按键会进入整个 `EditorShell` 和 `SheetCanvas` 更新链。
6. validation warning/information 只以异常字符串暴露，没有正式 confirmation state。
7. Cell editor kind 是固定联合类型并由 UI 条件分支扩张，没有可执行的 editor contract。
8. rich text 已能在 `CellData`/OOXML 中保存，但没有字符级编辑域和原子提交命令。

## 3. 官方行为基线

Microsoft Excel 官方合同：

- Double click 在命中文字位置进入 Edit；F2 在内容末尾进入 Edit；Formula Bar 在点击位置共享编辑内容。
- Edit 模式方向键移动文本 caret；Enter/Tab 提交并移动；Alt+Enter 插入单元格内换行；Esc 取消。
- 状态栏公开 Ready / Enter / Edit / Point，并在 Edit 下支持 Overtype。
- Point 是公式引用选择模式，不等价于普通单元格选择。

MESCIUS SpreadJS 官方合同：

- Ready / Enter / Edit 是公开 editor status，状态改变有正式生命周期。
- editor 能力、Formula text box、cell type/editor status 属于工作表编辑模型，不应由单个 DOM 输入框决定。

官方来源：

- <https://support.microsoft.com/en-us/excel/edit-cell-contents>
- <https://support.microsoft.com/en-us/excel/excel-status-bar-options>
- <https://support.microsoft.com/en-us/excel/use-formula-autocomplete>
- <https://developer.mescius.com/spreadjs/docs/features/cells/editing>
- <https://developer.mescius.com/spreadjs/docs/features/worksheet/editor-status>

## 4. 目标与非目标

### 4.1 目标

1. `CellEditDomain` 是编辑业务状态和状态迁移的唯一 owner。
2. Keyboard / Pointer / IME / Formula Bar / cell control 只发送 typed intent。
3. Ready / Enter / Edit / Point / Overtype、draft、caret、composition、reference selection、validation pending 和 surface focus 有一份 canonical state。
4. draft 期间零 model mutation；一次成功 commit 对应一次 semantic history transaction。
5. `text-input.ts` / `sheet.cell.commitText` 继续是普通文本词法解释的唯一 owner。
6. Formula、rich text 和结构化 editor 使用明确 typed commit command，不伪装为字符串。
7. 编辑热路径只更新轻量 edit projection，不重建完整 `UiSnapshot`，不触发 Canvas 基础层重绘。
8. 保护、权限、formulaHidden、spill child、merged target、Pivot、drawing 和 validation 由统一进入门/提交门 fail-close。

### 4.2 非目标

- 不把 edit draft 写入持久化 snapshot、OOXML、历史或 collaboration mutation。
- 不在 React 中新增日期、货币、百分比或公式 parser。
- 不保留旧 `EditSession` API 的别名或转发器。
- 不以延迟 commit、debounce 或隐藏 UI 卡顿代替架构修复。
- 不让 worker 接管逐键 caret/draft reducer；逐键 reducer 必须同步且 O(编辑文本长度或更低)，后台只处理可取消的候选索引/昂贵分析。

## 5. 状态所有权

### 5.1 持久化业务状态

Owner：`core-model`

- `CellData.value/formula/style/numberFormat`
- `CellData.richText`
- workbook-owned editor configuration
- validation rules、protection、merged cells、spill ranges

这些状态只在成功 semantic command 后变化。

### 5.2 编辑域临时业务状态

Owner：`CellEditDomain`（`spreadsheet-app`）

```ts
type CellEditorStatus = 'ready' | 'enter' | 'edit' | 'point';
type CellEditSurface = 'grid' | 'formula-bar' | 'formula-panel';

interface CellEditSession {
  target: CanonicalCellTarget;
  source: CellEditSource;
  status: Exclude<CellEditorStatus, 'ready'>;
  surface: CellEditSurface;
  editorKind: CellEditorKind;
  draft: CellDraft;
  caret: { start: number; end: number };
  composition: { active: boolean; text: string };
  overtype: boolean;
  referenceSelection: FormulaReferenceSelection | null;
  originalSelection: SelectionSnapshot;
  originalCell: CellData | null;
  validation: CellEditValidationState;
  baseCellFingerprint: string;
}
```

`CellEditDomain` 不拥有 workbook model、selection model 或 Canvas geometry。它保存进入编辑时的稳定快照，并通过 transition effects 请求这些 owner 执行动作。

### 5.3 Selection

Owner：`SelectionService`

- 普通 Ready 选择仍由 `SelectionService` 管理。
- Point Mode 的“公式引用选择”属于 `CellEditDomain.referenceSelection`，不能覆盖普通 selection。
- Point Mode 可把 reference range 投影给 Canvas，但不得把公式目标格的 original selection 丢失。

### 5.4 React/UI 状态

- DOM focus、popup measurement、textarea ref 是 Surface 局部 UI 状态。
- popup open/active item 若影响键盘优先级，则属于 `CellEditDomain` 的 canonical overlay state。
- React 不保存 draft/caret/referenceMode 的副本。

## 6. 单一数据流

```text
DOM Keyboard / Pointer / IME / FormulaBar / Cell Control
                         │
                         ▼
              Canonical CellEditIntent
                         │
                         ▼
                 CellEditDomain.dispatch
                 state transition + effects
                         │
          ┌──────────────┼───────────────┐
          ▼              ▼               ▼
  lightweight edit   selection/nav   semantic commit
     projection         effects           effect
          │                              │
          ▼                              ▼
 FormulaBar / overlay / status     CommandRuntime
                                  │
                                  ▼
                  sheet.cell.commitText / typed command
                                  │
                                  ▼
              model → history → calc → collaboration
                      → persistence → OOXML
```

普通 raw text 的提交链固定为：

```text
CellEditorBehavior.toCommitPayload(raw text)
  -> CellInputInterpretationContext
  -> sheet.cell.commitText
  -> interpretCellInput
  -> validation/write authority
  -> canonical mutation
```

## 7. Typed Intent 合同

所有 Surface 只能发送以下语义 intent；DOM event 不进入 domain：

```text
begin.directTyping / begin.doubleClick / begin.f2 / begin.formulaBar
text.insert / text.replace / text.deleteBackward / text.deleteForward / text.newLine
caret.move / caret.select / overtype.toggle
composition.start / composition.update / composition.end
status.toggle
reference.begin / reference.point / reference.extend / reference.move / reference.resize
reference.toggleAbsolute / reference.switchSheet
autocomplete.open / autocomplete.move / autocomplete.accept / autocomplete.close
validation.confirm / validation.reject
commit / commitAndMove / commitToSelection / cancel
pointer.cell / pointer.referenceHandle / surface.focus
```

`dispatch` 返回 typed outcome：

```ts
interface CellEditDispatchResult {
  handled: boolean;
  status: CellEditorStatus;
  effects: readonly CellEditEffect[];
}
```

UI 只依据 `handled` 决定是否继续交给普通 worksheet shortcut/pointer owner，不解释编辑状态。

## 8. 状态迁移与优先级

### 8.1 核心迁移

| 当前状态 | Intent | 结果 |
|---|---|---|
| Ready | printable | Enter，draft 从该字符开始 |
| Ready | F2 | Edit，caret=end |
| Ready | double click | Edit，caret=文字 hit position |
| Ready | Formula Bar focus | Edit，同一 active cell |
| Enter | printable | 更新 draft，保持 Enter |
| Enter | click editor | Edit |
| Enter/Edit | Alt+Enter | 插入 `\n`，不提交 |
| Enter/Edit | Ctrl+Enter | selection 原子提交，Ready |
| Edit | Arrow/Shift+Arrow | caret/text selection |
| Edit formula | F2 | Point |
| Point | Arrow/Shift+Arrow | 移动/扩展引用 |
| Point | pointer cell/range | 插入或替换引用，不提交 |
| Point | F4 | caret 所在引用绝对/混合循环 |
| Point | sheet switch | 保持 session，改变 reference sheet |
| Any edit | Enter/Tab | 原子提交并按方向移动 |
| Any edit | Esc | popup 优先关闭，否则 cancel |
| Any edit | validation warning/info | confirmation-required，模型不变 |
| composition active | keydown | IME owner，worksheet/domain shortcut 不执行 |

### 8.2 Popup 优先级

```text
IME composition
  > validation confirmation
  > editor-specific popup/list
  > formula autocomplete/function hint
  > CellEditDomain state transition
  > worksheet shortcut/navigation
```

## 9. CellEditorRegistry

Registry 是 editor kind 的唯一运行时分派点：

```ts
interface CellEditorBehavior<TDraft extends CellDraft = CellDraft> {
  readonly kind: CellEditorKind;
  canEnter(context: CellEditorContext): CellEditorEntryDecision;
  createDraft(context: CellEditorContext): TDraft;
  reduce(intent: CellEditIntent, draft: TDraft, context: CellEditorContext): TDraft;
  validate(draft: TDraft, context: CellEditorContext): CellEditValidationResult;
  toCommitPayload(draft: TDraft, context: CellEditorContext): CellEditCommitPayload;
  hitTestControl?(point: Point, context: CellEditorContext): CellControlHit | null;
}
```

首批正式 adapter：

- Text
- Number
- DateTime
- ValidationList
- ComboBox
- Checkbox
- Mask
- Formula
- RichText

禁止在 `SheetCanvas`/React 中按 kind 扩张 `if/else`。adapter 不重新实现 `text-input.ts` 的词法解析。

## 10. Formula 子域

Formula 编辑逻辑是 `CellEditDomain` 使用的无状态/增量计算服务，不拥有第二 session：

- tolerant tokenization：不完整公式可以继续编辑。
- strict parse：只在 commit 或明确校验时执行。
- caret token lookup 和 F4 AST rewrite。
- same-sheet/cross-sheet/range/defined-name/structured-reference。
- token ↔ range 稳定颜色映射。
- Point range create/move/resize。
- Formula AutoComplete 基于 caret 前 token，不破坏 caret 后文本。
- function、defined name、table、structured item 和 argument hint 统一候选模型。

候选索引按 workbook metadata revision 增量构建；逐键查询不得扫描 cells。昂贵索引构建可取消，并以 revision token 丢弃 stale result。

## 11. Validation、权限与拒绝路径

### 11.1 统一进入门

进入编辑前一次性解析：

- canonical merged target
- share role / permission
- sheet protection + locked/formulaHidden
- dynamic array source/spill child
- Pivot projection / drawing / hyperlink / cell control hit
- adapter capability (`allowEditInCell` 等)

拒绝必须返回 typed `CellEditError`，包含：

```text
code / message / sheetId / row / column / recovery
```

### 11.2 统一提交门

- Stop：保持 session、draft、caret，进入 blocking-error。
- Warning/Information：进入 confirmation-required；confirm 后使用同一 draft 重放一次提交，reject 返回编辑。
- remote target conflict：若 target canonical fingerprint 已变化，返回 `CELL_EDIT_REVISION_CONFLICT`，不得覆盖远端值。
- Ctrl+Enter：先对全部目标完成权限、保护、spill、validation 和 revision preflight，再提交一个 mutation descriptor；任何失败均为零写入。

## 12. Rich text 合同

现有 `CellData.richText` 是持久化真相，`value` 是公式/搜索使用的 plain-text projection。新增编辑能力必须满足：

- draft 以 run + plain text + character selection 表达，不把 run 信息压平成字符串。
- 字符级 bold/italic/underline/color/superscript/subscript 改变 run，不修改无关字符。
- commit 使用 typed rich-text command，同时原子更新 `richText` 与 plain `value`。
- OOXML 已有的 shared string/inline string rich text 保真链继续使用 canonical runs。
- Formula cell 与 rich text 互斥时 fail-close，不静默丢 run。

## 13. UI 与组件合同

### 13.1 轻量订阅边界

`WorkbookSession.getUiSnapshot()` 不再包含逐键变化的 `editSession`、draft、caret、composition 或 Point selection。

新增稳定外部 store 合同：

```ts
session.cellEdit.subscribe(listener)
session.cellEdit.getSnapshot()
session.cellEdit.dispatch(intent)
```

独立订阅者：

1. Formula Bar edit surface：只订阅 draft/caret/surface/popup。
2. Cell editor overlay：只订阅 target/draft/caret/editor kind/validation。
3. Cell mode indicator：只订阅 status/overtype。
4. Formula reference overlay：只订阅 reference projection，并仅 invalidates overlay layer。

`EditorShell` 和 `SheetCanvas` 基础层不得因为每个字符或 caret 移动而重渲染。

### 13.2 Surface 组件职责

- `CellEditor`：受控渲染、DOM selection 同步、IME 事件归一、发送 intent；无 Enter/Tab/Escape/F4 业务分支。
- `FormulaBar`：显示共享 draft、发送 focus/input/key intent；无独立 commit/cancel 决策。
- `SheetCanvas`：提供 geometry owner/overlay host 和 worksheet hit；不保存 editing boolean。
- `StatusBar`：显示 Ready/Enter/Edit/Point/Overtype 投影。

### 13.3 Geometry

唯一几何来源：render skeleton + `PaneMap`。

- canonical/merged rect。
- long text 向可用空白邻格扩张；不能固定 `overflow-hidden` 在单格宽度。
- multiline/wrap 保证 caret 可见。
- popup 可在 viewport 边缘翻转。
- scroll/zoom/freeze 只重新计算 overlay geometry，不改变 session。
- hit、selection、editor 与 commit address 完全一致。

## 14. 事务、历史、协作与持久化

```text
begin
  -> N 次 draft/caret/composition/reference transition
  -> 0 model mutation / 0 history / 0 persistence write
commit
  -> 1 semantic command
  -> 1 history transaction
  -> calc once per committed transaction
  -> collaboration broadcasts committed mutation only
  -> persistence checkpoints canonical model only
```

- Esc、validation reject、popup close：0 history。
- Ctrl+Enter：无论多少 selection cells，1 history transaction。
- presence 可发布“actor 正在编辑 cell/status”，但默认不广播每个 draft 字符；draft preview 若启用必须节流且不进入 revision。
- snapshot/protocol 不保存临时 CellEditSession。
- OOXML 只消费提交后的 canonical cell。

## 15. 性能硬边界

### 15.1 算法边界

- draft/caret/composition/overtype transition 与 workbook cell count 无关。
- 每个按键不得调用 `WorkbookSession.emit()`、`getUiSnapshot()`、model snapshot、Canvas base redraw、calc、history、persistence 或 collaboration mutation。
- Point pointer move 不扫描 workbook；range geometry 由 `PaneMap`/skeleton 直接定位。
- autocomplete 查询使用 Map/trie/prefix index；不得逐键扫描 worksheets/cells。
- React 订阅细分到 primitive/小投影；稳定 handler/controller 不随 snapshot 重建。

### 15.2 可测预算

在固定测试机和 production build 上记录原始结果，目标：

| 场景 | 预算 |
|---|---|
| 1–1,000 字符 draft reducer | p95 ≤ 1 ms，p99 ≤ 2 ms |
| 32 KiB 长文本插入/删除 | p95 ≤ 4 ms，p99 ≤ 8 ms |
| input event → editor/caret paint | p95 ≤ 16.7 ms，p99 ≤ 33.4 ms |
| IME composition update → paint | p95 ≤ 16.7 ms |
| Point pointer move/extend | p95 ≤ 4 ms，连续拖动保持 60 fps |
| 已构建索引后的 autocomplete query | p95 ≤ 4 ms，p99 ≤ 8 ms |
| 100k×20 workbook 单格 commit | p95 ≤ 20 ms（不含远端网络确认） |
| 100k×20 workbook 编辑热路径 | 与空 workbook 相比 p95 增幅 ≤ 20% |

内存边界：

- active session 为 `O(draft + richTextRuns + references + popupCandidates)`。
- autocomplete index 只索引 functions/defined names/tables/fields 等 metadata，不复制 cell matrix。
- session 结束后 draft、candidate、reference geometry 和 adapter-local state 均可回收。
- benchmark 同时记录 heap before/peak/after；连续 1,000 次 begin/cancel 后 retained heap 不线性增长。

## 16. Clean-break 删除清单

同一 PR 中物理删除：

1. `useCanvasInteraction.editingActiveRef`。
2. `CellEditor` 对 Enter/Tab/Escape/F4 的业务判断。
3. `FormulaBar` 对 Enter/Escape 的独立 commit/cancel 判断。
4. Canvas pointer down 的通用“先 commit 再 hit-test”分支。
5. `onInsertRef` 断链 prop。
6. `referenceMode: boolean` 旧模型。
7. `core-model/domain.ts` 中未使用的持久化样式 `EditSession` 类型。
8. `spreadsheet-app/types.ts` 中并行 `EditSession` DTO 及 `UiSnapshot.editSession` 热路径。
9. `WorkbookSession` 的 begin/setDraft/setCaret/composition/commit/cancel 分裂入口；消费者迁移到唯一 dispatch。
10. StatusBar 固定“就绪”。
11. React 按 `CellEditorKind` 扩张的运行时编辑分派。

禁止保留 deprecated alias、forwarder、fallback reader 或双写过渡期。

## 17. 实施 TODO

### A. 合同与 domain

- [ ] A1. 新建 typed `cell-edit` contract：status、session、draft、intent、effect、error、projection、lifecycle。
- [ ] A2. 实现纯状态迁移表，覆盖 Ready/Enter/Edit/Point/Overtype/IME/popup priority。
- [ ] A3. 实现 `CellEditDomain` external store；snapshot identity 仅在编辑投影变化时更新。
- [ ] A4. 实现 canonical entry/commit gate 和 typed fail-close error。
- [ ] A5. 实现 target fingerprint conflict detection。
- [ ] A6. 实现生命周期事件，保证 begin/change/cancel 无 model mutation。

### B. Adapter 与提交

- [ ] B1. 建立 `CellEditorRegistry`，注册 Text/Number/DateTime/List/ComboBox/Checkbox/Mask/Formula/RichText behavior。
- [ ] B2. Text/Number/DateTime 最终调用唯一 `sheet.cell.commitText`。
- [ ] B3. Formula adapter 使用 tolerant editing + strict commit parse。
- [ ] B4. 新建 Ctrl+Enter 原子 range commit command，完成全目标 preflight 和单 history transaction。
- [ ] B5. 新建 rich-text typed command，原子维护 runs/plain value。
- [ ] B6. validation confirmation 进入 domain 状态并支持 confirm/reject。

### C. Formula editing

- [ ] C1. caret token lookup、F4 rewrite 和 incomplete formula 合同。
- [ ] C2. Point same-sheet click/drag range。
- [ ] C3. cross-sheet Point session 保持与 sheet switch。
- [ ] C4. reference token color、border、move、resize。
- [ ] C5. Formula AutoComplete functions/names/tables/structured refs/argument hints。
- [ ] C6. 候选索引 revision/cancellation/stale-result fail-close。

### D. UI clean-break

- [ ] D1. 新建轻量 edit store hook 和独立 Formula Bar surface。
- [ ] D2. 将 editor overlay 从 `SheetCanvas` 大组件的逐键 props 中拆出。
- [ ] D3. Canvas interaction 全部改发 intent，删除 `editingActiveRef`。
- [ ] D4. Formula Bar/CellEditor 只做 DOM 事件归一和渲染。
- [ ] D5. StatusBar 投影 Ready/Enter/Edit/Point/Overtype。
- [ ] D6. reference overlay 只 invalidates overlay layer。
- [ ] D7. geometry 支持 merge/freeze/scroll/zoom/long text/multiline/popup flip。

### E. 边界与持久语义

- [ ] E1. 统一 protected/locked/formulaHidden/spill/merged/Pivot/drawing/control 进入门。
- [ ] E2. 统一 direct input 与 validation list/control 的 ownership。
- [ ] E3. collaboration 只广播 committed mutation；presence 不进入 revision。
- [ ] E4. snapshot/protocol/OOXML 确认不持久化 edit session；rich text round-trip 保真。
- [ ] E5. 删除所有旧类型、旧方法、旧 props、旧测试假设和 parallel state。

### F. 验证与性能

- [ ] F1. Domain 状态迁移成功/拒绝路径单元测试。
- [ ] F2. adapter、validation、revision conflict、Ctrl+Enter/rich-text 原子性测试。
- [ ] F3. keyboard/Pointer/IME/Point/autocomplete/geometry Playwright 矩阵。
- [ ] F4. 大 workbook benchmark 记录 p50/p95/p99、long task、render count、heap。
- [ ] F5. 加入断言：逐键不触发 `UiSnapshot` generation、model/history/calc/persistence。
- [ ] F6. 运行 typecheck、boundaries、unit、build、E2E、`git diff --check`。
- [ ] F7. 使用内置浏览器验收真实交互、console、network。
- [ ] F8. 用真实 XLSX rich text/validation/formula corpus 验证导入编辑导出；桌面 Excel 不可用则标记 Blocked。
- [ ] F9. PR 描述记录实现、删除项、性能证据、回滚与 Blocked；全部满足后关闭 #301。

## 18. 验收判定

只有以下条件全部成立才关闭 Issue #301：

1. `rg` 与类型图证明不存在旧编辑状态源/旧业务 keydown/并行 DTO。
2. 每个入口和 Surface 都通过同一个 typed intent dispatcher。
3. 每个 commit 都进入 canonical command/runtime/model 链。
4. 成功、拒绝、冲突、validation pending、IME 和 popup priority 均有测试证据。
5. 性能 benchmark 达标且证明耗时不随 workbook cell count 线性增长。
6. 内置浏览器真实操作、console 和 network 无相关错误。
7. rich text/validation/formula 的真实 XLSX round-trip 有证据或明确 Blocked。
8. PR checks 通过并合并；Issue 验收项逐条回读后才能关闭。

## 19. 回滚策略

该改动是 clean-break，禁止在运行时切回旧链。回滚只能整体 revert PR，使代码、类型、协议、测试与 UI 同步恢复；不得只恢复某个 Surface 的旧 keydown 或 `editingActiveRef`。
