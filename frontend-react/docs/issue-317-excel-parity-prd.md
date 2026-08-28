# Issue #317 — Excel Core Worksheet 95% Parity PRD

## 1. 产品目标

本 PRD 将 GitHub Issue #317 定义为当前工作簿编辑器的总验收合同，而不是一组独立的按钮补齐任务。目标产品基线为 Microsoft Excel for Microsoft 365 / Windows 的 Normal worksheet view、100% UI zoom；用户在 Issue 中提供的中文 HOME / INSERT 截图是固定 Ribbon 结构与视觉基线，Microsoft 官方资料是行为语义基线，SpreadJS v19 是浏览器交互深度参考。

最终可度量目标：

- HOME / INSERT 固定可见入口覆盖率 100%。
- Microsoft 官方 Windows 快捷键收录率 100%，并由统一语义 resolver 决定 owner。
- Excel Core Worksheet 可执行行为一致率不低于 95%。
- Cell / Grid / Canvas 核心交互一致率不低于 95%。
- HOME / INSERT Wide 视觉结构一致率不低于 95%，关键 group 与 surface 边界 100% 同构。
- Native Excel round-trip 不静默降级为 100%；不能编辑的原生内容必须 preserve-only，不能安全转换的能力必须显式 fail-close。

## 2. 当前基线与真实差异

开发基线是 Issue 创建时的最新 `github/main`（已包含 #301、#305、#311、#316 的已合入链路）。本 Issue 在该基线上继续收敛 Excel Core Worksheet 的统一语义与验收证据。

当前代码事实与主要缺口：

1. `ui-command-catalog`、`ShortcutRegistry`、KeyTips 已收敛到同一 registry/resolver；剩余验收重点是实际 UI 事件消费证据。
2. HOME / INSERT 的 Forms、Screenshot 已按宿主能力分类，不再用本地 button 或 worksheet snapshot 冒充 Excel 等价执行器。
3. `CellContentLayoutDomain` 现在统一 static、edit、AutoFit、rich-text、overflow、rotation、merge 和 caret 几何。
4. HeaderInteractionDomain 与 RangeDragDomain 已成为同一表头/选区交互链，行列菜单和 selection border drag 共用结构命令。
5. `GoToSpecialKind`、Fill Series、Conditional Formatting Rules Manager、Format Cells 已补齐对应的 typed surface 与命令入口。
6. Drawing selection、marquee、Tab z-order navigation、transform/arrange 与 Canvas 输入 owner 已接到 DrawingRuntime。
7. NativeDocument 与 Parity Manifest 已汇总 host/preserve-only/fail-close 分类；真实桌面 Excel 与 Browser 证据仍需在验收环境中完成。

## 3. 产品与架构设计

### 3.1 唯一事实源

新增 `ExcelParityManifest`，每一个原子能力包含：

```ts
type ExcelParityScope =
  | 'home' | 'insert' | 'shortcut' | 'cell' | 'grid' | 'selection'
  | 'clipboard' | 'table' | 'drawing' | 'object' | 'visual' | 'native-io';

type ExcelParityClass =
  | 'core-executable'
  | 'workbook-host'
  | 'external-office-host';

interface ExcelParityItem {
  id: string;
  scope: ExcelParityScope;
  class: ExcelParityClass;
  officialSource: string;
  commandId?: string;
  status: 'pass' | 'fail' | 'preserve-only';
  testId?: string;
}
```

Manifest 只记录真实 owner 与真实测试证据，不以按钮存在作为 pass。`core-executable` 进入 CoreParity 分母；系统窗口、VBA IDE、Copilot、Online/Stock 服务等归入 `workbook-host` 或 `external-office-host`，仍然必须被收录，但不伪造为本地可执行能力。

### 3.2 统一执行链

```text
Ribbon / KeyTip / Context Menu / Contextual Tab / DOM KeyboardEvent
                              ↓
                     ExcelFeatureRegistry
                              ↓
             CanonicalKeyGesture + Scope/State Resolver
                              ↓
                      Domain Intent / Command
                              ↓
        WorkbookModel + history + collaboration + permission
                              ↓
                Render / UI projection / NativeDocument
                              ↓
                         Parity evidence
```

同一 `featureId` 只能绑定一个 domain command。UI 只投影 registry，Canvas 只做 hit-test 和 gesture normalization，服务层不反向拥有 UI 状态。

### 3.3 状态所有权

| 状态 | 唯一 owner | 允许消费者 | 禁止行为 |
| --- | --- | --- | --- |
| 工作簿值、公式、样式、结构、对象 | `core-model` / domain command | render、history、native writer | UI 直接写模型 |
| Cell display/edit layout | `CellContentLayoutDomain` | static renderer、AutoFit、CellEditOverlay | 字符数经验估算、第二套 layout |
| 键盘 scope、KeyTips 层级 | `ExcelShortcutResolver` / `KeyTipState` | Canvas、Ribbon、Dialog、FormulaBar | 各组件各自 `if key` |
| 行列选择、resize、hide、insert/delete | `HeaderInteractionDomain` | Canvas、context menu、Ribbon | Header UI 手工改索引 |
| cell selection border move/copy/insert | `RangeDragDomain` | Canvas pointer layer、clipboard、StructuralTransform | 复制一套 header/canvas 逻辑 |
| Drawing selection/transform/arrange | Drawing domains | Canvas、Selection Pane、contextual Ribbon | 按 drawing kind 重复实现 |
| 原生内容所有权 | `NativeDocumentArtifact` | import/export、capability report | 简化对象覆盖未知原生 part |
| Parity 结果 | `ExcelParityManifest` | CI、PR、验收文档 | 手工声称完成 |

所有前置契约、权限、数据完整性或迁移状态不满足时 fail-close，保留错误码、对象和恢复动作，不返回默认值或假成功。

## 4. 实施切片与 TODO

### P0-A：Manifest 与统一 Surface Registry

- [x] 新建 `excel-parity.ts`，声明 HOME、INSERT、shortcut、cell/grid、drawing/object、native-io 的原子能力与官方资料。
- [x] 把现有 `RIBBON_COMMAND_CATALOG`、`RIBBON_TAB_SURFACES`、shortcut binding、contextual tab 和 context menu descriptor 接到同一个 `ExcelFeatureRegistry`。
- [x] 为每一项提供 command owner、permission、precondition、pass/fail/preserve-only 与 testId。
- [ ] 生成 `ExcelParityReport.json` 与人类可读摘要，不能用渲染数量替代执行证据。

### P0-B：Shortcut / KeyTips

- [x] 把 `ShortcutRegistry` 升级为 `CanonicalKeyGesture` + scope/state resolver，覆盖 Grid、Cell Edit、Formula Point、Formula Bar、Dialog、Ribbon、Pivot、Table、Drawing、TextBox、Comment。
- [x] 删除 `Ctrl+P -> commandPalette`，恢复 Print；Command Palette 使用不冲突的显式入口。
- [x] 收录 Issue 规定的完整 Microsoft Windows shortcut manifest，宿主级能力以 `workbook-host` / `external-office-host` 分类。
- [ ] 实现并通过真实 UI 验证 Alt/F10 KeyTips 的完整 Tab/Arrow/Enter/Space/Down/Esc 迁移和 `preventDefault` 合同；当前字母 KeyTips 已实现。
- [x] 将 Alt+F1、F11、Ctrl+E、F8、Shift+F8、Ctrl+Q、Ctrl+L/T、Ctrl+Shift+F2、Ctrl+Alt+V 等纳入统一 resolver。

### P0-C：CellContentLayout / Grid / Header

- [x] 将 `CellContentLayoutDomain` 作为唯一布局所有者，统一 static/edit/AutoFit 的 font runs、lines、rotation、indent、merge、zoom、caret geometry。
- [x] 删除 `draft.length * fontSize * 0.56` 和独立编辑器换行计算。
- [x] 实现 Left/Right/Center/General/Center Across Selection 的 `OverflowSpanResolver`；Wrap、Shrink、Fill、Justify、Distributed、merge、数字日期按规则阻断。
- [x] 让编辑 surface 可临时覆盖邻格，按真实 metrics 扩张并保证 caret/IME 可见，不改变邻格模型。
- [x] 将 Wrap、手工 newline、固定行高、AutoFit row/column 与统一 layout 关联。
- [x] 建立 `HeaderInteractionDomain`，同构支持 row/column/corner selection、Shift/Ctrl 非连续选择、resize preview、double-click AutoFit、hidden double-line unhide、context menu、insert/delete、clear、clipboard。
- [x] 所有本次整行整列结构变化接入 `StructuralTransform`，并沿用既有公式、merge、name、table、filter、validation、CF、comment、hyperlink、sparkline、pivot、drawing、print refs 和 selection 变换链。
- [x] 建立 `RangeDragDomain`，实现 selection border move/copy/insert 与整行整列复用 clipboard/structural command。

### P0-D：HOME 深层行为

- [x] 扩展 Format Cells 的 Number/Alignment/Font/Border/Fill/Protection typed model、混合状态、主题色、对角线边框、pattern/gradient、negative/date/locale 等属性。
- [x] 修复 `rotateDown` 的错误 180° 语义，采用 Excel vertical/down 方向模型。
- [x] 扩展 Conditional Formatting 为 Rules Manager：Edit、Applies To、优先级移动、Stop If True、formula/date/unique/average/native threshold。
- [x] 扩展 Go To Special：current array、row/column differences、precedents、dependents、All/Same conditional/data validation、constants/formulas subtype。
- [x] 实现 Fill Series 对话框（Rows/Columns、Linear/Growth/Date/AutoFill、Step/Stop/Trend/Date unit）、Fill Options、Justify、Flash Fill/Ctrl+E。
- [x] 让 HOME/INSERT Table 共用 Table domain；Forms 改为外部宿主能力，不调用 `createFormControl('button')` 冒充。

### P0-E：Drawing / Object Canvas

- [x] 建立 Drawing selection/transform/arrange/text-edit 的统一命令入口。
- [x] 支持 click/Ctrl/Shift/marquee、Tab/Shift+Tab z-order、Ctrl+Shift+Space、Selection Pane show/hide/reorder、group member、Esc 回到 grid。
- [x] 统一 move/resize/rotate/nudge/Ctrl-drag duplicate、snap、align/distribute、z-order、contextual tabs。
- [x] Shapes、Pictures、Controls、Chart、TextBox、SmartArt、WordArt、OLE identity 复用基础 selection/transform/arrange，不按对象类型复制逻辑。
- [x] Screenshot 改为明确 Camera/workbook snapshot 语义；系统窗口截屏归 host，不能伪称 Excel Screenshot。
- [x] Icons/3D/SmartArt/WordArt/Signature/Equation/OLE 对不可执行的宿主能力做 native preserve 或 fail-close 分类，不生成假等价对象。

### P0-F：Native I/O 与证据

- [x] 对齐 `NativeDocumentArtifact` ownership，round-trip wrapText、shrinkToFit、newline、column width/hidden/bestFit、row height/hidden/customHeight。
- [x] 未知合法 part/node/extLst 保留；已编辑 owned region 只改对应 native region。
- [x] 对 XLSX/XLSM/XLSB/BIFF 能力使用真实 codec 结论，不把 preserve-only 声称为 editable。
- [x] 任何不能安全转换的 native object 返回明确 `UNSUPPORTED_FEATURE`，不自动降级或双写。

### P1：视觉与高频交互

- [ ] 以 Issue 的 HOME/INSERT 中文截图为 Wide golden：group 顺序、separator、control hierarchy、split hit region、launcher、label wrapping、icons、hover/pressed/checked/disabled/focus-visible/mixed/menu-open。
- [ ] Grid golden 覆盖 active border、copy border、fill handle、formula reference、header selection、freeze/zoom/hidden/filter/validation/comment overlays。
- [ ] Drawing golden 覆盖 selection box、handles、rotation、multi-select、snap/alignment guide、marquee、z-order cursor。
- [ ] 所有业务 UI 使用既有共享组件与 Tailwind token，页面不散落原生交互元素或一-off 样式。

### P2：宿主与高级扩展

- [ ] 记录 Windows window、VBA IDE、Copilot、Online/Stock content、system Screenshot、OLE activation、digital signature execution 的 host classification 与 preserve contract。
- [ ] SpreadJS 扩展只进入 extension scope，不改变 Excel 固定 Ribbon 和 CoreParity 分母。

## 5. 失败与拒绝合同

- 无法解析快捷键 scope、KeyTips 层级或 command owner：拒绝消费事件，返回 `SHORTCUT_SCOPE_UNRESOLVED`，不让浏览器默认行为偷偷抢占 Excel Core shortcut。
- 无法取得真实字体/layout 或 merge/pane 几何：拒绝提交编辑/resize，不用默认尺寸掩盖问题。
- 结构操作无法计算完整 metadata transform：中止整笔 transaction，保留原模型和受影响范围。
- Forms、Screenshot、OLE、Signature、Equation、SmartArt、3D 等宿主能力没有权威执行器：标记 host/preserve-only 或 `UNSUPPORTED_FEATURE`，禁止本地简化对象伪装为 Excel 等价能力。
- 原生 OOXML/Binary part 无法安全映射：保留未知内容，禁止删除、静默转换或生成第二套写入链。

## 6. 验收矩阵

### Manifest / architecture

- [ ] `ExcelParityManifest` 覆盖所有 Issue 原子能力，官方快捷键收录率 100%。
- [ ] 只有一个 Feature Registry、一个 Shortcut Resolver、一个 CellContentLayout、一个 HeaderInteraction、一个 DrawingSelection/Transform/Arrange owner。
- [ ] `rg` 检查不存在旧 `Ctrl+P`、字符数 editor geometry、Forms→button 假映射、Screenshot→Excel 原生假映射和新旧双写 registry。

### Functional / structural

- [ ] Cell layout unit：CJK/Latin/mixed、bold/italic、newline、wrap、shrink、left/right/center、occupied neighbor、merge、center-across、rotation、indent、rich text、zoom。
- [ ] Editor browser：长文本、右对齐、居中、Alt+Enter、Wrap、fixed row、AutoFit、IME、merge、邻格占用、viewport edge、freeze、50/100/200% zoom。
- [ ] Header browser：单/多/非连续 row/column、Shift/Ctrl、corner、resize、AutoFit、hide/unhide、context menu、insert/delete、clear、whole-row/column clipboard。
- [ ] Structural undo/redo：公式、merge、names、table、filter、validation、CF、comments、hyperlinks、sparklines、pivots、drawings、print refs、selection。
- [ ] Native round-trip：XLSX/XLSM/XLSB/BIFF 适用 corpus；未知内容 hash 与 capability report 可解释。

### Gate

```text
CoreParity >= 0.95
HOME visible surfaces = 1.00
INSERT visible surfaces = 1.00
official shortcut catalog coverage = 1.00
no silent native-loss = 1.00
```

## 7. 视觉设计思路

Issue #317 的设计图基线是 Issue 中提供的 Microsoft Excel 中文 HOME / INSERT 截图：Wide 模式固定 group 顺序、分隔线、二/三行 control hierarchy、split button、dialog launcher 和 label wrapping；Compact/Narrow 只改变密度与溢出，不改变 feature ownership。实现复用现有 `ui-system` 的 `Button`、`DropdownMenu`、`Stack`、`Inline`、`Panel`、`Dialog`、`Text` 和 Tailwind token，不重新创建视觉组件体系。

Cell/Canvas 使用分层视觉模型：模型背景/边框层、文字布局层、selection/focus 层、transient gesture 层、native/error notice 层；overflow 不能擦掉邻格背景或边框，selection 必须保持最高相关 z-order，编辑 surface 无圆角且跟随 pane/zoom/freeze/merge 几何。

## 8. 交付策略

在 `codex/issue-317-excel-parity` 专用分支开发。所有 P0/P1 语义修改完成后再执行集中验证，不在开发中频繁编译。最终按一个 issue 一个提交形成单一功能提交，经 Draft/Ready PR 交付；PR 描述记录实现摘要、契约变化、验证证据、明确 Blocked 项与回滚方式。

## 9. 风险与回滚

- 风险：Issue 横跨前端、核心模型、原生 I/O 与浏览器宿主，任何第二套 resolver 或 UI 假状态都会造成长期语义漂移。
- 风险：当前 main 已含 #311/#305，但 #316 未合入，Cell geometry 变更可能影响 render-engine、Canvas 和 formula/editor overlay。
- 风险：外部宿主能力和真实桌面 Excel 可能不可用；只能报告 Blocked/preserve-only，不能伪造 acceptance。
- 回滚：以该 Issue 的单一提交整体 revert；不通过删除文件或 reset 工作区回滚用户改动。
