# Issue #304 Remaining Closure PRD

## 1. 目标

基于 Issue #304 的原始功能清单和当前 `main` 真实源码，补齐 HOME/INSERT 的本地 canonical runtime 行为。任何正常功能入口都不得要求 Microsoft 365、Windows capture、数字证书、OLE server、在线素材库或其他外部宿主；只有单个无法表达的输入变体才在领域边界返回可观察的 typed error。

本轮引用的视觉证据：

- HOME 原图：`C:\Users\kuo13\AppData\Local\Temp\codex-clipboard-fe420d56-7b06-41bf-801d-250843c45638.png`（1783×141）
- INSERT 原图：`C:\Users\kuo13\AppData\Local\Temp\codex-clipboard-c9bac183-6f0e-4b48-a9e6-c1ca1048dabe.png`（1905×135）
- 图表下拉参考：`C:\Users\kuo13\AppData\Local\Temp\codex-clipboard-f72c6d8c-423a-464e-9025-ad3f124cb001.png`、`C:\Users\kuo13\AppData\Local\Temp\codex-clipboard-ca67831c-bd91-4be5-88f0-dcf3e4802ccb.png`

## 2. 范围决策

### 本地可闭环（本 PR 完成）

1. Paste Source Theme：源主题从 clipboard payload 进入 dialog、planner 和同一事务，主题变更可撤销；未带主题的外部 payload fail-close。
2. Number：补齐 Excel 基础 gallery（General、Number、Currency、Accounting、Short/Long Date、Time、Percentage、Fraction、Scientific、Text、More Formats），底层值不被格式操作改写。
3. Conditional Formatting：把现有规则域投影成 highlight、top/bottom、duplicate、data bars、color scales、icon sets 及 priority/Stop If True。
4. 本地对象：Forms、Icons、3D OBJ、SmartArt、Screenshot、WordArt、Signature Line、Embedded Object、Equation 均由独立 payload、命令、历史、渲染和本地资产/metadata 回环完成。
5. Chart Map：Map 使用本地确定性分格投影；没有可表达的输入时返回明确的 `INVALID_CHART_SOURCE`，不依赖地图宿主。
6. 结构变换：截图/相机源范围、绘图对象、图标和本地对象随行列/工作表复制保持同一坐标语义。

### 全部本地化（本 PR 必须完成）

本产品不得依赖 Microsoft 365、Windows capture、数字证书、OLE server、在线素材库或其他外部宿主。Forms、Icons、3D Models、SmartArt、Screenshot、WordArt、Signature Line、OLE Object、Equation 必须由本地 canonical model、本地资产、编辑器、渲染器和持久化链路完成；网络不可用时仍可使用已 vendor 的资源和用户本地文件。

签名行采用本地 signer metadata/status，不宣称数字证书签名；OLE 采用本地 embedded/linked file object，不激活外部 COM server；Screenshot 采集当前应用可见画布/工作区，不请求操作系统屏幕权限；Equation 采用本地 OMML 子集编辑器和 renderer。无法表达的 OOXML 结构必须保留 opaque parts 并在编辑入口明确拒绝该具体变体。

Shapes、generic Chart native OOXML 和桌面 Excel 的互操作仍遵循现有 typed capability 边界；本地应用行为不依赖这些外部宿主。无法写成当前 OOXML family 的单个图表变体必须保留原始数据或返回明确错误。

## 3. 统一数据链与接口契约

- 所有新增行为从 `RibbonSchema → WorkbookSession → CommandRuntime → model mutation → history/collaboration/persistence → render/OOXML` 进入；React 组件只负责投影和临时交互状态。
- 现有 CellStyle underline 语义保持与既有 workbook contract 一致；本地对象不复用 CellStyle 作为伪装状态。
- `PasteSpecialSpec` 的 `source-theme` 必须携带 `ClipboardPayload.rangeMetadata.sourceWorkbookThemeRef`，缺失时返回 `PASTE_SOURCE_THEME_UNAVAILABLE`。
- Number/Conditional/Fill/Print models 继续使用现有 canonical mutation registry，成功路径和拒绝路径各有测试。
- 只有无法表达的单个 OOXML 变体才返回错误码、feature、affected object、reason、recovery；正常入口必须创建本地 canonical payload，不创建空 payload、不写双份状态。

## 4. 验收标准

- 每个本地功能均有 model/command/history/undo-redo/collaboration/persistence/OOXML 证据。
- 每个 fail-close 条件都有拒绝测试，且事务前后 workbook/history 不发生部分写入。
- 内置浏览器验证 HOME/INSERT 的真实菜单、dialog、disabled/error/loading 状态；记录 console/network。
- 1783×141 HOME、1905×135 INSERT 和图表下拉参考截图完成组件级 Pixel Difference Report。
- `npm run typecheck`、`npm run test:unit`、`npm run check:boundaries`、`npm run build` 全部通过。
- 桌面 Excel 不可用时，Native Excel corpus 仍标记 `Blocked`，不能宣称通过；本地浏览器功能不能因此依赖 Excel。

## 5. 回滚与迁移

本轮是 canonical contract 的一次性升级。旧 underline/paste 运行时值不作为 fallback 读取；无法安全迁移的快照在 hydrate 边界返回 `MIGRATION_REQUIRED`。回滚使用本 PR 单一 revert，不恢复旧 runtime 双写路径。
