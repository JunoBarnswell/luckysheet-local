# Issue #316 Cell UI Geometry and Header Interaction PRD

## 1. 产品目标

基于 `main` 的审查基线 `7d0c99f26f5708d207768110207ebd921c439b7b`，把单元格静态显示、行内编辑、文本测量、行列尺寸和表头交互收口为一个可验证的 Excel/SpreadJS 行为域。

用户看到的文字、编辑框和行列尺寸必须由同一份字体测量、换行结果、旋转/缩进占位、合并几何和 PaneMap 屏幕变换产生。禁止依赖字符数经验系数、邻格占用的静态规则或 UI 私有状态制造第二套几何。

## 2. 真实基线与问题

- `render-engine/src/cell-text-layout.ts` 已经是静态文字和 AutoFit 的测量入口，但结果只描述文字布局，编辑框仍在应用层另算尺寸。
- `cell-renderer.ts` 只把左对齐文本向右溢出；右对齐、居中、Center Across Selection、合并边界和数字/日期的阻断语义没有统一解析。
- `CellEditOverlay.editorRect()` 仍使用 `draft.length * fontSize * 0.56` 和邻格为空条件；右侧有值时长文本会被挤压，viewport 边缘和 freeze pane 也可能发生跳变。
- `CellEditor.tsx` 的普通文本输入使用 `overflow-hidden`，外层尺寸不足时 caret 和多行内容不可见。
- `SheetCanvas.tsx` 的行头、列头右键菜单由两个分支分别拼接，列头缺少整维结构操作，行头缺少剪贴板、清除和格式操作。
- `CanvasRenderEngine.contentRangeToScreenRects()` 能够返回多个冻结窗格矩形，但编辑器当前只取第一份矩形，不能保证 anchor pane 一致。

## 3. 统一产品链路

```text
Cell value / formula / rich text / style
                |
                v
       CellContentLayoutDomain
       - canonical text + font metrics
       - wrap / shrink / rotation / indent
       - merge and alignment span
       - static overflow boundary
       - edit surface + caret geometry
          |                         |
          v                         v
  Static Cell Layout          Edit Surface Layout
          |                         |
          v                         v
   Canvas Cell Renderer       CellEditOverlay / CellEditor
          \_________________________/
                      |
                      v
             Row/Column Geometry
                      |
                      v
             HeaderInteractionDomain
                      |
                      v
   selection / resize / AutoFit / hide / unhide / insert / delete / clipboard
```

### 3.1 `CellContentLayoutDomain`

唯一输入是 canonical cell projection、`CellRenderStyle`、cell/merged/alignment-span geometry、`mode: 'display' | 'edit'`、邻格可见性和当前 zoom。唯一输出包含：

- 实际字体和字体运行；
- 按真实 `measureText` 得到的行、宽、高；
- 静态显示矩形和 overflow span；
- 编辑矩形、文本内边距、可视内容区域和 caret 可视区域。

静态与编辑使用同一个测量器，但规则不同：静态 overflow 按对齐、占用、merge、presentation、viewport 边界停止；编辑 surface 以 draft、selection、caret、IME 可见为优先，可以暂时覆盖网格内容，不改变邻格模型和尺寸。编辑器 anchor 使用 `PaneMap.paneForCell()` 对应的屏幕变换；冻结窗格不取任意第一矩形。

### 3.2 `HeaderInteractionDomain`

表头只产生 typed intent，不直接改索引或拼接模型状态：

```text
select(target, additive, extend)
resize.begin/update/commit(target, pointer)
autofit(target)
hide/unhide(target)
insert/delete/clear(target)
move(target, destination, copy, insert)
```

Canvas 负责命中和指针/键盘手势，Context Menu 负责把同一 catalog 投影为菜单，CommandRuntime/StructuralTransform 负责实际事务。`ColumnDimensionController` 只负责 Excel 单位转换和 `CellContentLayoutDomain` 测量，不拥有 selection 语义。

## 4. 行为合同

### 单元格显示

- `General`/`Left` 向右、`Right` 向左、`Center` 向两侧寻找连续可用空间。
- `Center Across Selection` 先解析 canonical alignment span；不把邻格值当成编辑器边界。
- Wrap、Shrink、Fill、Justify、Distributed、Stacked、Barcode/Image/Checkbox presentation 不走普通文本 overflow。
- 数字/日期不因列窄而把原始文本溢出到邻格。
- 合并单元格以完整 merged rect 为绘制和编辑边界；邻格背景、边框和 overlay 不被文字层擦除。

### 编辑 surface

- `editRect.width >= cellRect.width`，随 draft 的真实字体宽度增长。
- Left/General 主要向右扩，Right 主要向左扩，Center 以 cell center 向两侧扩。
- 显式换行、Wrap、rich-text 多行使用真实行高；textarea 不以 `overflow-hidden` 掩盖不足。
- surface 在 viewport 边缘优先保留 anchor 和 caret；不足时使用内部滚动，不把 x 强行截断为 0。
- scroll、zoom、freeze、merge 变化只重新解析同一个 layout 结果，不创建第二个 editor。

### 行列头

- 列/行支持整维、连续、非连续、Shift 扩展、Ctrl/Meta 加选、Ctrl+Space/Shift+Space、corner Select All。
- 单维和多维 resize、双击 AutoFit、hidden double-line unhide 使用同一 intent 和事务。
- 行列右键菜单同构，具备 Cut、Copy、Paste/Paste Special、Insert、Delete、Clear Contents、Format Cells、Width/Height、Hide、Unhide。
- Insert/Delete 只调用现有 StructuralTransform/clipboard 域，统一更新公式、合并、名称、表格、筛选、校验、条件格式、批注、超链接、sparkline、Pivot、drawing、打印引用和 selection。

## 5. 状态和失败边界

- `CellContentLayoutDomain` 是纯几何计算；编辑 draft/caret/composition 仍由 `CellEditDomain` 持有，结果通过 typed props 进入 `CellEditor`。
- header selection、resize preview、AutoFit pending、hidden indicator 是瞬时 UI 状态；提交只产生一个 canonical command/history transaction。
- 无法解析字体、merge、pane、selection 或 dimension contract 时 fail-close，返回可观察错误并阻止提交；不使用默认偏移、空结果或静默降级。
- 不新增 layout compatibility bridge、旧字段别名、双写或 UI-only repair；完成时删除 `editorRect()` 经验估算和行列头不对称菜单路径。

## 6. 验收与证据

- Layout unit tests：CJK/Latin/mixed、粗斜体/字体、newline、wrap、shrink、左/右/中溢出、邻格、merge、Center Across、rotation、indent、rich text、zoom，并校验静态/AutoFit/edit 共用 golden measurement。
- Browser tests：长文本、右/中对齐、多行、固定行高、AutoFit、IME、rich text、merge、非空邻格、viewport 边缘、freeze、50/100/200% zoom；每项检查 caret、裁切、anchor、提交和 Esc 恢复。
- Header browser tests：行列选择/加选/扩展、corner、resize、多选 resize、AutoFit、hide/unhide、右键多选、insert/delete/clear、整行列 clipboard 及镜像行为。
- OOXML tests：`wrapText`、`shrinkToFit`、rotation/indent/alignment、newline、列宽/行高/hidden round-trip。
- 最终门禁：typecheck、相关单测、全量单测、boundary/contract、build、真实内置浏览器 DOM/console/network/screenshot/interaction。
- 若当前环境无法使用真实 Excel producer 或桌面 executable，只将该项标记为 `Blocked`，不把本地模拟当成 Excel 互操作通过。

## 7. 参考资料

- Microsoft Excel：Wrap Text、列宽/行高、选择、插入/删除、隐藏/显示、移动/复制。
- MESCIUS/SpreadJS：Cell Overflow、Wrap Text、AutoFit、Editing、Headers、Resizing、Rows/Columns、Designer Context Menus。

## 8. 回滚

改动通过单一 `codex/issue-316-cell-ui` 分支和 PR 交付；回滚方式为回退该 PR。不得在默认分支直接提交或保留废弃的平行几何链。
