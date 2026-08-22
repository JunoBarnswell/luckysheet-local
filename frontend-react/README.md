# React Sheets

独立的 React + Canvas + Node 全栈 Sheets 重实现目录。该目录不引用现有 Luckysheet、Univer OSS 或 Univer Pro 运行时包，也不包含 License 模块。

## 当前链路

```text
React UI
  -> CommandRuntime
  -> Mutation / Operation
  -> WorkbookModel / Sparse CellMatrix
  -> FormulaEngine / RenderPlan
  -> Canvas Scene / Layer / Viewport
  -> Node HTTP / WebSocket / SQLite
```

## 开发

```powershell
npm install
npm run dev:server
npm run dev:web
```

Web: `http://127.0.0.1:4180/`

Server health: `http://127.0.0.1:4181/health`

## 已完成的垂直切片

- `WorkbookSnapshotV1` 与稀疏 `CellMatrix`。
- Command、Mutation、Operation、Undo/Redo 注册运行时。
- 无动态执行的 Lexer/AST/FormulaEngine、RangeIndex、循环错误。
- Canvas Scene、Layer、Viewport、SheetSkeleton、RenderPlan、dirty range 和滚动计划。
- React Ribbon、FormulaBar、Canvas Sheet、Sheet Tabs、Status Bar、Feature Sidebar。
- 独立 SQLite Snapshot/Revision/Changeset 存储和 WebSocket presence/changeset ACK。
- 前端真实创建 Workbook 并提交 cell/sheet Mutation。

## 尚未关闭的范围

图表、透视表、形状、Sparkline、打印、XLSX 交换、编辑历史、完整协同冲突变换、数据验证、条件格式和完整键盘/剪贴板语义仍需按 `packages` 分层继续实现。当前 UI 的未注册功能会显示明确 disabled 状态，不使用 Mock 冒充完成。
