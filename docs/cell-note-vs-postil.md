# Cell Note 与批注（postil / `ps`）

LuckySheet 原有批注是单元格上的 **`ps`** 对象，由 `frontend/src/controllers/postil.js` 渲染成可拖动、可编辑的批注框，画布右上角画红色三角。

本 fork 新增的 **Note** 是另一套轻量备注，对照 Univer OSS `@univerjs/sheets-note`（CHANGELOG `#5125`），**不替代、不迁移 `ps`**。

## 并存规则

| | 批注 `ps` | Note |
|---|---|---|
| 存储 | `cell.ps`（value/isshow/left/top/width/height） | `cell.note` 或 `sheet.notes["r_c"]`；工作表级 `sheet.note` |
| UI | 可编辑浮动框 + 红三角 | 黄/橙三角 + hover 文本 |
| 交互 | 单击进入编辑、拖动、缩放 | 仅悬停展示，不打开 postil 框 |
| 协同键 | 随 cell 写入 | `notes` / `note` |
| 产品角色 | Excel 批注（Comment） | 单元格/工作表备注（Note） |

同一单元格可以同时有 `ps` 和 `note`：

- 红三角仍表示批注（靠单元格右上角外侧）。
- 黄三角略向内偏，表示 Note。
- hover 时 Note 用浅黄浮层；批注框逻辑不变。

## JSON

```json
{
  "v": 1,
  "m": "1",
  "ps": { "value": "这是批注", "isshow": false },
  "note": { "text": "这是 Note" }
}
```

工作表级：

```json
{
  "name": "Sheet1",
  "note": { "text": "本表用于对账" },
  "notes": {
    "0_0": { "text": "A1 备注" }
  }
}
```

`sheet.notes` 与 `cell.note` 指向同一语义；运行时 `noteCtrl` 会把 map 同步到格子上。旧工作簿没有这些字段时行为与原来完全一致。

## 不要做的事

- 不要把 `ps` 自动改写成 `note`，也不要把 Note 画成 postil 框。
- 不要引入 Univer thread-comment（`#6042`）数据模型。
- 不要改 `jfundo` / `jfredo` 命名。
