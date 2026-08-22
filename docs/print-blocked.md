# 打印能力 Blocked 评估

> Phase 4.4 评估文档。仅核实现状与可选方案，本轮不实施打印引擎。

## 1. 核实结论

`frontend/src/expendPlugins/print/` **存在名为 `print.js` 的文件，但长度为 0 字节，没有实现**。插件入口仍 `import { luckysheetPrint } from "./print"`，因此运行时 `luckysheetPrint` 为 `undefined`，打印插件无法初始化。

这与「缺 `print.js` 实现文件」等价：磁盘上有空壳，没有可执行打印逻辑。

## 2. 源码证据

| # | 路径:行号 | 要点 |
|---|---|---|
| 1 | `frontend/src/expendPlugins/print/plugin.js:2` | `import { luckysheetPrint } from "./print"` |
| 2 | `frontend/src/expendPlugins/print/plugin.js:6-14` | `dependScripts` / `dependLinks` 全部被注释，无真实依赖加载 |
| 3 | `frontend/src/expendPlugins/print/plugin.js:17-30` | `print()` 仅当 `luckysheetPrint` 真值时才摘掉 `Store.asyncLoad` 的 `"print"` 并挂 `Store.luckysheetPrint` |
| 4 | `frontend/src/expendPlugins/print/plugin.js:25-29` | 仅注入 `./expendPlugins/print/print.css` |
| 5 | `frontend/src/expendPlugins/print/print.js` | **文件存在，Length = 0**，无 `export` |
| 6 | `frontend/src/expendPlugins/print/print.css` | 仅有样式，无分页/打印引擎 |
| 7 | `frontend/src/controllers/expendPlugins.js:2` | `import { print } from '../expendPlugins/print/plugin'` |
| 8 | `frontend/src/controllers/expendPlugins.js:7` | 注册表 `'print': print` |
| 9 | `frontend/src/controllers/expendPlugins.js:16-20` | 文档示例 `plugins:[{name:'print'}]` |
| 10 | `frontend/src/core.js:27` | `// import { printInitial } from "./controllers/print"` 已注释 |
| 11 | `frontend/src/core.js:198` | `// printInitial();` 工作簿初始化不走打印 |
| 12 | `frontend/src/store/index.js:153` | `asyncLoad:['core']`，插件名会 push 进异步队列 |
| 13 | `frontend/src/core.js:146` | `Store.asyncLoad.push(...plugins.map(plugin => plugin.name))`，启用 print 后若实现缺失会卡住异步完成标记 |
| 14 | `frontend/README.md:32` | 官方引导高级打印使用 Univer |
| 15 | `frontend/README-zh.md` | 停维声明将「打印」列为 Univer 新增能力 |

目录实测（2026-08-22）：

```
plugin.js   1125 bytes
print.css    826 bytes
print.js       0 bytes   ← 空壳
```

不存在 `frontend/src/controllers/print.js`（`core.js` 的 import 已注释）。

## 3. Blocked 原因

1. **实现文件为空**：`print.js` 无法导出 `luckysheetPrint`，`plugin.js:22` 的 `if (luckysheetPrint)` 永远不进入。
2. **历史依赖被注释**：曾经打算从 `luckysheetPluginPrint.umd.js` / localhost:8080 动态加载，当前仓库无该 UMD。
3. **核心初始化切断**：`printInitial` 不再接入 `initialWorkBook`。
4. **Univer 打印为 Pro**：OSS 仓无 `@univerjs-pro` 打印包可抄；计划已禁止复制 Pro。
5. **产品契约缺口**：`window.luckysheet` 对外 API 清单中无 `print` / `printSheet` 方法，文档也未给出可调用打印 API。

因此 Phase 4 打印标记为 **Blocked**：不是「差一点就能开」，而是缺少完整打印引擎与 API 契约。

## 4. 可选方案（不实施）

| 方案 | 做法 | 工作量 | 风险 |
|---|---|---|---|
| A. 浏览器 `window.print()` | 对当前可见区或克隆后的表格 DOM 调用系统打印 | 小（1–3 天出占位） | 冻结/图表/canvas 失真；分页不可控 |
| B. `html2canvas` + 分页 | 仓库已有 `frontend/src/plugins/js/html2canvas.min.js`；按 `visibledatarow` 切片画布再拼 PDF/打印页 | 中（1–2 周） | 大表内存；合并单元格跨页；与 Excel 页边距不对齐 |
| C. 自研 `print.js` + 工具栏入口 | 补齐空壳 `print.js` 导出 `luckysheetPrint`，实现预览对话框、页眉页脚、缩放、选定区域 | 大（3–6 周） | 需新 JSON 字段（页边距/纸张）；协同广播未定义 |
| D. 导出 xlsx 后交给 Excel 打印 | 走现有 `exportXlsx` / `luckyexcel-node` | 小 | 依赖导出保真度；不是应用内打印 |
| E. 采购 Univer Pro Print | 不在本 fork 范围 | — | 与「保留 LuckySheet API」目标冲突 |

**建议（评估，非实施）**：短期用方案 D 满足「能打出来」；中期若必须应用内打印，选 B 作为最小自研，先做当前 sheet 可见区，再补分页。不要从 Univer Pro 搬打印引擎。

## 5. 验收口径（若未来解 Blocked）

- `print.js` 非空，且 `export { luckysheetPrint }` 有真实对象。
- `plugins:[{name:'print'}]` 后 `Store.asyncLoad` 不再残留 `"print"`。
- 至少能打印当前 sheet 的值 + 基础边框；公式打印显示值 `m` 而非 `f`。
- 文档声明与 Excel 分页/页眉页脚的保真度差距。
