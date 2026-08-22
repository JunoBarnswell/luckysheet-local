# LuckySheet Local Fork

本仓库是 **dream-num/Luckysheet 的独立维护 fork**，不是官方继续开发，也不是迁到 Univer。

- 官方仓：[dream-num/Luckysheet](https://github.com/dream-num/Luckysheet) 已 **archived**（约 2025-10-30），见 [#1454](https://github.com/dream-num/Luckysheet/issues/1454)、[#799](https://github.com/dream-num/Luckysheet/issues/799)。
- 官方推荐继任产品是 [Univer](https://github.com/dream-num/univer)，但 Univer **不是** LuckySheet 源码超集，不能 1:1 替换。
- 本 fork 保留 `window.luckysheet` API 与 `celldata` / `data` / `config` / `calcChain` JSON 契约，基线版本 **2.1.13**。

变更记录：[CHANGELOG.md](CHANGELOG.md)（本 fork）。上游 2.1.13 及更早：[frontend/CHANGELOG.md](frontend/CHANGELOG.md)。

---

## 维护范围 vs 官方 EOL

| | 官方 Luckysheet | 本 fork |
|---|---|---|
| 状态 | EOL，archived，只读 | 在本仓库继续修 P0/P1 与对齐 OSS 语义 |
| 发布 | 停在 2.1.13 | 见下方版本策略；尚未打 npm tag |
| 目标 | 引导迁移 Univer | **继续用 LuckySheet API 的应用** 能修公式/滚动/筛选等缺陷 |
| 不做 | — | 不重写成 Univer TS/React monorepo |

`frontend/README.md` 顶部的「请改用 Univer」是**官方原文**，对本 fork 的适用范围以本文为准。

---

## 与 Univer 的关系

- **借鉴**：只对照 `@univerjs/*` **OSS** 的公开语义与 CHANGELOG（如 dirty range、筛选列同步、SUBTOTAL 跳过隐藏行、IME composition）。
- **不迁移**：不把工作簿改成 Univer snapshot / 插件架构。
- **不抄 Pro**：禁止复制 `@univerjs-pro/*` 闭源包。

字段对照见 [`docs/field-mapping-ls-uv.md`](docs/field-mapping-ls-uv.md)。证据索引见 [`docs/evidence-book.md`](docs/evidence-book.md)。

---

## Pro 边界（明确不做）

下列能力在 Univer 侧属于 **Pro / Server**，本 fork **不实现、不移植源码**：

| 能力 | 本 fork 做法 |
|---|---|
| Chart（Pro） | 沿用 ChartMix；依赖改为本地 vendor，离线可加载 |
| Pivot（Pro `engine-pivot`） | 加固现有 `pivotTable.js` + `pivotTableBoundary`，不抄闭源引擎 |
| Print（Pro） | **Blocked**：`frontend/src/expendPlugins/print/print.js` 为 **0 字节**，见 [`docs/print-blocked.md`](docs/print-blocked.md) |
| OT 协同（Pro） | 现有 WebSocket + last-write-wins；不实现 OT。见 [`docs/collab-protocol.md`](docs/collab-protocol.md) |
| Import/Export Server | 走本仓 `luckyexcel-node` / `exportXlsx`，**不依赖 Univer Server** |

---

## 已修复（相对官方 2.1.13）

| Issue | 现象 | 本 fork |
|---|---|---|
| [#1004](https://github.com/dream-num/Luckysheet/issues/1004) | `setRangeValue` / `setCellValue` 带 `f`+`v` 不重算 | 合法 `f` 一律进公式路径；删 `file.data` 回滚；`markFormulaDirty` |
| [#504](https://github.com/dream-num/Luckysheet/issues/504) | 表格偏下首次点击整页滚顶 | `focusEditor(..., { preventScroll: true })` |
| [#794](https://github.com/dream-num/Luckysheet/issues/794) | 嵌套页单击跳顶 | 同上，覆盖编辑器 / 公式栏 / 筛选容器 |
| [#529](https://github.com/dream-num/Luckysheet/issues/529) | `destroy`→`create` 滚动条错位 | `applySheetScroll` 单点恢复；destroy 不改调用方 `scrollTop` |
| SUBTOTAL + 筛选 | 隐藏行仍计入 | `SUBTOTAL` 跳过 `config.rowhidden` |

官方仓已 archived，上述 issue 在 upstream **不会关闭**。

---

## Phase 2–4 能力摘要

**Phase 2 存储与渲染**

- 稀疏 `flowdata`：10 万逻辑行只存占用格；`cloneSheetData` 隔离引用。
- `dirtyRect` 局部重绘；合并格 merge-aware 扩展。
- 透视占位/数据框按 `pivotTableBoundary` 绘制。

**Phase 3 表格语义**

- CF / DV adapter；筛选列同步与 CUSTOM AND/OR；排序跳过隐藏行并拦截数组公式相交。
- Table 对象、Cell Note（黄三角）与 postil（红三角）并存。
- IME composition、查找公式/跨表、特殊粘贴、命名区域超链接、万/亿格式保留。

**Phase 4 插件与后端**

- 图表 CDN → `frontend/src/expendPlugins/chart/vendor/`。
- 导出增强 + 错误处理；协同版本号 / 切表 / 重连。
- 打印 Blocked；多实例仅评估。

---

## 版本策略

| 线 | 何时升 | 例子 |
|---|---|---|
| **2.1.x patch** | 只修正确性，不扩 JSON/API 能力 | #1004、#504、#794、#529、SUBTOTAL+筛选 → 记为 2.1.14 类 |
| **2.2.x minor** | 新字段或新能力 | Table / Note / CF·DV adapter / 稀疏网格 / 图表 vendor / 导出增强 → 记为 2.2.0 类 |

不把 Table/Note/adapter 塞进 2.1 patch。正式打 tag 前 `package.json` 可仍为 `2.1.13`。

---

## 构建与回归

```bash
cd frontend
npm run build
node tests/regression/run-sparse-grid.cjs
node tests/regression/filter-phase3.mjs
```

```bash
node luckyexcel-node/scripts/verify-export.js
```

`verify-export.js` **不启 HTTP 服务**，只对 `luckyToXlsx` 写盘逻辑做断言。整页导出需自启 `luckyexcel-node` 再 POST `/luckyToXlsx`。

浏览器 HTML（需先 `npm run build`，再用静态服务打开，避免 `file://` 拦脚本）：

| 页 | 路径 | 期望 |
|---|---|---|
| 公式联动 | `frontend/tests/regression/setRangeValue-formula.html` | A1 改为 10 后 B1=`20` |
| focus 不滚顶 | `frontend/tests/regression/focus-scroll.html` | 单击/进编辑后 `window.scrollY` 不变 |
| destroy 滚动 | `frontend/tests/regression/destroy-create-scroll.html` | 无 `scrollTop` 再 create 回 0；`scrollTop:300` 首次与 recreate 一致 |
| SUBTOTAL | `frontend/tests/regression/subtotal-filter.html` | 隐藏 2、4 后 SUBTOTAL=`9` |
| 稀疏性能 | `frontend/tests/regression/sparse-perf.html` | 只采集 Δheap/FPS，数字写入 `docs/perf-baseline.md`，不编造 |

本机未开浏览器时，**不要把上表标成已通过**。

---

## 文档索引

| 文档 | 内容 |
|---|---|
| [CHANGELOG.md](CHANGELOG.md) | 本 fork 变更 |
| [docs/evidence-book.md](docs/evidence-book.md) | 模块证据册 |
| [docs/api-inventory.md](docs/api-inventory.md) | 对外 API |
| [docs/issue-module-map.md](docs/issue-module-map.md) | OPEN issue → 文件 |
| [docs/field-mapping-ls-uv.md](docs/field-mapping-ls-uv.md) | LS ↔ UV 字段 |
| [docs/print-blocked.md](docs/print-blocked.md) | 打印 Blocked |
| [docs/export-fidelity.md](docs/export-fidelity.md) | 导出保真度 |
| [docs/collab-protocol.md](docs/collab-protocol.md) | WebSocket 协议（非 OT） |
| [docs/cell-note-vs-postil.md](docs/cell-note-vs-postil.md) | Note vs 批注 |
| [docs/perf-baseline.md](docs/perf-baseline.md) | 性能基线（未测不填） |
| [docs/multi-instance-eval.md](docs/multi-instance-eval.md) | 多实例评估 |
