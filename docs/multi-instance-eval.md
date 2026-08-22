# 多实例评估：Store 单例改为实例上下文

> Phase 4.6 **仅评估不实施**。对照 Univer `IUniverInstanceService`，估算把 LuckySheet 全局 `Store` 改成实例上下文的工作量与 API 兼容方案。

## 1. 现状

LuckySheet 2.1.13 是 **进程内单工作簿单例**：

- `frontend/src/store/index.js:1-170` 导出一个普通对象 `Store`，不是工厂。
- `frontend/src/core.js:45` `luckysheet = common_extend(api, luckysheet)`，`api.js` 全部方法读同一个 `Store`。
- `frontend/src/core.js:48-49` **每次 `create` 先 `method.destroy()`**，后写入 `Store.container` / `Store.luckysheetfile`。
- 前端 `src` 下约 **74** 个 JS 文件 `import Store`（实测 rg 计数）。
- DOM 大量使用全局 id：`#luckysheet-cell-main`、`#luckysheet-rich-text-editor`、`#luckysheetTableContent`（`draw.js:436`、`handler.js`）。
- 公式实现挂到 `window.luckysheet_function`（`functionlist.js:33`）供 `eval`。
- 模块级单例还有：`sheetmanage`、`luckysheetformula`、`pivotTable`、`imageCtrl`、`dataVerificationCtrl`、`server`、`scroll.js` 的 rAF 旗标。

因此同一页面 **不能** 同时存在两个独立工作簿；「多实例」今天只能串行 `destroy → create`。这正是 [#529](https://github.com/dream-num/Luckysheet/issues/529) 滚动状态串扰的结构根因之一。

## 2. Univer 对照（可借鉴语义，不可整包搬）

`C:\Users\kuo13\Projects\univer\packages\core\src\services\instance\instance.service.ts`

| UV API | 行号 | 语义 |
|---|---|---|
| `IUniverInstanceService` | 76-119 | 多 Unit 注册表 |
| `createUnit` | 108 | 按 snapshot 创建 workbook |
| `disposeUnit` | 112 | 销毁指定 unit |
| `getUnit` / `getAllUnitsForType` | 116-117 | 按 id / 类型取实例 |
| `focusUnit` / `getFocusedUnit` / `focused$` | 97-101 | 同一 app 仅 1 个 focused unit |
| `setCurrentUnitForType` | 104 | 当前操作目标 |
| `unitAdded$` / `unitDisposed$` | 78 / 86 | 生命周期可观察 |

UV 还有：Command/Mutation 带 `unitId`；Facade `FUniver` 可挂多个 workbook；`focus-editor.ts` 用 `ownerDocument` 隔离。

LS 没有 `unitId`，所有 mutation 默认打在 `Store.currentSheetIndex`。

## 3. 必须拆开的状态面

| 分组 | 代表字段 / 模块 | 单例风险 |
|---|---|---|
| 工作簿 | `luckysheetfile`, `currentSheetIndex`, `container`, `toJsonOptions` | 两个 create 互相覆盖 |
| 网格运行时 | `flowdata`, `config`, `visibledatarow/column` | 绘制读错表 |
| 选区/剪贴板 | `luckysheet_select_save`, `luckysheet_copy_save` | 跨实例粘贴串数据 |
| 撤销 | `jfundo`, `jfredo`, `clearjfundo` | undo 打到另一本簿 |
| 公式 | `calcChain`（在 sheet 文件上）+ `formula.js` 模块字段 | eval / execFunctionGroup 全局 |
| 协同 | `server.gridKey/updateUrl` + `cooperativeEdit` | WebSocket 绑错文档 |
| 插件 | `asyncLoad`, `chartparam`, `luckysheetPrint` | 图表 DOM id 冲突 |
| DOM | 固定 `#luckysheet-*` | 第二个实例选择器命中第一个 |
| 冻结/筛选/DV/图片 | 各 controller 对象字段 | destroy 重置也不支持并存 |

## 4. 改造方案（评估）

### 方案 A：实例上下文对象（推荐评估方向）

引入 `createWorkbookContext()`，把今日 `Store` 字段变成实例。对外仍挂 `window.luckysheet`，但：

```text
luckysheet.create(options) → 返回 instanceId
luckysheet.getInstance(id) / luckysheet.use(id)
luckysheet.destroy(id?)     → 缺省销毁当前 focused
```

内部：`api.js` 每个方法先 `const ctx = getActiveContext()`，再读 `ctx.flowdata`。

**兼容**：旧调用 `luckysheet.setCellValue(...)` 继续打在 **当前 focused 实例**（对齐 UV `focusUnit`）。新调用可传 `options.instanceId`。

### 方案 B：完整多实例 + 多容器 DOM 前缀

在 A 之上把所有 `#luckysheet-*` 改为 `#${container}-*` 或容器内查询。必须改 `createdom.js`、`handler.js`、`draw.js`、`formulaBar.js`、`sheetBar.js`。

没有 B，A 只能做到「同一时刻一个可见实例」，与今天 `destroy+create` 差别不大。

### 方案 C：iframe 隔离（产品 workaround，不改架构）

每个表格一个 iframe 加载独立 UMD。API 用 `postMessage`。工作量小，但丢失同页公式互引、主题一致、协同复用。

## 5. 工作量估算

| 项 | 触及面 | 人周（1 人熟悉本仓） |
|---|---|---|
| Store → Context 类型与工厂 | `store/index.js`、`method.js` destroy 重置 | 1 |
| 74 个 `import Store` 改为 `getStore()` / 注入 | 几乎全部 controller/global/function | 3–5 |
| `api.js` 109 个方法加 instance 解析 | `frontend/src/global/api.js` | 1–2 |
| 公式模块去全局 | `formula.js`、`functionlist.js` 的 `window.luckysheet_function` | 2–3 |
| DOM id 前缀化 | `createdom.js` + 事件选择器 | 3–4 |
| 协同 `server` 按 gridKey 多连接 | `server.js` | 1–2 |
| 插件/图表/透视单例 | chart plugin、`pivotTable.js` | 2 |
| 回归：双容器并存、#529、#504 | `tests/regression/` | 1–2 |
| **合计（方案 A+B 可发布）** | | **14–21 人周** |
| 方案 A 仅 focused 切换（仍单 DOM） | | **6–8 人周** |
| 方案 C iframe | | **1–2 人周**（产品层） |

风险不在「改一个单例」，而在 **公式 eval 全局 + 固定 DOM id + 模块级 controller 单例** 三处必须一起拆。只改 `Store` 会留下隐性串扰。

## 6. API 兼容方案

| 现有调用 | 建议行为 |
|---|---|
| `luckysheet.create(options)` | 保持；若已有实例：默认 **destroy 再建**（旧行为），或 `options.multi=true` 时创建第二实例 |
| `luckysheet.destroy()` | 销毁 focused；`destroy(id)` 销毁指定 |
| `setCellValue` / `setRangeValue` / `toJson` / `undo` | 默认 focused；`options.order` 仍表示 sheet 索引，**不**复用为 instanceId |
| 新可选 `options.instanceId` | 显式指定目标簿 |
| `luckysheet.flowdata()` | 继续返回 focused 的 `flowdata` |
| 协同 `gridKey` | 必须 1:1 绑定 instance，禁止两实例共享同一 `server` 对象 |
| `jfundo` / `jfredo` 命名 | **禁止改名**（计划硬约束）；每实例各有一对栈 |

**不兼容但必须文档化的点**：

- 两个实例不能共享 `calcChain` 跨簿引用（LS 本就没有跨 workbook 公式）。
- 固定 id 未改完前，禁止同 document 双挂载。
- `changLang` 今日会 `toJson()` + `create()`，会毁掉第二个实例。

## 7. 与 P0 缺陷的关系

- [#529](https://github.com/dream-num/Luckysheet/issues/529)：`destroy` 用 `defaultStore` 覆盖全部字段（`method.js:448-453`），再 `create` 时 `execF`（`sheetmanage.js:940-977`）与 `restoreselect`（`sheetmanage.js:1146-1166`）重复写 `scrollTop`。多实例改造能根治，但 **P1 应先修单实例生命周期**，不必等待本评估落地。
- [#504](https://github.com/dream-num/Luckysheet/issues/504)：focus 滚顶与单例正交；`focus({preventScroll:true})` 可独立做。
- 本 Phase **不实施** Store 拆分。

## 8. 结论

把 Store 单例改成实例上下文 **可行**，语义上应对齐 UV 的 `createUnit` / `disposeUnit` / `focusUnit`，而不是引入 Command 体系。要达到「同一页两个可交互表格」，必须同时做 **Context + DOM 前缀 + 公式/controller 去全局**，量级约 **14–21 人周**。对外 API 应以 **focused 实例保持旧签名** 为兼容底线，新增 `instanceId` 而不是改 `order` 含义。当前 fork 的 P0/P1 不应被本改造阻塞。
