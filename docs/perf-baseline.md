# LuckySheet 性能基线（Phase 2）

> 对照对象：Phase 0 稠密 `flowdata`（`buildGridData` 预分配 `row × column` 的 `null` 矩阵，`luckysheetDrawMain` 全量 `clearRect`）。  
> **浏览器 FPS / `performance.memory` 本机未开浏览器采集，下表对应列保持「未测」，不编造。**  
> Node 侧契约与堆增量已于 2026-08-23 实测，见第 4、5 节。

## 1. 目标场景

| 场景 ID | 工作簿 | 说明 |
|---|---|---|
| S1 | 10 万稀疏格 | `row=100000`, `column=20`，`celldata` 仅 100000 个 `{r,c,v}`，均匀落在前 20 列；无合并、无条件格式 |
| S2 | S1 + 合并跨越 | 在可见区边界放 8 个 `rs/cs ≥ 3` 的合并格，验证 dirty 区被扩展 |
| S3 | 热写 | 在 S1 上对当前可见区连续 `setCellValue` 200 次 |

默认窗口：画布约 `1280×720` CSS 像素，`devicePixelRatio=1`，默认行高 19、列宽 73。

## 2. 指标定义


| 指标 | 符号 | 定义 | 采集点 |
|---|---|---|---|
| 初始化堆增量 | `Δheap_init` | `create()` 完成后 − 创建前的 `performance.memory.usedJSHeapSize` | `luckysheet.create` 返回且首帧 `luckysheetrefreshgrid` 结束 |
| 占用格数 | `occupied` | 稀疏层实际存下的非空单元格数，期望 = `celldata.length` | `occupiedCellCount(Store.flowdata)` 或 `getGridData(file.data).length` |
| 逻辑维度 | `logicalRC` | `flowdata.length × flowdata[0].length`，应等于 `file.row × file.column`，**不得**为此预分配 `null` | 初始化后立刻读 |
| 导出体积 | `celldataN` | `toJson().data[i].celldata.length`，应等于占用格；导出对象**不得**再带稠密 `data[][]` | `luckysheet.toJson()` |
| 滚动 FPS | `fps_scroll` | 连续滚动 3s 内 `requestAnimationFrame` 次数 / 3 | 自定义 rAF 计数器 |
| 滚动帧时延 | `p95_frame_ms` | 同上 3s 内相邻 rAF 间隔的 P95 | 同一计数器 |
| 局部重绘面积比 | `dirty_area_ratio` | 一次 `setCellValue` 后 `clearRect` 宽高乘积 / 画布 CSS 面积 | 在 `luckysheetDrawMain` 打日志或 Performance mark |
| 热写耗时 | `t_setcell_200` | 200 次可见区写入（含公式 dirty flush）的 `performance.now()` 差 | 脚本计时 |

**通过标准（相对 Phase 0，而非绝对分数）**

- `occupied` 必须等于稀疏格数；禁止出现 `new Array(100000)` 且逐行 `new Array(column).fill(null)` 的物化。
- `Δheap_init(S1)` 应明显低于 Phase 0 同场景（Phase 0 约为 100000×20 个槽位 + 包装对象）。
- 小范围写入时 `dirty_area_ratio` 应远小于 1；合并格跨越 dirty 边界时重绘区必须覆盖整个 merge。
- `fps_scroll` / `p95_frame_ms` 只与**同机器、同浏览器、同 Phase 0 录屏**对比，不跨设备宣布“提升了多少”。

## 3. 测量方法（本机浏览器）

本仓库 Node 环境只能验证稀疏契约（见下节），**不能**代替浏览器 FPS/堆内存。

1. 在 `frontend` 执行 `npm run build`，用静态页打开 `frontend/tests/regression/sparse-perf.html`（或任意挂了 `dist/luckysheet.umd.js` 的页面）。
2. 打开 Chromium DevTools → Memory / Performance。不要用任务管理器的整页内存冒充 JS 堆。
3. 初始化脚本（可直接贴到页面）：

```javascript
const celldata = [];
for (let i = 0; i < 100000; i++) {
  celldata.push({ r: i, c: i % 20, v: { v: i, m: String(i) } });
}
const before = performance.memory && performance.memory.usedJSHeapSize;
luckysheet.create({
  container: "luckysheet",
  data: [{ name: "S1", row: 100000, column: 20, celldata: celldata }]
});
requestAnimationFrame(() => {
  const after = performance.memory && performance.memory.usedJSHeapSize;
  const file = luckysheet.getluckysheetfile()[0];
  console.log({
    deltaHeap: after && before != null ? after - before : "performance.memory 不可用",
    logicalRC: [file.data.length, file.data[0] && file.data[0].length],
    celldataN: luckysheet.toJson().data[0].celldata.length
  });
});
```

4. **滚动 FPS**：Performance 面板录 3s 平滑拖动纵向滚动条；或用 rAF：

```javascript
let n = 0, last = performance.now(), gaps = [];
function tick(now) {
  n++;
  gaps.push(now - last);
  last = now;
  if (n < 180) requestAnimationFrame(tick);
  else {
    gaps.sort((a, b) => a - b);
    console.log({ fps: n / 3, p95: gaps[Math.floor(gaps.length * 0.95)] });
  }
}
requestAnimationFrame(tick);
```

5. **dirty 面积**：对可见区内一格 `luckysheet.setCellValue(r, c, 1)`，在 Performance 里看 `luckysheetDrawMain` 是否只 `clearRect` 局部；S2 再确认合并主格被一起清掉。
6. 把测得数字记入下表，**不要回填未测值**。

## 4. 记录表


| 场景 | 环境 / 机器 | Δheap_init | occupied | celldataN | fps_scroll | p95_frame_ms | dirty_area_ratio | 备注 |
|---|---|---|---|---|---|---|---|---|
| Phase 0 S1 | 浏览器未测 | — | — | — | — | — | 1（全画布 clear） | 稠密基线，无对照录屏 |
| Phase 2 S1 | 浏览器未测 | — | — | — | — | — | — | 打开 `sparse-perf.html` 后补 |
| Phase 2 S1 Node | Node v24.15.0 / win32 x64 / 2026-08-23 | heapUsed Δ **35.54 MB**（`process.memoryUsage`，**不是** `performance.memory`） | **100000** | **100000** | 未测（无 rAF） | 未测 | 未测（无 canvas） | `measure-sparse-heap.cjs`：`createSparseGridFromCelldata` 100000 格，`logicalRC=100000×20`，`storeRows=100000`，`create_ms=42`，rss Δ 56.9 MB |
| Phase 2 S2 | 浏览器未测 | — | — | — | — | — | ≥ merge 并集 | #7394 merge-aware |
| Phase 2 S3 | 浏览器未测 | — | — | — | — | — | — | 热写 |

Node 数字只说明稀疏层**不会**为 100000×20 预分配空槽。不能用来宣称浏览器滚动 FPS 提升。

浏览器补测：`frontend` 执行 `npm run build` 后，用静态服务打开 `frontend/tests/regression/sparse-perf.html`（不要 `file://`），按第 3 节操作，把 Δheap / FPS 填回本表 Phase 2 S1 浏览器行。

## 5. 仓库内可自动验收（无需浏览器）

在仓库根或 `frontend` 执行：

```bash
node frontend/tests/regression/run-sparse-grid.cjs
node frontend/tests/regression/measure-sparse-heap.cjs
```

`run-sparse-grid.cjs` 直接加载 `sparseGrid.js` / `dirtyRect.js`。2026-08-23 结果：**全部 PASS**（exit 0），其中包括：

- 逻辑维度 `100000 × 20`，3 个占用格时 `occupied=3`，backing `store.size=2`（只存有格的行）；
- `data[r][c]` 读写外观；空行不分配行数组；
- `toCelldata` 导出 3 条 `{r,c,v}`；
- `cloneSheetData` 隔离引用且不稠密化；
- `ensureSparseSize` 只改逻辑维度（扩到 100011×26 后 occupied 仍为 4）；
- dirtyRect 合并扩展为 `0..2 × 0..2`；小滚动增量 dirty 从行 21 起；大跳转回退全量。

`measure-sparse-heap.cjs`（S1 同构、无画布）：`occupied=100000`，`celldataN=100000`，`logicalRC=100000×20`，构建 42 ms，heapUsed Δ 35.54 MB。数字见第 4 节。
