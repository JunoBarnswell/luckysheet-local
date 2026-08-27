# CellEditDomain 性能证据

## 测试环境

- CPU：12th Gen Intel(R) Core(TM) i7-1260P，12 cores / 16 logical processors
- 内存：15.7 GiB
- Node.js：v24.15.0
- 命令：`npx tsx --test packages/spreadsheet-app/src/cell-edit/performance.test.ts`
- 日期：2026-08-28

## 当前结果

| 场景 | p50 | p95 | p99 | 其他 |
|---|---:|---:|---:|---|
| 1,000 次同步 draft reducer | 0.0096 ms | 0.0217 ms | 0.0805 ms | max 1.1248 ms；history 66,000 bytes |
| 32 KiB 长文本插入 | 0.3071 ms | 0.4055 ms | 0.4785 ms | max 0.4785 ms |
| 100k 同列 AutoComplete 查询 | 0.0006 ms | 0.0036 ms | 0.0070 ms | max 0.1270 ms |
| 100k 同列 AutoComplete 索引构建 | — | — | — | 85.231 ms；heap delta 33.565 MiB |

所有已执行预算均通过：

- 普通 reducer：p95 ≤ 1 ms，p99 ≤ 2 ms。
- 32 KiB reducer：p95 ≤ 4 ms，p99 ≤ 8 ms。
- 已构建索引查询：p95 ≤ 4 ms，p99 ≤ 8 ms。
- 同列索引按 chunk 构建，可取消；缓存最多保留 8 列。

## 尚待浏览器验收

以下项目必须在 production build 与内置浏览器中记录，当前不能由 Node benchmark 代替：

- input event → editor/caret paint 的 p50/p95/p99。
- IME composition update → paint。
- Point Mode 连续拖动 FPS、long task 与 Canvas 基础层 redraw 次数。
- 100k×20 工作簿单格 commit 的端到端时间。
- 连续 1,000 次 begin/cancel 后的 retained heap。
