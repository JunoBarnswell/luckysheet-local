# 共享内核整合交接与验收记录

本文件记录 `codex/web-excel-shared-kernel` 相对基线 `4fa80a2a879e6f5e0ac626f2fa45b6dd983a2cd4` 的最终整合判断。批准范围仍以 [Web Excel 共享内核重构与体验增强 PRD](web-excel-shared-kernel-prd.md) 为准；64 项逐项状态、证据、删除路径和阻塞项只在 [机器可读台账](web-excel-shared-kernel-status.json) 中维护。

## 最终产品与架构判断

工作簿产品链已经收敛为：

`UI intent → typed command → Rust WASM preview/geometry/formula → Java OIDC/ACL/transaction → Rust native authority → immutable pages/history/outbox/artifact → committed revision → browser page replica`

以下边界均为单一所有者：

- `PaneMap` 拥有表头、冻结窗格、命中、选择、编辑和提交坐标；四个冻结区域互斥。
- Rust `kernel/core` 使用固定 1024×32 列式页、manifest、内容哈希和有界缓存；隐藏行列属于 visibility projection，不改变 canonical cell read。
- Rust `kernel/commands` 拥有 mutation validation、结构变换、保护、撤销与 restore；一次事务只发布一个 revision/history record，失败不发布部分页或 metadata。
- Rust `kernel/formula` 拥有 parser、引用 AST、依赖图、spill、函数目录、数值/日期语义和服务上下文；native/WASM 使用同一实现。
- Rust `kernel/analytics` 拥有 typed column source、filter ownership/visibility、query、pivot、取消、预算、稀疏 viewport 和分页 drilldown。
- Rust `kernel/native-document` 拥有 OOXML package graph、导入、重写和保真边界；未知 part/node/extension/macro 被保留，数字签名和不安全转换显式拒绝。
- Java 只拥有身份、ACL、事务、三方言持久化、连接器调度和 outbox；旧 Java mutation/snapshot reducer 已删除。
- React 只拥有交互、视觉和 committed page replica；local-only、mirrored、offline commit、浏览器 Pivot worker 和 TypeScript OOXML 执行链已删除。
- 云端 ack 是唯一 saved 事实源。未配置 OIDC/backend 时，Hub 和 workbook route 返回可观察错误，不创建本地工作簿、不显示假成功。

## 已删除的旧设计

本次 clean-break 删除了 Java `mutation/**`、`WorkbookSnapshotValidator`、全量 checkpoint entity；删除浏览器 `offline-queue`、OT rebase、collaborative undo、本地 asset migration/native document store；删除 TypeScript Pivot source index、block source、worker/task protocol/dense projection；删除 TypeScript OOXML archive/import/export/binary codec。边界检查禁止重新引入这些运行时所有者、legacy manifest、双写或 fallback reader。

数据库 V7–V12 为显式迁移边界。V10 只归档旧 snapshot，V12 验证 canonical cutover；运行时只接受 manifest 11、host protocol 1，不读旧字段。应用与数据迁移必须整体回滚，不能单独恢复任一旧消费者。

## 当前整合验证

2026-09-08 使用当前分支源码完成：

| 门禁 | 结果 | 覆盖 |
|---|---|---|
| `cargo test --workspace` | PASS | analytics 23、commands 12、core 4、formula library 21、formula migration 11、formula runtime 16、geometry 5、host 7、native-document library 7 + integration 5；doc tests 通过 |
| `npm run build:kernel` + `npm run verify:kernel` | PASS | 同一 Rust 源生成 native host、WASM 与 Vite 产物；manifest/protocol/hash 一致 |
| `npm run typecheck` | PASS | 全前端 TypeScript |
| `npm run check:boundaries` | PASS | contracts、mutation registry、artifact provenance、421 项 acceptance matrix、formula parity |
| `npm run test:unit` | PASS，242/242 | canonical model/command/render/formula/persistence/permissions/query/pivot/native transport 等 |
| `npm run test:performance` | PASS | 百万行 Rust 稀疏 Pivot + 分页 drilldown 约 3.06 秒；native 输入预算拒绝；45 项 render/edit 性能与几何用例 |
| backend `mvn test` | PASS，80/80 | H2 migration、ACL、native host、pages/history、query proof、artifact/import/export transaction |
| H2 authenticated SQL smoke | PASS | real server create=201、manifest=200、native cell commit=201、revision 1 page publication、anonymous=401 |
| `npm run test:e2e` | PASS，6 passed / 102 skipped | kernel artifact、permission/elastic contracts、Hub、未知 workbook fail-close、1672×941 Hub geometry；跳过项不计通过 |

真实 in-app browser 在 `http://127.0.0.1:4180/workbooks` 检查了 1024、1366、1920 宽度：Hub 有有效内容，无框架错误遮罩，console 无 error/warning；未配置云身份时创建和未知工作簿均显式 fail-close。该结果只证明 disconnected Hub，不替代 connected editor 验收。

性能入口已从被删除的旧 TypeScript Pivot/OOXML 测试迁到当前所有者：Rust 百万行分析、Rust native-document 预算，以及前端 PaneMap/render/cell-edit 性能。当前证据不覆盖 PRD 中所有 P95、浏览器峰值、真实连接器或 2000 万格导入预算。

## 64 项状态

当前机器台账为 64 项：21 `verified`、22 `implemented`、21 `blocked`、0 `not-started`。

- `verified`：本地命令直接覆盖该项完整验收子句。
- `implemented`：生产链和拒绝路径已实现，尚缺一项或多项完整验收证据。
- `blocked`：生产链已实现，但批准验收依赖本机不存在的真实身份、服务、数据源、桌面产品或 PR 托管环境。每项阻塞条件写在机器台账中。

不得将 E2E skip、synthetic OOXML self-roundtrip、H2 单方言或 disconnected Hub 计作对应外部验收通过。

## 剩余外部验收

| 范围 | Blocked 条件 | 完成证据 |
|---|---|---|
| connected editor、保存、协作、UI | 缺 `E2E_OIDC_AUTHORITY`、`E2E_OIDC_CLIENT_ID`、真实用户文件及运行中的 backend | 102 个连接态 E2E 在真实身份下执行；console/network、刷新恢复、断线拒写、双客户端、键盘/IME/屏幕阅读器记录 |
| PostgreSQL/MySQL/Redis | 本机未提供三个真实服务 | PR workflow 的 fresh/upgrade/SQL smoke、事务、outbox、多实例重连结果 |
| live connector/query | 未配置真实 connector 与高基数数据源 | bounded heap、首块延迟、cancel/stale proof、结果分页指标 |
| 百万/2000 万真实工作簿 | 无连接态 dense/sparse/multisheet 语料 | PRD P95、browser/native peak memory、取消无残留、保存重开 |
| Excel/WPS | 无桌面 Excel/WPS 和 real-producer XLSX/XLSM corpus | producer hash、导入→编辑→导出→Excel/WPS reopen 差异；宏、未知扩展和签名边界 |
| GitHub 交付 | PR 尚待创建 | PR 描述、远端 checks、契约迁移、Blocked、整体 rollback 记录 |

## PR 交付与回滚

PR 必须以本分支为 head、`main` 为 base。描述需列出 shared-kernel ownership、公开 DTO/协议变化、V7–V12 迁移、删除的 legacy 设计、上述验证、所有 Blocked 条件和检查链接。

回滚单位是本 PR 加数据库迁移边界：先停止写入并备份 canonical pages/history/artifacts，再回滚应用与数据库到基线对应版本。不得只回滚前端、Java 或 Rust，也不得恢复 runtime legacy reader、双写或 snapshot fallback。已导入的未知 OOXML/macro 原包须与数据库备份一起保留。
