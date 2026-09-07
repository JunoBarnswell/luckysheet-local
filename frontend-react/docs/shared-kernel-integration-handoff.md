# 共享内核整合输入与剩余整改

本文件汇总已有智能体报告及其落盘交接文档，作为 PRD 的执行输入。报告中的测试结果为该智能体当时报告的局部证据，不代表当前整合源码已重新验证，不提升整项验收状态。未经重新核对的接口风险保留为待确认，不推断为已修复。

## 已有交接信息

| 来源 | 已报告实现 | 已报告验证 | 尚未闭环 |
|---|---|---|---|
| analytics_integration / kernel/analytics/README.md | 列式源索引、revision cache、typed filter、稀疏 Pivot、12 种聚合、多级汇总、TopN、ShowAs、viewport、分页 drilldown、hash join、分组查询、取消与预算 | `cargo test -p kernel-analytics --lib` 23/23，含百万行用例，报告耗时 3.08 秒；WASM check 通过 | session source-register、主线程同步 WASM、TS engine 和 dense 结果转换仍存在；connector 批次、proof、spill sort、计算字段接线、collation、布局、筛选视觉与真实验收待完成 |
| backend/docs/native-integration-handoff.md | ACL/native command/pages/history/outbox 事务；immutable pages/manifest；文件分块任务；query proof；删除 Java mutation/snapshot runtime；V7–V12 三方言迁移和测试代码 | 明确未编译、未运行测试 | host prepare/default create/undo/restore/copy/file ABI 需逐项复核；查询全量物化、offline DB importer、旧生成契约和用户偏好待整改 |
| kernel/formula/README.md | 统一 parser/reference AST、依赖和 dirty closure、spill、服务化取消/可见性/外部执行边界、检查与 trace API、函数能力目录 | 报告 21 个 library + 16 个 runtime integration tests 通过 | host metadata/services 未闭环；共享公式模板、聚合索引、格式 locale、migration parity、分页 spill、WASM/native 差分、百万行和真实 Excel 待完成 |
| kernel/README.md / docs/verification/README.md | protocol 1、manifest 11、1024×32 页和统一 build/verify 入口 | 文档本身不构成运行证据 | 全链消费者、前后端统一验证、浏览器 console/network、真实 producer/reopen 证据仍需完成 |

## 整合 TODO（隶属既有 PRD，不扩充产品范围）

| TODO / PRD | 当前状态及问题 | 整改方案 | 验收边界 |
|---|---|---|---|
| I01 / A06、P04、P06 | host analytics.execute 使用 stateless 入口，丢失跨请求源缓存 | 工作簿独占 AnalyticsRuntime；成功 open/close 清理生命周期；revision 校验先于执行 | 跨工作簿隔离、重复计算、旧 revision 拒绝、取消恢复；不能替代前端任务验收 |
| I02 / P01–P06、B06–B07 | frontend source-register 和 dense tree 仍与 Rust 新契约并存 | 将 session、任务 transport、稀疏结果消费者及 proof 一起切换，删除旧计算链 | 真实任务取消、分页、首块延迟、无主线程全量计算，旧结果不可发布 |
| I03 / F01–F10、A01 | 公式模块的 services/metadata API 与 host 当前调用未完全对齐 | 接通 worksheet/name/table/context、可见性和取消服务，统一 command 后失效与分页读取 | 原回归案例、失败不发布、跨端差分、真实文件与性能证据 |
| I04 / A03、B01–B05、B08、B10 | Java 交接列出多处未确认的 host ABI | 对照实际 host dispatch 和所有 Java 调用逐项核对 create/prepare/undo/restore/copy/import/export，移除失效契约 | native+H2 成功/拒绝路径、artifact provenance、权限、原子回滚 |
| I05 / B06、B11 | 查询仍全量物化；数据库存量切换 importer 未实现 | connector 批次接入 Rust、结果页 proof；显式离线迁移和证明，运行时不读旧快照 | 内存预算、取消、三数据库 fresh/upgrade，未实现前不得标记完成 |
| I06 / D01–D05 | 现有局部测试无法证明当前整合源码通过 | 完整开发后统一构建和验证，按实际证据更新原 64 项状态并交付 PR | frontend/backend/Rust/WASM/browser；真实 Excel 缺失标 Blocked；PR checks 必须记录 |

整合顺序：先公共 host 生命周期与协议，再公式和分析消费者，再文件/持久化/迁移收口，最后统一验收。保留所有既有修改；不新增智能体。PRD 原有 64 项是完整范围，本表不是缩减交付清单。

## 本轮整合证据

I04/A03 撤销链路整改：`history.undo` 现在是 Rust command 的显式 canonical 分支，不进入普通 mutation registry。服务端只传从 immutable history 读取并重新校验的 record；客户端提交的 inverse mutations 不参与执行。HistoryRecord 同时记录 metadataBefore/metadataAfter，页和 metadata 都必须仍等于目标操作的 after state；不相交页可保留，重叠页或 metadata 后续变化返回 `UNDO_CONFLICT`，revision、manifest、history、operation 均不产生部分提交。撤销结果始终发布 baseRevision+1，before 页通过同工作簿 content-addressed persistence 验证并由客户端按需补读。Rust command 12/12 通过，包含权限、页冲突、非重叠页、metadata 成功与冲突；真实 native/H2 的权威历史、跨主体拒绝、重叠冲突共 3/3 通过，另 persistence 6/6 通过。

I04/B04 复制文件身份整改：自研新建 OOXML 在受 Markup Compatibility 保护的 `urn:react-sheets:workbook:1` 属性中写出 canonical worksheet id，重开时按 namespace URI 解析；外部文件无该扩展时继续采用 OOXML `sheetId`。这删除了只存在于内存的临时映射，使新建→导出→重开→复制保持同一 sheet identity。Rust native document library/integration 共 12/12 通过；真实 native/H2 的复制原包当前值与恢复历史页测试 2/2 通过。真实 Excel/WPS 对该扩展的保存行为仍属于 D04 Blocked 验收，不在此宣称通过。

后端重跑结果（2026-09-07）：`WorkbookCatalogServiceTest` 5/5、`KernelPersistenceServiceTest` 6/6，共 11 项全部通过，使用最新 `cargo build -p kernel-host` 生成的真实 native 进程与 H2。命令为 Maven `-Dtest=WorkbookCatalogServiceTest,KernelPersistenceServiceTest test`，本机附加 `-DargLine=-Djdk.net.unixdomain.tmpdir=D:/code/luckysheet-local/.tools/verification/sockets`。原 TEMP 短路径下 JDK Unix-domain pipe connect 失败；已用真实 HttpClient 初始化确认指定目录可用，不是跳过连接器或 mock HTTP。JSON 断言现比较序列化的完整字段和值，不依赖 Jackson IntNode/LongNode 实现类别。此验证尚不证明历史恢复/原文件复制的完整端到端场景，须继续补齐相应验收，亦不代表 64 项整体通过。

I04/B11 集中后端验证发现：`cargo test -p kernel-host --lib` 当前 5/5 通过且最新 native 构建通过；Maven catalog/persistence 共 11 项运行，1 failure、8 errors，不能标记后端通过。持久化比较的根因是 JSON 整数在内存 LongNode 与重读 IntNode 表示不同，统一比较整数数值且保留其它类型严格相等；不修改测试数据。H2 启动错误为新 V9/V11 的 CLOB 与 JPA LONGVARCHAR 不一致，未发布迁移改为与现有 H2 页表一致的 text。历史引用 publication 同时计入已核验变更页集合，防止后续 history 校验再次误拒。修复后需重跑该组测试。

I04 恢复链路整改：原 Java `restore` 调用无 native 实现，不能通过旧 revision 重新 open 冒充提交。现由 core `restore_manifest` 验证同一工作簿和历史 revision、生成当前 revision+1、页差量与 metadata before；host 校验 owner 并在控制帧预算验证后发布 staged 状态。持久化层对未附带 payload 的变更页核验同工作簿下既有 immutable 内容再发布，历史页不通过大 base64 控制帧传输。成功/越权/外工作簿/stale 测试已补，本轮尚待集中运行。前端缺页补读及撤销提交消费者还需同批收口。带原生数字签名的包在重写导出时一律显式拒绝，不能用相同 revision 绕过签名保真边界。

I04 创建契约核对：HTTP `CreateWorkbookRequest.sheets` 可省略，Java 仅在非 null 时发送，而 native host 原先强制要求该字段，导致正常空白工作簿创建失败。整改为 native 创建入口拥有默认工作表（完整 Excel 行列容量、零已分配数据页）；显式提供非法、null 或空列表时拒绝，不以默认值覆盖错误输入。已有工作簿 identity 冲突继续拒绝。

I04 复制与导出核对：Java 已调用 `copy`，host 尚无入口，现实现显式新身份边界：复制 metadata 和页内容 hash、重置 workbook/page revision 为 0，页数据仍经 Java 验证复制和按需加载。Java 复制原生文件时在同事务中以原包为保真来源、目标 canonical pages 为数据来源重新导出，不能只关联旧 artifact 从而让下载返回旧单元格。导出统一使用 `pagesDirectory/pages.json`，删除 Java 冗余输入 `pagesManifestFile`；native 显式接收并核对 `format`。未经实现的跨格式转换返回 UNSUPPORTED_FEATURE，不能假成功。以上复制/导出本轮实现尚未运行整合验证，原生签名、宏、未知 part 的真实文件验收仍待完成。

I01 已接入 `KernelHost.analytics`，由每个工作簿独占运行时；成功 open 和 close 移除缓存，外层 revision 校验先于分析执行。`cargo test -p kernel-host --lib analytics_tests`：2 passed、0 failed，覆盖两个同 revision 工作簿的数据隔离、重复查询、重新打开、关闭后拒绝、stale revision 拒绝、取消后恢复。现有依赖仍有 unused/dead-code warnings。本结果不覆盖浏览器任务 transport、内存/P95 指标或完整 PRD 验收，不据此提升 P04/P06 为 verified。
