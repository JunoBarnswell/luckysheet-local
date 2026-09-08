# v10 持久化到 v11 kernel 的显式切换

本流程只供停机维护时的离线迁移。生产 API 不接受 v10，运行时没有 archive reader、snapshot fallback 或双写。SQL migration 的版本号与 workbook 格式号不同：V10 建立归档，V11 建立 document task 表，V12 执行 v11 workbook 切换门禁。

## 执行顺序

1. 停止所有应用实例、作业和数据库写入方。制作并验证可恢复的数据库备份，以及新文件存储目录的备份。MySQL DDL 不保证全事务回滚，备份必须在任何 migration 之前完成。
2. 使用匹配数据库方言的 Flyway locations，将升级目标设为 `9`，完成旧的 repeatable migration 后再继续。不能让旧的 `rows.permuted` rewrite 在归档之后改变 operation envelope。
3. 将 Flyway 升级目标设为 `10`。检查 workbook、checkpoint、operation 的归档数量、UTF-8 原始内容 SHA-256、revision 与源表一致；原始 OOXML BLOB 已由 V9 保存到 `workbook_native_artifact_migration_archive`。此时不要启动新应用；旧 snapshot 列仍为 NOT NULL，新运行时也不能向它写入。
4. 离线迁移器读取归档。Rust `workbook-kernel-migrate` 是格式转换器；仅输出 artifact 或 manifest 并不构成数据库导入成功，也不能据此设置 `history_verified=true`。
5. 可信离线数据库导入器完成以下证明后，在一个数据库事务中写入 immutable manifests、pages、history、文件 artifact 引用以及 `workbook_v11_migration_proofs`。任一功能或历史记录无法转换时停止，并保留原始归档，不得伪造历史或跳过 revision。
6. 文件 artifact 写入必须先完成原子文件发布，再验证文件存在、原始字节 SHA-256、长度和 metadata 保真。所有原始 OOXML bytes 必须保留。导入器把新引用绑定到当前 workbook revision，不能用新生成的空包替换旧包。
7. 重新核对源表未变化，再执行 Flyway `migrate` 到最新版本。V12 检查全部存量 workbook 的证明与 manifest，并检查 snapshot/checkpoint/operation 和 OOXML 归档；缺少任何证明会触发 `workbook_v11_migration_required_ck`，在 DROP 前失败。
8. 门禁成功后才启动新应用。使用有效身份验证打开、页读取、提交、重新加载、历史恢复与原始文档导出。archive 和 proof 保留，后续删除需要独立的数据保留决策。

空库可以直接 migrate 到最新版本；它没有待证明的旧数据。存量库缺少离线导入器或不能证明全部历史时，升级状态为 **Blocked**。禁止通过手工插入证明、禁用 CHECK、删除待迁移 workbook 或移走 V12 文件绕过门禁。

## 归档契约

`workbook_v10_migration_archive` 主键为 `(source_kind, unit_id, revision)`。`source_kind` 是 `WORKBOOK` 或 `CHECKPOINT`；前者 `revision` 是 workbook 当前 revision，后者为 checkpoint revision。字段包括 `snapshot_revision`、`snapshot_json`、`source_checksum`、`original_checksum`、`source_created_at`、`source_updated_at`、`archived_at`。

`source_checksum` 固定是 `snapshot_json` 原始 UTF-8 字节 SHA-256 的小写十六进制，不对 JSON 重排或重新序列化。`original_checksum` 保留 checkpoint 原有 checksum，必须单独验证；不能把重新计算的 hash 当作旧 checkpoint 本来正确的证据。

`workbook_v10_operation_archive` 保留 `operation_id`、`unit_id`、`revision`、`actor_subject`、`client_sequence`、`base_revision`、`envelope_json`、`source_checksum`、`committed_at`。其 checksum 同样覆盖原始 UTF-8 envelope 字节。归档没有 runtime 外键删除链，不随用户删除 workbook 消失。

## Proof v1 契约

`workbook_v11_migration_proofs` 仅由离线导入器写入，无生产 Entity/Repository：

| 字段 | 要求 |
| --- | --- |
| `unit_id` | 主键，与源 workbook 一致 |
| `proof_version` | 固定 `1` |
| `source_revision` | 原 workbook 当前 revision |
| `source_snapshot_revision` | 原 snapshot revision，不能重置为零 |
| `source_checksum` | 对应 WORKBOOK 归档原始 UTF-8 snapshot 的 SHA-256 |
| `manifest_checksum` | 已写入当前 revision manifest_json 原始 UTF-8 字节的 SHA-256 |
| `checkpoint_count` | 本 workbook 全部 CHECKPOINT 归档数量 |
| `operation_count` | 本 workbook 全部 operation 归档数量 |
| `history_checksum` | 完整历史验证报告原始 UTF-8 字节 SHA-256；报告应外部留存，可重算 |
| `pages_verified` | 仅在全部当前及历史页 payload/hash/长度/descriptor 与 native 校验通过后设 true |
| `history_verified` | 仅在全部归档 checkpoint 校验、完整 revision 连续性、operation replay 与每个 checkpoint 结果逐一证明后设 true |
| `verified_at` | 完成验证的实际时间 |

历史验证报告必须按确定顺序记录每条 operation/checkpoint 的身份、源 hash、源 revision、目标 manifest hash、页证明结果与原始文档 hash；包含实际验证工具版本和失败项。`history_checksum` 不能是任意占位字符串。原始历史缺口、错误 checksum、未知 mutation、无法重建旧 checkpoint，均属于停止条件。SQL 门禁验证证明与源身份的绑定，无法执行 Rust 页解码，也无法读取文件存储；这些必须由独立离线导入器完成，SQL 测试不代替该验收。

## 失败与回滚

- V12 在 CHECK 上失败：旧列/表尚未删除。读取错误、检查缺失 proof 或变化的源数据，修复实际迁移问题。不要清理归档。
- MySQL/H2 失败可能留下 guard 表或 Flyway failed history。确认没有 DROP 执行且备份可恢复后，按 Flyway 官方 repair 流程修复 failed entry 再重试；guard 的创建和清空可重入，但这不代表破坏性 DDL 可回滚。
- V12 已经完成或发生部分 DDL 失败：停止所有写入，恢复完整的升级前数据库和文件存储备份后再部署旧应用。不能仅增加 snapshot 列、回填空 JSON 或切回旧应用假装回滚。
- 三方言 SQL、H2 SQL 门禁测试、真实 PostgreSQL/MySQL migration、离线导入器与 native 页/历史验证是不同验收项；只通过其中一项不能宣称存量迁移已完成。
