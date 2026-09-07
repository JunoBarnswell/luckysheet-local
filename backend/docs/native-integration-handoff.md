# Native backend integration handoff

此批实现已落盘，未执行编译或测试；按用户要求由 root 串行完成后统一验证。不能把本文件当作完成/通过声明。

## 已落盘

- `WorkbookOperationService` 删除 nullable kernel、Java reducer/replay、旧 snapshot DTO/reader；统一 server ACL → DB lock/idempotency → native command.prepare → page.load → command → manifest/pages/history/operation/outbox 同事务发布。rollback afterCompletion 丢弃 native staged state。
- `KernelPersistenceService` 按 unit/checksum 存 immutable page，按 unit/revision 存 immutable manifest。所有读写校验 checksum/length/descriptor。history 保存页引用与 metadata before，readChangeSet 从同一 refs 重建，不写全工作簿 cells checkpoint。
- `WorkbookCatalogService` 原生 create/copy/import/export；任意 artifact PUT 已删，artifact 为 durable task storage 文件引用，revision/hash/length/format 绑定。导入/导出通过 pages.json `[ {descriptor,fileHandle} ]`，每项为 raw LSPG page file。
- `NativeDocumentTaskService` 持久化 chunk upload、offset/hash/length、取消、同事务 publication/idempotence；1 GiB 文件/8 MiB chunk/64 KiB buffer。final cancellation lock 在 task 与 workbook 共用事务裁决。
- `QueryExecutionProofService` 起始 revision/token、READY/result hash、actor/expiry/cancel/once consume，同 native commit 事务；native query.load 收到 server `trustedResult`。
- 删除整个 Java mutation 包、WorkbookSnapshotValidator、WorkbookSnapshotResponse、旧 JPA snapshot/checkpoint runtime fields/实体/API。
- V7 pages/manifests、V8 history、V9 query proof + archive旧artifact BLOB/new durable paths、V10 snapshot/archive/proof、V11 import tasks、V12 proof gate/drop旧snapshot，三方言齐备。V12 既有数据库需要 offline importer；没有伪造 migration proof。
- 新真实 native/H2 success/reject integration tests、page persistence、query proof、upload tasks、cutover gates tests。未运行。

## 对 frontend 的确切 API

- POST `/api/workbooks` body `{unitId,name,sheets?,spaceId?,folderId?,source?}`。source仅native，拒client snapshot。返回 `{unitId,revision,manifest,checksum}`。
- GET `/{unitId}/manifest?revision=`（可省revision）→ raw v11 manifest。
- GET `/{unitId}/pages/{sheetId}/{pageRow}/{pageColumn}?revision=` → raw PagePayload。
- POST `/{unitId}/kernel-operations` body `{operationId,baseRevision,clientSequence,mutations:[{id,sheetId,params}],intent?}` → raw ChangeSet。
- POST `/{unitId}/operations` OperationEnvelope → `{operation,changeSet}`，同一 service core。
- POST `/{unitId}/checkpoints` → `{workbook:{unitId,revision,manifest,checksum},created:false}`，每个commit已存manifest checkpoint。
- POST `/{unitId}/native-document-artifact` `{revision,fileName,format}` → artifact metadata（包含revision）。GET流文件，stale拒绝。
- POST `/api/workbook-imports/tasks` → task；PUT `/tasks/{id}/chunks?offset=`；POST `/tasks/{id}/commit` → task `{result?:WorkbookImportResponse}`；GET/DELETE task。原multipart只接file/name/space/folder。

## root 必須串行补齐的接口与风险

> Root 整合进度（2026-09-07）：以下第 1–4 项已按 canonical host contract 实现并通过针对性 Rust/native/H2 测试。`command.prepare` 已存在；create 省略 sheets 由 native 创建默认工作表；undo 使用服务端已验证 HistoryRecord 与 descriptor 引用并在冲突时 fail-close；restore/copy 发布连续 revision 或独立 revision-zero identity。第 5 项 source artifact 复制已接入原包保真导出并验证当前单元格写入，但真实 Excel/WPS producer 保存仍 Blocked。第 6–10 项继续按原说明整改。历史交接文字保留，用于说明原始断点。

1. 当前 native host 尚未确认实现 `command.prepare`。Java发送与command完全相同 envelope，期望 `{pages:[{sheetId,pageRow,pageColumn}]}`。必须真实 native决定所需页/公式依赖，不靠吞DATA_PAGE重试。空白页不应在需要hydrate的目录中。
2. Native `create` 当前要求 sheets，而HTTP契约允许省略；默认sheet必须在native create实现。不要Java重建canonical默认metadata。
3. Native `history.undo` 尚需服务端 proven before page 装载：Java现传 `params.history`（来自DB）。不能把所有before pages塞16MiB control frame；可用host staging/任务文件。公开client inverses不参与undo，committed log记录history.undo+targetOperationId。
4. Native `restore` 当前Java期望 `{unitId,baseRevision,operationId,accessRole:'owner',targetManifest}` → ChangeSet。Native `copy` 期望 `{sourceUnitId,sourceRevision,targetUnitId,name}` → `{manifest,...}`。Java copyPages 会复制可信blob并保存revision0 checkpoint，不要再publish重复checkpoint。
5. 当前 Catalog.copy 尚未转移 source artifact provenance；复制xlsm/unknown OPC应在native copy/export路径保留原包，不能因为页面复制成功就算完整copy。请补证据并保持viewer复制权限。
6. `native_document` owner契约：import `{unitId,name,fileHandle,outputDirectory}` → `{manifest,pagesManifestFile,artifact:{revision,checksum,byteLength,format,codecRevision}}`；export `{unitId,revision,fileHandle,format,sourceFileHandle?,sourceChecksum?,sourceRevision?,pagesDirectory,pagesManifestFile}` → `{artifact}`。Java `fileHandle` 是受控临时目录内尚不存在的输出文件；要确认native blank-workbook export无source文件也能生成合法OPC。native_io最后消息称import revision1/artifact.identity嵌套，这与已定revision0/direct artifact冲突；已要求其统一，root必须检查最终代码。
7. B06尚未实现：`QueryExecutionService` Java全量算子/结果materialization仍存在，等待 `/root/analytics_integration/columnar_query` 外部connector批次ABI。B07 proof目前 `{columns,rows}`；需要随B06变为结果页manifest。不能仅提高限额或把入口改unsupported就宣称完成。
8. Explicit offline DB importer尚未实现，V12存量cutover会Blocked。Rust v10 artifact转换不等于checkpoint/operation历史证明。详见native-kernel-cutover.md。
9. GeneratedWorkbookContract.java仍是旧生成文件（v10常量+mutation capability map）；已无semantic consumers，仅CreateWorkbookRequest用MAX_WORKBOOK_NAME_LENGTH。root须随contracts生成器整体退役该旧生成规则；不要重建Java authority。
10. Catalog UserState 中旧autoSave/autoSync/offlineCache/defaultCreateLocation字段仍存在，云唯一持久化前端若已删须统一公开DTO/DB用户偏好语义。不是已迁移宣称。

## 验证顺序

先完整补完上述host/契约，再统一 `cargo build -p workbook-kernel-host`，Java21/Maven测试（native tests默认找repo/target/debug/workbook-kernel-host.exe，缺文件硬失败）。检查 `WorkbookOperationServiceTest`、`WorkbookCatalogServiceTest`、`KernelPersistenceServiceTest`、`QueryExecutionProofServiceTest`、`NativeDocumentTaskServiceTest`、`NativeKernelCutoverMigrationTest` 及原H2授权测试。补真实DB并发取消、Postgres/MySQL fresh+upgrade matrix、真实OIDC frontend页/commit/network/console、nativeExcel真实corpus。没有部署或Git commit。
