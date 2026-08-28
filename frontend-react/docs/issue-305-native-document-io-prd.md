# Issue #305 NativeDocumentIO PRD

## 目标

把导入/导出从 XLSX 中心转换器收口为原生文档协议 I/O。文件先由唯一 `NativeFormatDetector` 识别，再交给对应原生 codec；codec 产出 `NativeDocumentArtifact` 与 canonical Workbook projection。Save 保持原 family/variant/profile，Save As 才允许显式转换。

```text
bytes -> NativeFormatDetector -> NativeDocumentCodecRegistry
      -> NativeDocumentArtifact -> WorkbookProjection -> Model/History/UI
```

禁止 BIFF/XLSB/ODS/SJS 先转 XLSX；禁止 API 别名、桥接层、双写、legacy fallback 和 UI 假成功。未拥有但合法的 native parts/records/streams 保存在 artifact，不能进入 WorkbookSnapshot。

## 格式与所有权

| family | detector | 本轮 codec | Save 语义 |
|---|---|---|---|
| OOXML | OPC ZIP + workbook content type/profile | 现有 OOXML codec，改挂 NativeDocumentIO | 保持 Strict/Transitional 与 xlsx/xlsm/xltx/xltm/xlam |
| text | BOM/encoding/dialect sniff | CSV/TXT/PRN/DIF/SYLK 原生 dialect codec | 活动工作表语义，保留编码/BOM/分隔符/换行 |
| xmlss | SpreadsheetML 2003 namespace/root | XML Spreadsheet 2003 codec | 原生 XMLSS，不转 OOXML |
| ods | ODF ZIP + `content.xml` mimetype | ODF package codec | 保持 ODS package |
| sjs | SpreadJS ZIP markers | SpreadJS SJS package codec | 保持 SJS JSON parts/unknown fields |
| ssjson | SpreadJS JSON schema markers | SSJSON schema codec | 保持 SSJSON |
| xlsb/biff | CFB/BIFF12 magic | 独立 binary codec boundary | 不借道 OOXML；CFB 目录、BIFF/BIFF12 基本单元格记录和原始未知流/部件由 binary graph 原生读写，公式表达式和未拥有的结构保持 preserved-only 或 typed fail-close |

每个 artifact 记录 `format`, `sourceBytes`, `checksum`, `nativeGraph`, `ownership`, `compatibility`。`WorkbookSnapshot` 只保存 canonical projection，不保存 opaque bytes。

## Save / Save As

- Save 使用 artifact.format 对应 codec，只重写 editable-owned 区域，未拥有区域保持原始 bytes。
- Save As 明确传入 target format，生成 conversion report；任何丢失都逐项记录 `converted/lost/blocked`。
- 后缀与内容冲突时内容优先，并产生 diagnostic。
- 修改会破坏数字签名覆盖范围时 fail-close，不输出伪有效签名。

## 资源与失败边界

所有解析/序列化在 `NativeDocumentWorkerPort` 执行并受统一资源上限约束（输入 bytes、entries、inflated bytes、CFB streams、records、XML depth、cells）。错误必须携带 code、format、对象位置、原因和恢复动作；不允许空快照或静默降级。

## 验收

- 统一 registry 直接注册各 family codec；生产调用方不再依赖 `Xlsx*` 顶层 API。
- CSV/TXT/PRN/DIF/SYLK、XMLSS、ODS、SJS、SSJSON 完成真实 native round-trip。
- OOXML 保持现有未知 part/relationship/VBA/custom XML 保真能力。
- XLSB/BIFF 走独立 record/CFB graph，不经过 OOXML；不能安全编辑的 record 明确保留或拒绝。
- Save/Save As、format detector、dialect、artifact ownership、unknown preservation、worker cancellation 均有成功/拒绝测试。
- 未修改的 OOXML、文本、ODS、SJS、SSJSON 文档通过 `sourceSnapshotHash` 证明后直接返回原始 source bytes；修改后只进入对应 native writer。
- 原生 Excel/SpreadJS/Office corpus 若环境没有真实 producer 或 executable，验收项标记 `Blocked`，不伪称通过。

## 回滚与迁移

这是一次性 canonical I/O contract 迁移。旧 `Xlsx*` 运行时入口在同一 PR 删除，调用方改用 NativeDocumentIO；旧 snapshot 在 hydrate 边界只接受显式离线 migrator，运行时不读取 legacy 字段。回滚为单一 PR revert。

## 本轮明确边界

BIFF5/BIFF8/CFB 与 BIFF12/XLSB 已具备独立 record/CFB/package graph：读取工作簿目录、工作表边界、数字/字符串/布尔/空白单元格、共享字符串和 BIFF12 行上下文；Save 只改写这些被明确建模的单元格，未知 record、stream、package part 原样保留。公式 token、宏执行、图表/控件、复杂 row/column 属性和其他未拥有结构不会被猜测或转换，变更触及这些结构时以 typed `NATIVE_DOCUMENT_UNSUPPORTED` 拒绝。

当前仓库没有 Excel/LibreOffice 产出的真实 BIFF5/BIFF8/XLSB corpus，也没有桌面 Excel executable，因此真实 producer interoperability 验收继续标记 `Blocked`；本地 synthetic corpus、解析/重写 round-trip、资源限制和 worker 链路可作为自动化证据，但不替代该外部验收。
