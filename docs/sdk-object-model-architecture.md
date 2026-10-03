# SDK 对象模型与多工作簿架构

2026-10-03（UTC）。用户要求：统一 SDK 入口、类似 Aspose 的对象模型、完整 Excel/在线 Excel 能力、多工作簿与跨工作簿函数；先确定架构，一次完成实现批次，再统一验证。本文在代码实现前确定约束；不是宣布所有能力已实现。

## 对象与所有权

`createSpreadsheetSdk` 是唯一组合入口。一个 SDK scope 对应一个认证 subject 和资源生命周期；workbooks 按 unitId 拥有唯一运行实例。Workbook 拥有 WorksheetCollection；Worksheet 由稳定 sheetId 标识；Cell/Range 属于显式 Worksheet，不能从当前选中格、路由、文件名或数组序号取得业务身份。对象保存身份和动作，不缓存另一份单元格真相。

所有程序写入与 Web 操作共享 canonical command → mutation → model → calculation/history/collaboration/persistence。SDK 对象模型拥有参数/地址/工作簿身份规划，内部 runtime port 仅访问已有 canonical cell resolver 与命令事务，不建立模型副本、第二套函数引擎或绕过授权的写入路径。公开 Workbook 不返回 Session、model、commands、API client、token、CSRF 或可变 snapshot。已有 Web Session 暴露仍是待删除的迁移项，新增对象 API 不宣称这一边界已完整完成。

工作簿 open 等待实际 ready，失败/超时/身份退休返回带 operation、对象身份、cause、recovery 的 SdkError。同一 ID 重复打开返回同一对象；共享的打开操作有一个生命周期。close 释放该对象的所有权；被 Web 视图租用的同一 session 继续工作。关闭一个工作簿不关闭其他工作簿；SDK dispose/subject 切换使全部旧对象失效。Web 的 StrictMode probe 复用同一运行实例，最终释放才退还 session 租约。

## 公开 API 本批范围

```ts
const sdk = createSpreadsheetSdk();
await sdk.auth.initialize();
const sourceEntry = await sdk.workbooks.create({ name: 'Source' });
const source = await sdk.workbooks.open(sourceEntry.unitId);
const targetEntry = await sdk.workbooks.create({ name: 'Target' });
const target = await sdk.workbooks.open(targetEntry.unitId);
const sheet = source.worksheets.at(0);
await sheet.cells.get('A1').setValue(10);
await sheet.cells.get('A2').setValue(20);
await target.externalLinks.bind(source, 'Source.xlsx');
await target.worksheets.at(0).cells.get('B1')
  .setFormula(`=SUM('[Source.xlsx]${sheet.name}'!A1:A2)`);
await target.externalLinks.refresh();
const result = await target.worksheets.at(0).cells.get('B1').read();
await target.save();
source.close();
target.close();
await sdk.dispose();
```

setValue 保留字面值（字符串 `=1+2` 不自动变公式）；setFormula 明确声明公式。read 返回 authored value、可见 formula、calculatedValue 和 formulaHidden 标志。Range 是有界地址对象，提供 cell 寻址与惰性遍历；批量原子写、格式/复制/结构变换属于后续真实领域动作，不添加空方法。普通 sparse、data-block、record/view 的读取和写地址使用已有同一 canonical resolver。隐藏行列不改变 cell 数据；range access 与 formula-hidden 权限仍生效。

create 只接受 SDK 的名称、模板与目录 ID；身份和初始文档由同一 SDK planner 生成。Web create/retry 消费方在本批同时迁移，删除其 raw snapshot、unitIdFactory 和 template planner 调用。服务端仍是现有 snapshot create 协议；不是数据版本迁移。顶层对象类型各自独立文件，域的私有 WeakMap 仅存身份/规划闭包，不缓存单元格数据。可重复 ready 调用共享同一 Promise，关闭会拒绝等待者。

## 多工作簿计算

本批复用现有 ExternalLinkBinding、FormulaEngine 与 Java subject-safe external input endpoint。绑定保存 sourceUnitId 与稳定 sheetId，token 只作为公式语法别名。绑定前提交 source 的待发送操作；target 用正常 mutation 保存定义并提交，刷新从服务端按当前 subject 取得 sourceRevision/accessRevision 与可读输入。结果不能直接从另一 Workbook 的内存 snapshot 复制，也不能把来源的待提交数据当成已提交版本。

source close 不移除 target 中持久化的 binding。刷新失败的 denied/broken/unavailable 状态可观察；拒绝访问后清空派生输入，计算报告 #BLOCKED!，不能继续暴露旧授权值。跨 SDK scope 的对象不能直接绑定，防止身份混用。

完整依赖图后续必须支持链式依赖、明确的环/迭代计算策略、图读取期间版本一致性、来源变化传播、关闭/撤销权限清理与重算；本批直接来源刷新不能标记这些项通过。多个工作簿的写事务需 Java/H2 多工作簿授权、锁顺序、共同提交点、失败回滚和历史协议，不以依次 await 多个保存冒充原子事务。

## 全能力对齐清单

| 领域 | 目标对象能力 | 当前验收 |
|---|---|---|
| 入口/生命周期 | create/open/import/save/save-as/export/dispose、多 workbook registry | 本批实现 named create/open/objects/save；全格式 Pending |
| Worksheet/Cell/Range | typed value/formula/rich text/style/merge/clear/fill、范围原子写 | 本批基本 cell 与 range 寻址；其余 Pending |
| 计算 | 完整函数目录、UDF、spill、names、manual/volatile、跨 workbook 图 | 直接 external binding 本批验收；完整图/函数对齐 Pending |
| 结构/引用 | 行列/cell shift、跨 sheet/workbook move、统一引用图 | Pending |
| 表与数据 | Table/filter/sort/subtotal/validation/conditional format、所有 TabularSource | 部分 SDK data actions 已有；完整 Pending |
| 分析 | Pivot/PivotChart/slicer/timeline、what-if、query/linked records | Pending |
| 对象 | chart/image/shape/connector/text/control/barcode/camera/sparkline/OLE | 首图等细项已验；完整 Pending |
| 文档互通 | XLSX/XLSM/XLSB/XLS/ODS/CSV/Text/SJS/SSJSON，读/编辑/写/保留分别验收 | Pending；desktop Excel Blocked |
| 输出 | page setup/print area/titles、preview/PDF/图片/打印 | Pending |
| 在线 | auth/users/roles/range access、presence/history/replay/reconnect | 部分 Pass；完整 Pending，真实 OIDC Blocked |
| 宿主/SDK 发布 | 浏览器无编辑器调用、可注入 Worker/文件/下载/打印端口、Node/Java 文档宿主 | 浏览器本批验收；Node/Java host Pending，当前缺少 browser Worker 明确拒绝 |
| 边界/交付 | Web 仅 sdk/ui-system，完整公开类型/包发布、全部门禁与逐项 UAT | Fail/Pending；不得以本批对象 API 宣称全部完成 |

## 本批实施前 UAT

| ID | 操作和成功断言 | 拒绝/生命周期断言 | 初始状态 |
|---|---|---|---|
| OO-01.a | root open → worksheets → cell/range；按 ID/name/索引寻址；同地址对象稳定；选区变化不改变写地址 | 无 session/model/command/credentials；A0/XFE1/越界/非法值 typed 拒绝且 snapshot 不变 | Pending |
| OO-01.b | typed literal、formula 与 calculatedValue；style 保留；隐藏行仍可读 | viewer 写拒绝；range-hidden 读拒绝；formula-hidden 不泄漏公式；失效对象不发请求 | Pending |
| MWB-01.a | 同 SDK 同时打开两本；重复 open 为同对象；只关闭来源后目标仍可操作 | subject/SDK dispose 退休旧对象；并发/StrictMode 不产生重复 session/writer | Pending |
| MWB-02.a | 实际 Java/H2 中 source A1=10/A2=20；target SUM=30；source 更新 A2=40，提交/刷新后 target=50 | foreign scope/非法 alias 拒绝；来源 403 后 external cells 清空、#BLOCKED!，状态可观察 | Pending |
| MWB-02.b | 保存、刷新页面、实际服务端 snapshot 与 native file 同一公式/身份 | 真实浏览器 console/pageerror/network 检查；不使用 HTTP route mocks | Pending |

完成整个代码批次后统一执行 TypeScript/Vite、SDK 成功/拒绝测试、既有边界、全量 unit、真实 Java/H2 浏览器 UAT，并记录失败；不按测试输出临时拼补架构。全部原始产品项与新增多工作簿项通过，才允许合并。
