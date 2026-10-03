# SDK 完整实施契约与验收批次

2026-10-03（Asia/Taipei）。用户授权实施全部研究计划，同时支持 ERP 内嵌和独立应用。完整目标仍包含原始 49 项产品验收；本文增加实施批次和验收条件，不把研究、声明、保留或单测通过当成功能完成。

## 统一所有权

公开入口为 createSpreadsheetSdk。Workbook/Worksheet/Cell/Range 和全部领域对象按稳定 ID 与显式地址操作；Web 和程序调用消费同一 command → transaction → model → calculation/history/collaboration/persistence/document 链。宿主仅拥有 I/O、线程、凭证来源和登录导航，不拥有另一份模型、函数引擎或权限真相。一次完成一个已确定契约的实施批次，再统一验证；失败记录原因，下一次修正先完成契约审查，不随测试逐点修改语义。

## 身份与凭证

所有凭证来源（local、SDK OIDC、host-session、external-bearer）统一经过 AuthDomain。外部来源只交付凭证与变化通知；服务端产生受验证的 identity/context，客户端传入的 subject、角色、用户名、工作区标识都不能授予身份。ERP 固定研究版本 edc2482abb0dd5597d42428d070578b5ff17eafe。

ERP 浏览器会话 Cookie 是 HttpOnly，令牌由 Gateway 保存和续期。内嵌接入必须经 Gateway、Main 的工作区授权与正式 Excel 服务路由；独立应用可使用同一 SSO 的授权码流程或宿主已有且 audience 正确的 bearer。现有 MCP token exchange 不支持通用 Excel 换票。新增服务不能绕过 ERP 的受控部署、运行租约和当前权限。

身份 authority/subject 与租户、应用、任职、会话及 contextVersion 分别建模。两个来源只有通过可信服务端身份映射才能归为同一主体。工作区、主体或会话改变，旧对象、缓存、协作连接与待提交操作退休；同主体同上下文的有效凭证续期不创建第二个 owner。HTTP、WebSocket、资产、下载、导出、跨簿输入和恢复日志都使用同一上下文。每个请求和异步结果有上下文边界，旧请求不能用新身份提交或发布结果。全局身份管理与工作簿 ACL 是不同授权；ERP 管理员不能自动成为全部工作簿 owner。

## 输入前置故障

读取被撤销、隐藏、缺失或不可用的授权输入，产生包含 reason/source 的 FormulaInputFault；对外仍返回相应 #BLOCKED!/#REF!/#N/A。它不同于合法来源中的普通 Excel 错误值。根求值边界记录实际消费到的前置故障；流式、数组、名称、间接引用、派生单元格、record 和 Worker 不能用 COUNT/IFERROR/ISERROR/忽略错误选项把故障转换为数字或成功状态。没有消费的 IF 分支不读取来源。普通错误值仍按 Excel 函数语义处理。失效来源清空派生输入，恢复后以授权版本重算。

## 完整交付批次

| 批次 | 领域和交付 | 完成要求 |
|---|---|---|
| I1 | 统一身份、外部凭证、OIDC 与 host-session 生命周期 | SDK/Java/Web/WS/资产/恢复日志同一受验证上下文，成功和拒绝验收 |
| I2 | ERP 内嵌及独立身份接入 | 正式 Gateway/Main/Java 路由与授权，两种入口同一身份映射；真实 ERP 验收 |
| C1 | 授权输入前置故障 | COUNT 等所有消费函数阻断，普通错误语义、Worker、恢复正常 |
| C2 | 多簿依赖图、循环、变化传播 | 来源版本一致、过期任务不发布、关闭与撤权清理、迭代策略明确 |
| T1 | 多簿事务/复制移动/history | Java/H2 共同提交点与回滚、授权锁顺序、整体拒绝 |
| O1 | Workbook/Worksheet/Range/Cell 完整领域动作 | 批量、清空、填充、结构、复制、引用与 atomic undo，删除 Web 越层消费 |
| O2 | Style/RichText/Theme/Names/保护 | 程序对象与 UI 同一 owner、格式与权限真实保存/往返 |
| D1 | Table/filter/sort/DV/CF/数据源 | 明确范围、统一 owner、全部输入源与公式/引用一致 |
| D2 | Pivot/PivotChart/slicer/timeline/what-if | 缓存与视图生命周期、计算字段、真实数据刷新 |
| G1 | 图表/图片/形状/批注/控件/对象 | create/edit/delete、资产、锚点、引用与原生往返 |
| F1 | 455 个官方函数目录对齐 | 逐函数参数/类型/精度/错误/日期/数组向量与 Worker；当前缺失 320 个，优先财务和现代数组 |
| X1 | 全格式与转换 | 每格式独立检测/read/edit/write/preserve/render，实文件；不以保留计编辑 |
| P1 | page setup/分页/打印/PDF/图片/矢量 | 同一布局 owner、字体/分页一致、真实输出 |
| S1 | 加密/签名/属性/XML maps/VBA 项目 | 真实文件安全与文档语义；项目管理与执行能力分别声明 |
| R1 | 模板填充/流式处理/取消/资源预算 | 完整失败回滚、无数据截断或虚假成功 |
| H1 | Node/Java/browser 宿主与 SDK 发布 | 单一语义链、宿主注入 I/O/线程、公开类型/构建包/独立消费者 |
| A1 | 全量门禁及 UAT | 原 49 项与新增项逐条 Pass、PR 指定 head 全部检查通过后合并 |

## 实施前 UAT

| ID | 成功细项 | 拒绝/边界细项 | 初始状态 |
|---|---|---|---|
| AUTHX-01 | 已登录 ERP 打开两本 Excel，无二次账号登录；保存/刷新/下载同身份 | 未登录不建立 Workbook；Cookie 不进入快照 | Pending |
| AUTHX-02 | 独立应用授权码复用 SSO；与内嵌映射到同一主体 | 无会话或需要 MFA/同意时按 provider 规则处理 | Blocked（真实 SSO 环境尚不可用） |
| AUTHX-03 | 外部有效 bearer 完成 create/open/write/save/reopen | 错签名/issuer/audience/过期/缺声明均拒绝，服务器无写入 | Pending |
| AUTHX-04 | 并发初始化/续期单一 owner；同上下文旧对象继续有效 | 续期失败撤销凭证，旧请求结果不发布 | Pending |
| AUTHX-05 | 同用户切换 tenant/app/employment 后打开新对象 | 旧多个 Workbook/Cell/缓存/journal/WS 全部退休，跨范围访问拒绝 | Pending |
| AUTHX-06 | 登出、禁用、撤权、重新登录按新授权读取 | 所有 HTTP/WS/资产/导出路径立即失效，不使用旧缓存 | Pending |
| AUTHX-07 | 合法 Cookie 写入；合法 bearer 无 Cookie 的独立调用 | 缺失/错误 CSRF 拒绝 Cookie 写；凭证不发往任意来源 | Pending |
| AUTHX-08 | viewer/editor/owner 与 ERP 资源权限组合逐条验证 | ERP 管理员无 ACL 不能读写；伪造用户或租户头拒绝 | Pending |
| AUTHX-09 | 来源授权下 SUM/AVERAGE/COUNT/MIN/MAX 正确；更新和恢复重算 | 撤权后五种函数与 IFERROR/ISERROR/AGGREGATE 均保留输入故障 | Pending |
| CALCF-01 | 合法源中普通 #N/A 可被 COUNT 忽略、IFERROR 处理 | 授权输入 fault 不可被这些处理逻辑吞掉 | Pending |
| CALCF-02 | inline、Worker、spill/名称/派生单元格相同结果 | 不消费的 IF 分支不触发来源读取；坏故障元数据拒绝 | Pending |
| AUTHX-10 | 重连与并发多个工作簿同一 subject/context | 旧上下文异步响应、未提交操作不跨身份恢复 | Pending |
| AUTHX-11 | 公开 SDK 独立消费者完成完整真实流程 | public API/文件/日志无 token、CSRF、密码和内部 owner | Pending |
| FULL-01 | 上表每批对象调用 → 操作 → 保存 → 重开 → 原生输出逐项对照 | 无权限、非法参数、存储/版本失败整体不改变模型/资产/包 | Pending |

每个完成的功能批次提交到现有草稿 PR #349；ERP 变更使用独立 codex 分支和草稿 PR。验收证据记录实际代码 head、环境、步骤、断言、console/network 与失败恢复；真实 SSO 和桌面 Excel 缺失仍为 Blocked。当前 54 项单测失败、COUNT 撤权失败及其余原验收项均未因本计划自动通过。

## C1 实现批次验证

统一实现输入 fault 类型、根求值输入边界、流式/稀疏/数组/名称/派生读取观察、外部与隐藏范围故障、Worker 元数据校验及 audit trace 后，统一执行验证。TypeScript/Vite build Pass；calculation-domain 主进程 450/450 Pass，补充结构/表达式 5/5 Pass。其中新增 4 组检查覆盖 12 种消费式、三类来源失败、恢复、Worker、普通错误、懒分支、隐藏范围和坏元数据。

首次真实浏览器执行被 E2E provenance 的 clean-source 前置条件拒绝，未进入业务步骤（`/tmp/sdk-input-fault-browser.log`）。产品实现保持冻结，提交同一批完整源代码后，在干净的明确 head 上执行浏览器验收；不删除或绕过来源门禁。完整单测旧 54 项失败仍需后续逐项验证。
