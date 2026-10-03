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
| F1 | 455 个官方函数目录对齐 | 逐函数参数/类型/精度/错误/日期/数组向量与 Worker；本批新增12财务函数，executable共148（官方目录重合147+SJS.TABLE），当前缺失308个；其余财务和现代数组仍待实施 |
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

每个完成的功能批次提交到现有草稿 PR #349；ERP 变更使用独立 codex 分支和草稿 PR。验收证据记录实际代码 head、环境、步骤、断言、console/network 与失败恢复；真实 SSO 和桌面 Excel 缺失仍为 Blocked。原有 54 项单测失败和其余未完成验收项不会因计划或局部门禁通过而自动通过。

## C1 实现批次验证

统一实现输入 fault 类型、根求值输入边界、流式/稀疏/数组/名称/派生读取观察、外部与隐藏范围故障、Worker 元数据校验及 audit trace 后，统一执行验证。TypeScript/Vite build Pass；calculation-domain 主进程 450/450 Pass，补充结构/表达式 5/5 Pass。其中新增 4 组检查覆盖 12 种消费式、三类来源失败、恢复、Worker、普通错误、懒分支、隐藏范围和坏元数据。

首次真实浏览器执行被 E2E provenance 的 clean-source 前置条件拒绝，未进入业务步骤（`/tmp/sdk-input-fault-browser.log`）。产品实现保持冻结，提交同一批完整源代码后，在干净的明确 head 上执行浏览器验收；不删除或绕过来源门禁。完整单测旧 54 项失败仍需后续逐项验证。


C1 干净产品 head `70c3920499b50ab6fa636662f00582ddc16d9e47` 的真实补充 browser **2/2 Pass**：撤权五函数阻断及恢复步骤实际执行通过，subject 切换与 active dispose 通过。AUTHX-09、CALCF-01、CALCF-02 的输入故障细项通过真实授权场景及新增成功/拒绝语义检查；其他身份项仍待实现/验收。完整 unit 1555/1501 Pass/54 Fail，与旧失败标题无增减。详细证据见 sdk-product-uat.md 的 C1 执行记录。


## I1 身份批次的冻结契约

统一 local / OIDC / external-bearer / host-session 凭证 owner；输入配置删除旧 oidc 选项，使用 source 判别契约，Web 和所有消费者同时迁移。OIDC 浏览器 profile 只用于凭证生命周期，身份通过 Java auth/session 验证。source 不提供可授予权限的 subject；公开 context 仅包含 authority、subject、ACL principal、scope、非秘密会话 nonce、tenant/app/employment/contextVersion 和 contextId。

REST/WS 采用同一受验证 JWT identity 转换，主体 key 由可信 authority 与工作区命名空间产生。Workbook 存储 identity_scope，owner/ACL/space-role 计算前先拒绝异域 scope；目录查询、空间角色、ACL/member 写同样拒绝跨 scope。V14 为显式存储迁移：既有 local 资源保持 local；旧未限定 OIDC subject 必须在停机迁移边界以可信 issuer/subject 映射重绑 workbook/ACL/space/member/user-state/range principals，不做运行时别名或回退。未迁移资源不会被新身份自动接管。迁移前备份数据库；回滚恢复备份和匹配代码。

每个 runtime API/凭证闭包/fetch/响应体和 WS 捕获已验证 context；切换或登出退休对象、目录、cache、assets、协作和恢复 journal。journal 使用完整 contextId。同 context 凭证续期先验新凭证，保留 owner；旧上下文请求不能携新凭证写入或发布结果，不自动重试写入。Cookie 模式要求 CSRF；host-session 从 Gateway owner 获取 CSRF，保留路由 prefix，bearer omit cookies。凭证仅用于固定可信 HTTP/WS 来源。local nonce 与 JSESSIONID 分离并在登录时轮换；local logout 关闭对应 context 的 WS，JWT/guest WS 到期关闭。

部署需声明受验证 token 的 session claim（默认 sid）；tenant/app/employment 必须共同存在且提供非负整数 contextVersion。可信 subject/authority 映射仅为部署配置。JWT 提前撤销仍取决于 IdP 的撤销/实时授权能力；I2 的 Main 断言和当前授权不可被此批的签名验证代替。

I1 先完成上述整个实现批次，再执行 build/SDK/Java/浏览器验证。AUTHX-03/04/05/07/10/11 先以成功及拒绝契约检查覆盖，真实外部身份集成不以测试凭证或本地 issuer 算作 Pass。AUTHX-01/02/08/真实 ERP 切换由 I2 交付验收；全部验收前 PR 仍草稿。

I1 统一验证：build/typecheck、现有 boundaries Pass；SDK 42/42 Pass；Java 327/327 Pass + package 成功。首轮 Java 326 tests 只有迁移测试仍断言版本 13，而新正式迁移为 14；产品实现保持冻结，只同步该测试的显式版本契约，并增加 v13→v14 数据/owner 不变、不擅自发明 OIDC 映射及重复启动历史不变的验收。签名/issuer/audience/过期/缺 sid 五类拒绝测试和 namespace/scope/WS 检查通过。随后在干净提交 head 执行真实本地浏览器；真实 ERP/SSO 项尚未通过。

I1 产品 head `1100d699cc8f8d724fcf367dde3e3276666e4af3` 的完整真实 Java 21/H2/Chromium UAT **15/15 Pass**（`/tmp/sdk-identity-browser.log`，`/tmp/sdk-product-uat-GLepoX/evidence`），无 HTTP route mocks。完整 unit 1555/1501 Pass/54 Fail，失败标题仍无增减；不会据此合并。外部 IdP/ERP 生命周期集成仍 Blocked，不将签名测试等同真实 SSO。

## C2 多簿图冻结契约与实施前细项

删除单 binding 输入端点/客户端方法，统一 ExternalCalculationGraph：root、受验证 subject、每个来源 unit 的 revision/accessRevision、计算快照/hidden ranges 或明确拒绝状态。服务端从 canonical snapshot externalLinks 构造闭包，图门控行串行化 topology 改变，再按 unitId 排序锁定现有来源，刷新 ORM 读取并验证闭包；一次响应中的来源版本固定。不会复制持久化依赖模型。V15 只创建图门控行，普通单簿值事务仍由原来的事务链拥有。

提交后的最终候选 externalLinks、客户端绑定前置和新建/导入/复制快照在同一个图契约检查循环、来源和工作区授权。循环直接 CIRCULAR_DEPENDENCY 拒绝；没有假装实现多簿迭代。64 本、每本 100000 输入和整图 1000000 输入预算是明确 UNSUPPORTED_FEATURE 拒绝，没有截断。

客户端按照拓扑顺序使用同一 FormulaEngine/Worker 计算每个来源；C1 输入 fault 沿 B→C 再次消费仍不可吞掉。版本/绑定/runtime/context 不匹配不发布；失败来源移除旧输入。来源权限撤销是图内 denied 节点，不包含快照，目标合法授权仍可获得 typed 失败图；不能把合法目标的 HTTP 200 当作失败来源已授权。

WebSocket 增加正式 calculation.subscribe / calculation.changed 契约；闭包和订阅来源由服务器产生。订阅是事件投影，不是第二个模型、凭证 owner 或影子 WorkbookSession。目标当前 VIEWER 权限在每次通知前再验证；消息只标识已绑定来源变化，不发送来源数据。commit/access/lifecycle 事件触发目标重新获取授权图；关闭来源 Workbook 不取消目标的服务器依赖，关闭目标/退出 context 会撤销所有订阅。刷新期间收到新事件只标记新一轮工作，不发布旧任务。

预设 MWB-03.a：A=10，B=A*2，C=SUM(B)、COUNT(B)、IFERROR(B)，一份授权图重算为 20/1/20；A 提交 40 后，不调用手动 refresh，C 自动变为 80/1/80，保存重开一致。MWB-03.b：关闭 A 的公开对象后仍观察真实服务端 A 的更新；撤销叶子 A 权限后三个 C 结果均 #BLOCKED!，图无 A 快照；恢复后自动恢复授权版本。MWB-03.c：尝试 C→A→B→C 绑定及伪造 REST externalLink.set，均 CIRCULAR_DEPENDENCY 拒绝，所有快照、revision、history/operations 不变；并发两个方向绑定至多一个成功。MWB-03.d：坏 schema、重复 node、坏版本、缺节点、越界、预算失败与过期 runtime 不能发布/留下旧来源缓存；inline/Worker 同结果。所有细项先 Pending，之后记录真实执行。

当前图捕获仅接受已物化输入。带数据块的来源统一返回 UNSUPPORTED_FEATURE，避免缺少 block hydration 时读成空值；完整流式来源由 R1 批次处理。

### C2 事务契约审查后的完整修正批次

首轮 b8a9fbdd 浏览器 7 Pass/3 Fail/6 未执行；实际三簿计算、自动更新、撤权、恢复和重开断言通过，但图 GET 的额外 409 导致 network/console 验收失败，原始 trace 在 `/tmp/sdk-graph-first-attempt-test-results`。409 body 是 ORM 乐观版本冲突，不是声明的 EXTERNAL_GRAPH_CHANGED。未锁读取的 JPA entity 留在 persistence context，随后升级 pessimistic lock 验证旧版本，说明锁前 entity 生命周期违反版本捕获契约。不得隐藏这些 409 或重试请求。

完整修正采用一次规范拓扑闭包收集（仅内部锁集，不是持久化/read model），锁前 flush/clear 去掉未锁 entity，按稳定 ID 获取所有规范依赖的数据库锁，再清理 ORM 并产生授权响应。拒绝节点的后代也只内部锁定，响应只遍历授权可见闭包，禁止暴露拒绝节点的快照或内部后代；这样权限变化不会改变已锁集合。新建/复制/导入、永久删除和历史恢复都经过同一图门控；恢复候选必须重新检查循环和授权。来源 trash/purge 是 broken 输入并发送生命周期失效通知，restore-from-trash 重新获取授权版本。

新增实施前细项 MWB-03.e：事务已有旧 source entity 后另一真实事务写入 40，图必须读到新的版本与值且无 409；锁前缓存不能代表来源版本。MWB-03.f：A→B 曾合法，移除后 B→A，再恢复旧 A→B 必须整体拒绝，无新 operation/revision；来源 trash→restore→trash/purge 自动触发 #REF!→80/1/80→#REF!，坏来源节点无 snapshot。修正整个上述事务链后再统一验证，既有拒绝断言保持。

e8bf2736 第二轮 15/16 Pass；所有旧图 409 已消失，新增 trash/restore/purge 的依赖结果均通过。最后 network 断言发现仍打开的来源自身 WS 被关闭后触发 access/snapshot 403：生命周期仅关闭连接，尚未退休宿主对象。下一完整所有权批次新增正式 server-only workbook.lifecycle.changed，单 Session 先处理生命周期再关闭 WS，dispose 通过内部 lifetime port 通知公开 Workbook 退休；ApplicationRuntime 现有 release 回收 session/object/catalog open cache。source 的 dependent subscriptions 保留，目标重取图；恢复创建新对象，旧 Cell/Range 永久 RUNTIME_DISPOSED。不再让已退休来源重新请求自身 snapshot，不放宽 network/console。

实施前 MWB-03.g：来源 A 打开时 trash 后旧 Cell 报 RUNTIME_DISPOSED，C 自动 #REF!；restore 后重新 open A 能读 40、旧 Cell 仍拒绝；purge 不发生来源自己的 access/snapshot 403；来源 lifecycle 不能被客户端伪造，也不能退休别的 unit。Session 直接 dispose 同样退休全部公开句柄，仍打开的其他工作簿保持可用。

## O1.1 范围与工作表对象的冻结契约

此批交付矩阵值/公式、不可变范围读取、样式/number format、清空、填充、合并、工作表身份与排序、行列结构/大小/可见性、pane 和规范 undo/redo。完整 O1/O2、跨簿移动和 T1 仍不能仅凭此子批通过。所有地址显式绑定工作表稳定 ID，不消费 Web 的当前 selection。对象只调用私有领域访问器；公开 SDK 不暴露 Session、model、dispatch 或权限 owner。

异构矩阵使用一个 sheet.cells.commitMatrix 命令，按行优先顺序准备完整输入，要求位于明确的 canonical extent 内（先以 Worksheet.growExtent 扩展），限制单次 10000 个单元格、精确矩形、有限标量或以 = 开头的公式。准备全部完成后才应用 cell.set，沿用 writeAuthority、spill、checkbox、DV、权限和单一历史/协作链；任何失败由同一事务回滚。公式和 typed value 保留已有样式，清除旧公式计算/来源和 rich text。禁止通过循环公开 Cell.setValue 伪造批量原子提交。记录/只读投影目标明确拒绝，不越过 record owner；数据区等其他 owner 由现有 dispatch 前置处理。

读取先对整块地址和真实来源授权，再加载所需内容、等待一次规范计算，并重新授权；最后同步组装不可变结果，不逐 cell 启动外部图/计算。最大 10000 输入是明确资源拒绝而非截断。样式和 rich text 是 authored snapshot，不假装包含最终 render/computed style。copyValuesTo 是授权读取结果后对目标的一次值写入，不是 T1 的跨簿移动/一致读事务；错误值目前不是可写 CellValue，明确拒绝，不能转成字符串。

Worksheet 的 rename/remove/add/duplicate/reorder、行列 insert/delete 和 pane 等使用现有规范命令/Java structural planner，不在 SDK 重写引用。工作表删除后其句柄访问明确失败；同一历史 undo 恢复稳定 ID 后句柄重新有效。Workbook close/context/lifecycle 退休仍永久有效。undo/redo 通过现有 CommandRuntime，权限或过期结构版本失败必须返回 typed 拒绝，不能仅显示 UI notice 然后宣告成功；没有历史返回 false。修改与其他对象 API 一样进入 pending canonical operation，flush/save 是持久化完成边界。

实施前全部 Pending：O1.1-a 独立公开 SDK 一次写 D8:E9 的混合有限值/字面 = 文本，选区仍 A1，单次 undo/redo 整体恢复，保存重开一致；O1.1-b 一次公式矩阵计算、公式输入坏值/尺寸/预算/非有限值/隐藏或 viewer 目标整块拒绝，快照与历史不变；O1.1-c 样式/number format 保留于值写入，contents clear 保留格式，formats clear 保留值，不可变快照不可回写；O1.1-d sheet add/rename/reorder/duplicate/remove 与规范引用、稳定对象/删除后访问/undo 恢复/保存重开；O1.1-e 行列 insert/delete、像素大小与隐藏、pane 实际保存和恢复，隐藏值仍可读；O1.1-f fill/merge/unmerge 的实际值、公式移动与历史，数据损失明确确认；O1.1-g 两簿 copyValuesTo 成功和源隐藏/错误值拒绝，来源不修改且不宣称跨簿原子事务。真实浏览器所有步骤记录 console/network；真实 Excel 互操作仍 Blocked。

O1.1 首轮冻结源差异在 /tmp/sdk-objects-first-pass.patch。SDK 51/52：唯一失败是验收 fixture 使用不存在的 sheet.dataValidation.add，实际正式入口为 sheet.dv.add；不添加产品别名。计算 457/457、boundaries Pass。build 的唯一失败是异步闭包中 source.region 的类型收窄丢失。完整异步契约复核还要求所有 source/query 前置检查完成后才启动加载，以及矩阵/descriptor 在 await 前捕获调用输入；下一修正批次统一落实这些条件。实施前 O1.1-h：setInputs 返回 Promise 后调用者改变原矩阵，不得改变已捕获的写入意图；缺任一 query 不启动先前查询、整块拒绝，不出现未处理 Promise 拒绝。

O1.1 产品 head 83261a156c45316e3c6dab60e8098ff5eaf9a047 全浏览器 17/19 Pass：原 16 项及 viewer/错误值复制 Pass；矩阵/样式/历史的业务步骤完成后，在关闭全部 Workbook 后重开失败；结构场景在 rename 提交失败。原 trace /tmp/sdk-objects-first-browser-test-results，证据 /tmp/sdk-product-uat-wc2OyW/evidence。rename 的真实 REST operations 返回 201，server StructuralPatch v10 包含完整公式改写，但本地 rename effect 对单元格公式遗漏 owner deltas，故客户端无法确认，不能重试提交或忽略 server patch。

下一整批所有权修正：createSpreadsheetSdk 持有 SDK 根 lease 直至 dispose，不把最后 Workbook/UI release 等同身份/context 退休。身份切换仍统一退休旧目录/对象。rename 删除单独的 cells/rules/names/advanced owners 改写，调用已有 planWorkbookFormulaRewrite，返回所有规范 before/after facts；保留不可编辑 preserved metadata 的明确拒绝，不降级解析或放宽结构断言。实施前 O1.1-i：关闭全部 Workbook 后同 SDK 重开、原 catalog 仍有效；dispose/身份变化旧 catalog 仍拒绝。O1.1-j：rename cell/CF/DV/名称/持久化公式 owner 的本地与 Java facts 一致，一次 ACK 后 undo/redo 正确；篡改服务器 before/after 或未知 owner 拒绝且不进入 terminal，原模型/历史不修补。

## O1.1 计算声明与无损快照修正批次（实施前冻结）

真实 e435d678 UAT 17/19 通过；矩阵撤销恢复了 authored values 却留下清空后的计算输入，删除工作表的撤销 REST 409。保留 `/tmp/sdk-objects-ownership-browser-test-results` 与 `/tmp/sdk-product-uat-PvpoCy/evidence`。对隔离 H2 的只读 canonical operation replay 证明撤销候选与服务端 preimage 仅差 `pane.activePane`：合法可选字段被前端 snapshot/fromSnapshot/duplicate 自动补出。服务端 `UNDO_RESULT_MISMATCH` 校验保持不变。

删除 runtime 的 FORMULA_SYNC_MUTATIONS / DIRECT_CELL_WRITE_MUTATIONS / VISIBILITY_MUTATIONS；MutationRegistrationMetadata 必须提供 calculation 声明，包含 inputs（none/cells）、visibility（布尔）、spillBlockers（none/ranges/table-deltas）、mode（布尔）和可选 typed context。注册时缺字段、未知字段/枚举或无效 context 均拒绝；不存在隐式按 mutation 名称补行为的通道。所有 production owner 和正式测试注册同步迁移，forward/inverse 独立声明各自真实行为。结构变换仍使用实际 StructuralTransformResult；声明负责普通输入、可见性和模式，context 重建/增量事实仍只有原 FormulaEngine/Worker owner。输入与几何 roots 合并，不能由几何更新覆盖输入同步；manual 模式同步 inputs 但不自动求值。record writes 使用其 canonical CellMatrix 投影，关系成员变化继续从现有 relation owner 派生。新增 find.replaced 同样声明真实输入变化，workbook.restore 声明重建。

WorksheetModel.snapshot/fromSnapshot/duplicate 验证并复制原 pane，不归一化合法可选字段；删除无消费者的 runtime normalizeWorksheetPane。字段缺失与字段显式存在均精确保留，不在运行时改版本或补默认值。PaneMap 坐标与 OOXML 格式边界不增加新的持久化 owner。

统一代码批次落地后一次执行 frontend build/SDK/calculation/boundaries/full unit 和合并后 Java 全测试/package，再在干净 head 上跑完整真实 SDK UAT。继承 main/security 变更已经保留于 371bc4d6，不以旧 Jar 或旧验证证明新 head 通过。所有失败原始证据保留，不降低服务端完整恢复、权限或公式断言。

脏根地址补充契约：canonical cellAddressKey 已记录完整 sheet/row/column，清空输入仍必须导出这个根。pending 根导出与遍历共用 key 解码，不以 live cells 查找/过滤，不新建第二份脏根集合。Worker bootstrap pendingRoots 与后续 input deltas 表达同一来源变化，manual 不自动重算，显式 F9 同步结果。原冻结的 k..o 条件保持不变。

## O2.1 名称、富文本和工作表保护（实施前冻结）

本子批扩展公开对象入口：Workbook.names（list/byName/define/remove，DefinedName.snapshot/setFormula/remove）；Cell/Range.setRichText(text,runs)，Cell.setStyle/setNumberFormat/setBorders 复用原 Range 格式 owner；Worksheet.protection（list/set/remove）。只保存句柄身份，名称、富文本、保护状态仍从 canonical model 读取。返回的配置递归 immutable，关闭/身份切换后所有新对象退休。独立消费者不取得 Session/model/dispatch/token。

名称采用正式 DefinedNameModel，scope 必填，sheet/anchor 用明确稳定 ID；旧 workbook.name.set 的 value 别名删除，UI 与测试正式调用同时改为 formula，并补齐 anchor 传递。相对名称不猜测 active cell，原 engine 对缺 anchor 的拒绝继续保留。定义、修改、移除与撤销使用已有 name.set/name.remove，增量同步唯一 FormulaEngine。

富文本使用现有 commitRichText / commitRichTextCells：plain text 必须等于 runs 拼接，整个目标预授权和准备后一个事务；不解析以 = 开头的富文本为公式，不改变格式/selection，不逐 Cell 写。拒绝隐藏、readonly/record/projected owner、预算/extent、坏 runs、viewer、DV 或 spill child，不写部分内容。Cell 的格式动作按显式单格 Range 调用原 owner。

保护的 canonical 校验由 core rule validator 与 Java reducer 分别在各自正式边界落实，scope/range/sheet/allow 布尔/未知字段均严格校验，错误不进入 model/history/server storage。消除 Number.MAX_SAFE_INTEGER 伪造整张范围：sheet-wide 保护使用已存在的 Java 空 affected-range 全表语义，range 保护使用真实范围；owner ACL 与工作表保护仍独立，设置保护不授予 ACL。worksheet/range 才由 Worksheet.protection 修改；workbook structure/password 安全能力属于 S1，不能用单张 sheet rule 假装实现。不存在的保护 remove 拒绝且不产生 history。当前规范状态 snapshot/load 保留全部合法元数据，不改版本、不隐式迁移。

Theme 和模板库 owner 的完整改造留给 O2.2；本批不把部分字段暴露算 O2 全量完成。实现全部子批改动后再统一 build/SDK/calculation/full unit/boundaries/Java 和真实 UAT；每完成一子批推送同一草稿 PR。

## O1.2 同表范围移动：实施前冻结

公开 Range.moveTo(destination) 只接受同 Workbook、同 Worksheet、相同矩形尺寸且各不超过 10000 单元格；跨表/跨簿移动属于 T1，明确 UNSUPPORTED_FEATURE，不能拆成 copy+clear。移动为显式 cut/replace，目标已有内容被替换并由规范历史完整恢复；句柄仍绑定原地址，不追随单元格。命令入口为已有 sheet.range.move，经 Session 的 resolved-write 前置、Java structural planner、CommandRuntime 和完整 preimage/history 提交。离线不能绕过 planner。

删除 TS 与 Java 的 moved-formula-owner 跳过路径。移动区域内公式、normal sourceFormula 与 barcode formula 同样通过已有 moved-reference transform；指向移动源的绝对/相对引用改为目标，区域外引用保留，不按 copy 偏移。TS 在源地址预先准备并记录目标 afterAddress，Java 在移动后的目标解析原 beforeAddress；两者使用同一正式 StructuralPatch v10 事实。失效计算 cache 不回填；规则、名称、图表及其他 persisted owner 继续通过原 transform。保留可逆性、未知/公式组、区域/元数据不支持、权限与目标重叠拒绝。无快照/协议升级，无额外模型或逆操作栈。

实施前全部 Pending：O1.2-a 移动包含值与公式的矩形，区域内/外及跨工作表依赖正确，原地址空、目标旧内容替换，绝对引用与 sourceFormula/barcode 一致；单独移动公式但其输入不移动时引用不变。O1.2-b Java 与前端 before/after 地址和公式一致；单次 undo/redo 完整恢复源、目标旧值、外部引用并经真实 REST 确认。O1.2-c 重叠、错尺寸、跨表/簿、超预算、已退休对象、离线、viewer/hidden、未知公式组和不可逆引用前置拒绝，模型/历史/原输入不改变。O1.2-d 独立公开 SDK 真 Java/H2/Chromium 保存重开、导出实 xlsx 并原生解析/服务器重导入后结果相同，console/network 干净；桌面 Excel 仍 Blocked。先完成这一整个实现批次再统一检查，失败冻结证据与代码，不逐断言改语义。

O1.2 整批实现后统一检查：SDK64/64、calculation-domain549/549、boundaries Pass；full unit1681/1641 Pass/40 Fail，范围移动原失败真实修正，其他失败标题无新增。实际移动区域内外引用、绝对/未移动引用、sourceFormula/barcode、before/after地址、cache失效及公式组/不可逆拒绝都有断言。统一构建、Java及干净head真实UAT结果待记录，不以本段局部结果宣告验收完成。

统一 build/typecheck Pass，Java21 package Pass：370 tests、368 Pass、2 Skip（外部 PostgreSQL/MySQL 环境缺失，仍 Blocked）；日志 `/tmp/sdk-o12-build.log`、`/tmp/sdk-o12-java.log`。产品源保持本批冻结，提交后执行完整真实 Java/H2/Chromium UAT，不能以通过的单元断言替代真实撤销/原生往返。

O1.2 产品00853d8e首次真实完整UAT22/23 Pass；移动已ACK，第一次undo flush被Java409拒绝，尚未执行后续redo/native步骤。证据 `/tmp/sdk-o12-first-browser-test-results`、`/tmp/sdk-o12-browser.log`、`/tmp/sdk-product-uat-8eTAwV/evidence`。完整审查确认 inverse 为range.move加两个目标旧值的cell.restore；服务端 structural undo的removed-cell事实仅处理rows/columns.deleted，遗漏move替换的目标。不能放宽guard，也不删除旧值恢复断言。

下一修正契约：deleted-cell helper clean-break为structuralRemovedCellFacts，由每条目标操作的真实中间preimage和已有正式mutation descriptor提取move destination sparse cells；行列删除继续原事实来源。每个未匹配的cell.restore必须精确匹配sheet/address/previous且只消费一次事实。无关、source原值、错坐标/工作表、篡改旧值、非restore、重复恢复均拒绝。操作最终还必须完整恢复原preimage；该校验和ACL/范围授权/版本/subject/current-history条件不变。新增O1.2-e（实施前Pending）：正式registry生成move事实→reverse move+精确target restores完整恢复，全部上述恶意输入拒绝且originalpreimage不变；真实browser原断言全部保留。无协议或存储迁移。

O1.2恢复事实修正统一验证：Java21 package Pass，371 tests、369 Pass、2 Skip；新完整preimage恢复与七类篡改/重复/遗漏恢复拒绝断言通过，原行列删除的中间preimage和其他非法structural undo测试继续通过。boundaries Pass。前端产品源未改变，沿用00853d8e的build/SDK64/calculation549/full unit40结果；提交修正后再跑完整真实UAT，不弱化任何原browser断言。日志 `/tmp/sdk-o12-undo-java.log`、`/tmp/sdk-o12-undo-boundaries.log`。

### O1.2 验收结果及本次交付边界

2026-10-03，产品head `f9cac28bf1840a106d60a6a0ac19ade202ca9a8c`，完整真实Java21/H2/Chromium SDK UAT **23/23 Pass**，零HTTP mocks、retries0；日志 `/tmp/sdk-o12-undo-browser.log`，证据 `/tmp/sdk-product-uat-dJzwIP/evidence`。O1.2-a/b/d实际通过：内部/绝对/区域外未移动/外部及跨sheet引用、源清空/目标覆盖、首次undo恢复88/99、redo、单独移动公式输入不偏移、保存重开、实际sdk-range-cut.xlsx原生解析与服务器import后公共SDK结果相同；既有viewer场景实际拒绝move并保留源/目标。e由Java新正式preimage/精确事实及恶意输入测试通过。c的错尺寸/跨sheet/overlap/viewer由真实browser通过，跨Workbook/预算/offline/退休与公式组/不可逆引用由SDK/core/Java检查通过；独立真实hidden移动验收仍未追加，不把整个原O1父项升Pass。范围移动新浏览器console/network诊断通过；本轮Vite在旧测试关闭WS时记录一次代理ECONNRESET，未出现新增范围场景的browser diagnostics失败，保留原始日志。

构建/SDK64/calculation549/boundaries Pass；Java371/369 Pass/2 Skip。完整unit1681/1641 Pass/**40 Fail**，不是完整门禁通过。ERP/SSO、桌面Excel和外部SQL环境仍Blocked；完整计划的其他批次仍Pending。2026-10-03用户最新明确要求“提交代码并推送，合并这个pr吧”，本次按该指令提交当前已实现批次并请求正常GitHub合并；不将合并记作完整Aspose能力或全部UAT验收通过，不绕过GitHub分支保护。首次22/23失败的原trace和409证据保留。
