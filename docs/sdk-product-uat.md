# SDK 产品目标与逐项 UAT

基线：JunoBarnswell/luckysheet-local，main，22d82876fd63fa253e51f9d4ffbbe5b288518409（Merge pull request #348）。2026-10-03 开始，分支 codex/sdk-product-kernel。

用户要求直接实施，附件“本轮只审计、不写代码”不作为本次指令。附件产品目标全部保持；不以包装 WorkbookSession、重导出内部包、改包名冒充 SDK 完成。

## 产品设计与交付约束

SDK 拥有认证/身份/凭证/组合根/工作簿目录与生命周期/Excel 领域/原生文档。Web 只接收只读快照、capability、typed error，转交用户动作和浏览器宿主接口。写链唯一：SDK action → domain → command → mutation → canonical model → formula/history/collaboration/persistence projection。

角色来自生成契约，系统管理员与工作簿权限分离。SDK 不公开 token/CSRF/任意业务 HTTP。Server workbook 是业务权威；恢复 journal 与 cache 不创建离线业务权威。旧入口在相应功能迁移时直接删除，不新增兼容层。每个功能单独 commit 并推送同一个草稿 PR。

验收只用 Pending/Pass/Fail/Blocked。测试不能代替真实浏览器或桌面 Excel。全部必需项 Pass 且 PR 检查通过后合并。

## UAT 细项（实施前制定）

执行后逐项记录实际命令、测试标题、结果和证据。未执行不得标 Pass。

| ID | 前置条件与操作 | 预期结果与拒绝路径 | 状态 |
|---|---|---|---|
| ROLE-01 | 生成 TS/Java 契约并扫描全部角色消费者 | 一个生成 WorkbookRole 来源；非法角色拒绝；wire labels 不变 | Pass |
| ROLE-02 | 四种角色分别打开活动/已删除工作簿菜单 | SDK capability 决定编辑/共享/移动/恢复/purge；Web 不解释角色 | Pending |
| AUTH-01 | 新部署使用初始化凭据建立管理员 | 真实认证与 CSRF 轮换；凭据不进入公共快照/工作簿 | Pending |
| AUTH-02 | 正确/错误密码登录、登出、刷新 | 身份跟随服务器；失败不建立身份；登出撤销会话 | Pending |
| AUTH-03 | OIDC 登录/回调/silent renew/过期/登出 | SDK 拥有生命周期；Web 不接触 token；过期不能沿用身份 | Pending |
| AUTH-04 | 网络失败/未知模式/损坏会话响应后重试 | typed error 含操作与恢复办法；失效时清除身份/凭据 | Pending |
| ID-01 | 管理员创建/禁用/启用用户、重置密码 | SDK 完成真实请求；禁用/重置使旧会话失效 | Pending |
| ID-02 | 普通用户调用管理员动作 | SDK 拒绝且服务器独立授权；无任意 request 公共入口 | Pending |
| RT-01 | StrictMode 挂载/清理/重挂载/最终卸载 | 唯一 SDK API/persistence/asset/session owner；释放资源 | Pending |
| RT-02 | subject 切换且有工作簿与恢复记录 | 释放旧 owner；journal 按 subject 隔离，不能跨身份重放 | Pending |
| HUB-01 | 创建→目录→打开→编辑→保存→刷新 | 服务端为唯一 authority，版本/快照一致；错误不变本地空表 | Pending |
| HUB-02 | 重命名/复制/移入文件夹/星标后再打开 | SDK 生命周期动作与服务端/目录一致，无双写 | Pending |
| HUB-03 | 删除→恢复；再次删除→永久删除 | 权限与转换正确；未删除或非 owner purge 拒绝 | Pending |
| HUB-04 | 快速切换目录/搜索/文件夹并取消旧请求 | 旧响应不覆盖新投影；folder cycle 可观察错误 | Pending |
| DOC-01 | 含公式/样式/DV/CF/table/objects 的 XLSX 导入编辑导出 | 显式 capability，唯一 owner，未知 parts 保真 | Pending |
| DOC-02 | XLSM/VBA/chart/pivot/OLE 等 preserve-only 文件 | 操作前可查 editable/preserved/unsupported/host-owned；破坏性编辑原子拒绝 | Pending |
| DOC-03 | XLSB/XLS/ODS/XMLSS/CSV/Text/SJS/SSJSON 实文件往返 | 逐格式区分读/编辑/写/保留，不以 preserve 计实现 | Pending |
| DOC-04 | 桌面 Excel 打开导出文件、重算、保存、重新导入 | 无修复提示，公式/数据/对象正确；无桌面 Excel 则 Blocked | Pending |
| EDIT-01 | 值/公式/富文本/格式/merge/clear/fill/Undo/Redo | SDK 写链、权限、history/persistence/server 一致 | Pending |
| INPUT-01 | mouse/Shift/Ctrl/Enter/Tab/F2/IME/name box/autocomplete | SDK 输入语义，选中/编辑/提交地址一致；失败不丢内容 | Pending |
| PANE-01 | frozen panes/hidden rows/columns/zoom/scroll 后编辑 | PaneMap 唯一 owner，pane 互斥；hidden 不影响 canonical read | Pending |
| SIZE-01 | 行列 resize/autofit/multi-column/取消手势 | SDK 拥有尺寸计划/worker，宿主只转发事件/测量 | Pending |
| CLIP-01 | copy/cut/paste/special/transpose/theme/跨 sheet cut | 一个 RangeMoveTransaction 覆盖源目标/formula/names/DV/CF/table/drawing/history/collab | Pending |
| STRUCT-01 | 插删行列/cell shift 含全部引用对象 | canonical transform 更新全部 owner，无各域核心场景拒绝 | Pending |
| STRUCT-02 | 跨 sheet move/部分限定 range/sheet rename/delete/reorder | 统一 Reference Ownership Graph/server patch，selection/pending/history 一致 | Pending |
| TABLE-01 | header/total row/table column 下插删/shift | 唯一 table transaction，structured refs/filter/extent/history 一致 | Pending |
| FORM-01 | dynamic array/spill/names/structured/external/volatile/manual/worker | client/server 同语义；spill/过期任务原子拒绝 | Pending |
| DV-01 | formula/custom/list/named/cross-sheet validated cells | server 按同一计算语义接受/拒绝，无 formula-under-validation split | Pending |
| DATA-01 | sort/worksheet-table filter/manual hidden/outline | 唯一 filter owner、row visibility/permutation | Pending |
| DATA-02 | Subtotal 全部聚合/advanced filter/DV/CF | 声明完整函数集，formula/structural/collab/persistence 一致 | Pending |
| DATA-03 | block/query/linked record 编辑重算与结构变化 | server authority，错误不默认空结果；source 类型不改变语义 | Pending |
| CHART-01 | 跨 sheet/title/axis/series/label formulas 与复杂 chart | SDK 拥有来源与语义，Web 不限制，native owner 一致 | Pending |
| PIVOT-01 | range/table/block/query/linked source 与 chart/slicer/timeline | 统一 TabularSource，核心引擎不按 provenance 分叉 | Pending |
| OBJ-01 | drawing/shape/connector/image/textbox/control/barcode/camera/sparkline | SDK hit/transform/permission/history/native owner 一致 | Pending |
| REVIEW-01 | comment/note/link/revision/audit/restore | 权限与 server history 一致，结构变化引用正确 | Pending |
| COLLAB-01 | 双浏览器 presence/cursor/commit/replay/reconnect/resync | operation 唯一写链，guest WS 授权，无丢失/重复 | Pending |
| COLLAB-02 | Undo/Redo 与 remote move/paste/shift/sheet/table/permutation 交错 | canonical rebase 或明确事务冲突，无大面积 invalidate | Pending |
| ACCESS-01 | 四 workbook roles × range hidden/read/edit × protection | 分层权限，server 独立拒绝越权 | Pending |
| OUT-01 | page setup/print area/title/preview/PDF/export/save as | SDK 拥有版式/文档，Web 只下载/打印宿主动作 | Pending |
| BOUND-01 | 扫描 Web import/API/Worker/model/command/token | 仅 sdk/ui-system 业务依赖，无内部路径/HTTP/worker | Pending |
| BOUND-02 | 注入越层 import/相对路径绕过/未声明依赖/cycle | 门禁全部拒绝；正常 graph 通过 | Pending |
| GATE-01 | build/typecheck/unit/boundaries/contracts/Java/browser/PR checks | 保存实际结果；全部通过后 ready 并 merge 指定 head | Pending |

## 已确认入口与删除要求

- ApplicationServicesProvider 创建 API/persistence/assets/session options，需迁入 SDK。
- Web auth/session.ts 与 oidc.ts 拥有 fetch/CSRF/UserManager，迁移后删除原文件。
- protocol WorkbookAclRole、permission ShareRole、persistence WorkspaceRole、Web WorkbookRole、Java enum 重复定义，统一生成契约。
- workbook-session.ts 8204 行，需领域拆分，不得新增 facade 后宣称完成。
- Web column-dimension-controller/autofit worker/command-controller/canvas interaction 需逐域迁移。

## 执行记录

完整产品目标未通过前保持草稿 PR，不合并。

2026-10-03：ROLE-01 已通过 TS/Java 生成契约与成功/拒绝测试；大写/未知角色被拒绝，canonical wire labels 与数据库 enum 常量不变。前端 build、boundary/contracts/mutation registry gate 通过。完整 Java/H2 测试 45 suites、319 tests、0 failures/errors/skipped（临时 Corretto JDK21 与 Maven3.9.11，代理使用系统 trust store）。

SDK auth/identity/runtime 13 项测试通过，包括 CSRF rotation、公共快照不泄露凭证、未知模式/损坏响应/网络失败、管理员拒绝路径、subject owner 切换、StrictMode 与最终释放。Web 旧 auth/session.ts、auth/oidc.ts 与 composition root 构造逻辑已删除；目录 service 实现移入 SDK，旧 service 文件删除。公共 hook 的会话领域拆分仍未完成，不能将 useWorkbook 计作完整 SDK Workbook API。

真实浏览器 UAT 已编写 e2e/sdk-product.spec.ts，需隔离真实 Java/H2 数据目录，SDK_UAT_ENABLED=1。当前尚未执行通过：仓库 provenance gate 要求 clean source tree，因此先提交实现后执行验收；不降低门禁。

