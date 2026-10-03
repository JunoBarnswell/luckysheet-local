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
| ROLE-02 | 四种角色分别打开活动/已删除工作簿菜单 | SDK capability 决定编辑/共享/移动/恢复/purge；Web 不解释角色 | Pass |
| AUTH-01 | 新部署使用初始化凭据建立管理员 | 真实认证与 CSRF 轮换；凭据不进入公共快照/工作簿 | Pass |
| AUTH-02 | 正确/错误密码登录、登出、刷新 | 身份跟随服务器；失败不建立身份；登出撤销会话 | Pass |
| AUTH-03 | OIDC 登录/回调/silent renew/过期/登出 | SDK 拥有生命周期；Web 不接触 token；过期不能沿用身份 | Blocked |
| AUTH-04 | 网络失败/未知模式/损坏会话响应后重试 | typed error 含操作与恢复办法；失效时清除身份/凭据 | Pending |
| ID-01 | 管理员创建/禁用/启用用户、重置密码 | SDK 完成真实请求；禁用/重置使旧会话失效 | Pass |
| ID-02 | 普通用户调用管理员动作 | SDK 拒绝且服务器独立授权；无任意 request 公共入口 | Pass |
| RT-01 | StrictMode 挂载/清理/重挂载/最终卸载 | 唯一 SDK API/persistence/asset/session owner；释放资源 | Pass |
| RT-02 | subject 切换且有工作簿与恢复记录 | 释放旧 owner；journal 按 subject 隔离，不能跨身份重放 | Pending |
| HUB-01 | 创建→目录→打开→编辑→保存→刷新 | 服务端为唯一 authority，版本/快照一致；错误不变本地空表 | Pass |
| HUB-02 | 重命名/复制/移入文件夹/星标后再打开 | SDK 生命周期动作与服务端/目录一致，无双写 | Pending |
| HUB-03 | 删除→恢复；再次删除→永久删除 | 权限与转换正确；未删除或非 owner purge 拒绝 | Pass |
| HUB-04 | 快速切换目录/搜索/文件夹并取消旧请求 | 旧响应不覆盖新投影；folder cycle 可观察错误 | Pending |
| DOC-01 | 含公式/样式/DV/CF/table/objects 的 XLSX 导入编辑导出 | 显式 capability，唯一 owner，未知 parts 保真 | Pending |
| DOC-02 | XLSM/VBA/chart/pivot/OLE 等 preserve-only 文件 | 操作前可查 editable/preserved/unsupported/host-owned；破坏性编辑原子拒绝 | Pending |
| DOC-03 | XLSB/XLS/ODS/XMLSS/CSV/Text/SJS/SSJSON 实文件往返 | 逐格式区分读/编辑/写/保留，不以 preserve 计实现 | Pending |
| DOC-04 | 桌面 Excel 打开导出文件、重算、保存、重新导入 | 无修复提示，公式/数据/对象正确；无桌面 Excel 则 Blocked | Blocked |
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
| BOUND-01 | 扫描 Web import/API/Worker/model/command/token | 仅 sdk/ui-system 业务依赖，无内部路径/HTTP/worker | Fail |
| BOUND-02 | 注入越层 import/相对路径绕过/未声明依赖/cycle | 门禁全部拒绝；正常 graph 通过 | Pending |
| GATE-01 | build/typecheck/unit/boundaries/contracts/Java/browser/PR checks | 保存实际结果；全部通过后 ready 并 merge 指定 head | Fail |

## 用户追加的对象模型与多工作簿目标（实施前设计）

详见 sdk-object-model-architecture.md。本轮先冻结所有权、生命周期、版本、权限与对象 API，再一次实现完整批次，最后统一验证；完整 Excel 对齐不以入口类或少数公式通过代替。

| ID | 前置条件与操作 | 预期结果与拒绝路径 | 状态 |
|---|---|---|---|
| OO-01 | SDK → Workbooks → Workbook → Worksheets → Cell/Range；真实显式寻址、typed 值/公式、保存与关闭 | 同一 canonical 数据/权限/计算/写链；不暴露可变 model/session/API；非法地址、失效、越权 typed 拒绝 | Pass |
| OO-02 | 按所有 Excel/Aspose 类对象逐能力调用、编辑、保存、重算、往返 | 完整领域对象与 read/edit/write/preserve capability，不能只有 Web 菜单可用 | Pending |
| OO-03 | 浏览器无编辑器、Node/Java 文件/Worker/输出宿主 | 宿主仅 I/O/线程，SDK 拥有语义与文档；缺失宿主明确拒绝 | Pending |
| MWB-01 | 一个 SDK 同时打开多个 workbook；重复打开；关闭一个、subject 切换、全部释放 | 按 unitId 唯一 owner、对象与租约；不影响其他 workbook，无旧身份读写 | Pass |
| MWB-02 | 真实跨 workbook SUM/范围/其他函数，source 更新→提交→refresh→target 重算与保存 | source ID/sheet ID/subject/revision/access 定义明确；撤权清缓存、#BLOCKED!、恢复重算 | Pass |
| MWB-03 | 多级依赖图、环/迭代、自动传播、版本变化/关闭/撤权并发 | 统一图、明确循环语义、无过期授权/过期任务结果 | Pending |
| MWB-04 | 多 workbook copy/move/批量写/undo，任一权限/版本/存储失败 | 服务端共同事务提交或整体拒绝；不能将依次保存声称原子事务 | Pending |

本批细项 OO-01.a/b、MWB-01.a、MWB-02.a/b 的完整成功/拒绝步骤已在架构文档中预先制定；执行结果在统一验证后记录。原有 42 项保持完整；追加 7 项后共 49 项。

## 已确认入口与删除要求

- ApplicationServicesProvider 创建 API/persistence/assets/session options，需迁入 SDK。
- Web auth/session.ts 与 oidc.ts 拥有 fetch/CSRF/UserManager，迁移后删除原文件。
- protocol WorkbookAclRole、permission ShareRole、persistence WorkspaceRole、Web WorkbookRole、Java enum 重复定义，统一生成契约。
- workbook-session.ts 8204 行，需领域拆分，不得新增 facade 后宣称完成。
- Web column-dimension-controller/autofit worker/command-controller/canvas interaction 需逐域迁移。

## 逐条执行记录

验收范围没有缩减。12 个已执行的浏览器场景不等于上面 42 个完整产品项全部通过；覆盖不足的父项继续 Pending。PR #349 保持草稿，未合并。

### 已通过的可复现细项

在 **34c44dbe025da5d298bc0aedafadbd97731902b9** 的干净源码上运行 `SDK_UAT_EVIDENCE_DIR=/tmp/sdk-uat/evidence npm run test:sdk-uat`：**12/12 Pass，1.5 分钟**。真实 Chromium + Java 21/H2，无 HTTP route mocks。每次使用全新临时 H2，结束后关闭 Java/Vite。源码/运行记录由现有 provenance gate 验证。脚本为 `frontend-react/e2e/sdk-product.spec.ts`，每项都检查浏览器 console、pageerror 和失败请求；错误密码产生的预期 401 单独校验。

| 细项 | 实际动作与断言 | 结果与证据 |
|---|---|---|
| AUTH-01 / ID-01 | 浏览器初始化管理员；创建 editor/commenter/viewer；禁用→旧会话失效；启用→重新登录；重置密码→旧会话失效 | Pass；identity-users.png |
| AUTH-02 | 错误密码 401 且仍匿名；正确密码登录；登出后刷新保持匿名 | Pass；auth-logged-out.png |
| ID-02 | 三种普通用户不显示管理入口；管理员页面拒绝；真实 `/api/admin/users` 返回 403；SDK 单测拒绝管理员动作且不发请求 | Pass；role-editor/commenter/viewer.png + SDK auth tests |
| ROLE-01 / ROLE-02 | 生成 TS/Java 唯一角色；四种角色的 capability；实际目录菜单；活动和回收站 purge/trash 越权请求带真实 CSRF，返回 FORBIDDEN | Pass；owner-catalog.png、三个角色截图、trash-purged.png；非法角色拒绝单测 |
| RT-01 | SDK 唯一组合根；StrictMode 清理取消与最终释放；销毁后的旧 actions 拒绝；浏览器真正打开工作簿 | Pass；internal/runtime.test.ts + workbook-persisted.png |
| HUB-01 | 新建服务器工作簿，Canvas 输入 A1，保存，轮询服务端快照，刷新后读回同一值 | Pass；workbook-persisted.png |
| HUB-02.a | 重命名、星标/取消星标状态、给三角色设置真实 ACL | Pass；owner-catalog.png；复制、文件夹移动和重新打开尚未验收，HUB-02 父项 Pending |
| HUB-03 | owner 删除→恢复→再次删除→永久删除；非 owner 对活动/已删工作簿 purge 均被服务器拒绝 | Pass；trash-purged.png |
| SIZE-01.a | 多列宽度设 0 隐藏，设 12 原子取消隐藏；Undo/Redo；20pt 行高；含文本/空列 AutoFit；刷新后的服务器尺寸 | Pass；sdk-dimensions.png；手势/全部宿主路径未覆盖，SIZE-01 父项 Pending |
| SIZE-01.b | stale AutoFit、取消、worker 失败及非法批量尺寸无部分写入；Worker 被释放；缺少 Worker 宿主明确拒绝 | Pass；SDK dimensions + internal/runtime tests |
| DATA-01.a | 真实 Canvas 输入数据；降序排序→Undo 完整单元格恢复；worksheet filter 开/清/关；创建 Table，filter off/on 与 autoFilter owner 同步；clear 保留按钮 | Pass；sdk-data-undo.png；manual hidden/outline 全部语义未完成，DATA-01 父项 Pending |
| DATA-01.b | 删除重复行后服务端快照正确，Undo 精确恢复整个 worksheet cells | Pass；sdk-data-undo.png；Java 拒绝范围外、其他 sheet、篡改内容与非 restore mutation |
| DATA-02.a | SUM 分类汇总公式与 outline；Undo→Redo→再次 Undo，精确恢复空白目的地与 absent outline | Pass；sdk-data-undo.png；全部聚合及 advanced filter/DV/CF 尚未完成，DATA-02 父项 Pending |
| EDIT-01.a | 分列覆盖已占用/空目的格，Undo 完整恢复分列前 cells；刷新后服务器快照一致 | Pass；sdk-data-undo.png；完整编辑父项仍 Pending |
| AUTH-04.a | 未知模式、坏响应、网络失败撤销身份；可重试；snapshot 不泄露 credential/CSRF；销毁后不发布 | Pass（SDK 单测）；浏览器故障注入未完成，AUTH-04 父项 Pending |
| STRUCT-01.a | 109 项结构/引用/尺寸 focused checks：隐式 formula rule anchor presence、歧义身份拒绝、不可变 owner carriers | Pass（单测）；所有引用 owner/客户端服务器完整一致性未完成，STRUCT-01 父项 Pending |
| COLLAB-02.a | 重复 Undo/Redo 使用新 durable operation 身份；旧/重复 binding 拒绝；多段删除按每段 preimage 恢复公式 | Pass（真实 DATA 场景 + history/Java 单测）；远程交错重放等未完成，COLLAB-02 父项 Pending |
| DATA-03.a | 普通/排序 block 预取不自动重试已失败 block；显式 retry 可成功；校验/长度错返回 typed storage error | Pass（19 项 focused checks）；其他 source/query/linked 场景 Pending |
| DOC-01.b / REVIEW-01.a | 先上传越界 hyperlink ref 的 XLSX，typed 拒绝且没有创建服务器工作簿；同一对话框重试合法文件，打开、键盘全选替换 A1、Ctrl+S、刷新、目录下载、解析真实下载 XLSX，四种 target 与值保留 | Pass；native-hyperlinks-edited.png 与 native-hyperlinks-edited.xlsx；DOC-01/REVIEW-01 完整父项仍 Pending |
| DOC-01.c | 实际插入 48×32 PNG 浮动图片、提交引用操作后读取资产、保存→服务端快照→刷新→目录导出→验证下载 ZIP 中 media 原字节/relationship/重新导入 drawing | Pass；first-image-persisted.png、first-image.xlsx；SDK 缺失资产/错误 metadata/hash 拒绝；H2 HTTP 上传成功/400/403/415 拒绝；完整父项仍 Pending |
| CHART-01.a | 自定义 view3D capability preserve-only、原字节保存；编辑后 typed 拒绝；合法默认 3D 饼图 codec | Pass（OOXML 单测）；完整 chart/browser/Desktop Excel 仍未通过 |
| DOC-01.b identity | 原生数字 sheetId 与业务 ID 分离；一次映射所有 native owners；自定义/交换身份、原生 sheet 重排、保留数字 ID、v3 导入边界升级、重名/重复原生映射拒绝 | Pass（2 个原生 focused tests，均含多项成功/拒绝断言）；完整 native corpus/Desktop Excel 仍不等同通过 |

以上截图与日志位于 `/tmp/sdk-uat/evidence/`；runner 汇总 `/tmp/sdk-uat/result.log`。临时文件不提交进仓库，CI 将日志/截图/provenance 打包为 artifact。

### 真实发现与修复

1. 工作簿在 collaboration 初始同步完成前不能进入 ready，否则浏览器操作遭到离线结构授权拒绝。
2. 204 响应必须完成 response body 消费，避免成功操作产生真实网络 aborted 错误。
3. absent 尺寸覆盖与 outline 不能被 Undo 写成默认值或空对象；canonical row/column resize 使用 nullable 删除覆盖，outline.set null 删除 absent outline。
4. 分列等 range.set 的逆操作必须覆盖所有写入坐标，含原本空格。
5. Redo 是新的 durable operation；再次 Undo 必须指向新提交的 id/base revision。
6. 纯结构删除的逆操作需允许原删除范围内、值完全一致的 cell.restore；用 canonical committed replay 还原每一删除段的前置状态，最后继续验证整个 preimage 精确相等。越界或篡改依旧拒绝。
7. 已失败 block 不能因 Canvas 预取循环不断变成 loading；显式读取才重试。
8. Ribbon 菜单的 React key 按 canonical menu id 设置，消除真实浏览器 console 警告。
9. 原生 XLSX 的 dimension 不包括合法空白格引用，在导入边界扩展 hyperlink anchors/targets 的 canonical extent；越界仍 typed 拒绝。
10. 导出以前把任意业务 ID 写进原生 sheetId，导入只重映射单个 sheet，造成跨表引用丢失。现在统一分配/保留合法数字 ID，用 metadata v4 明确绑定业务 ID，一次重映射 native owners；v1–v3 只在显式导入边界升级。重复身份在 export 前置检查拒绝，不等到图形或链接序列化时才失败。

11. 首次图片 drawing 根节点为空自闭合时，原有关闭标签替换无法插入 anchor。改为唯一 XML tree parser/serializer，校验 native drawing namespace/根节点；保留未知节点与属性，非法包 typed 拒绝。
12. 目录导出缺少 SDK 既有 AssetStore；现在使用同一个 unit/subject owner。上传二进制 content-type、MIME header 和 nullable metadata 契约已与 Java 对齐；错误媒体类型/缺失 header 由服务器分别报告 415/400。
13. 图片引用未 ACK 时，Canvas 资产读取被服务器正确拒绝 403。现在读取遵守既有 canonical queue commit barrier，未连接/拒绝返回 AssetReferenceError；没有放宽服务端授权或增加读取重试。
14. 原生 chart/view3D 查询漏掉 chartSpace 根，导致自定义视角丢失。现在从原生合法树读取并报告 preserve-only，未编辑保留原 bytes，编辑拒绝。
15. Pivot detail Undo 删除 worksheet 前，先在同一 canonical mutation 撤销该表专属 region/source，保留外部引用检查；没有忽略删除检查。真实浏览器和全部服务端 parity 仍 Pending。

每个已完成功能或实际缺陷修复均为单独 commit 并推送同一个草稿 PR；没有提交到 main。

### 门禁记录与未通过项

- `npm run build`：Pass（TypeScript + Vite）。
- `npm run test:sdk`：36/36 Pass；auth/identity/runtime/dimensions/data 成功与拒绝路径。
- protocol focused：33/33 Pass。
- `npm run check:boundaries`：Pass（包含 generated contracts/registry/stack/provenance/既有 acceptance matrix）；`npm run test:calculation-domain`：446 + 5 tests Pass。不能据此宣称 Web 已只依赖 SDK。
- `npm run test:native-codecs`：17/17 Pass；`npm run test:pointer-gesture`：7/7 Pass。
- OOXML suite：**63/63 Pass**；首次图片 DrawingML、合法/非法 preserved XML、业务/native sheet identity、自定义原生 3D 饼图保留与编辑拒绝。Camera typed error 与 Pivot OPC ContentTypes fixtures 已按真实契约修正。codec 通过不等于完整原生互通验收。
- 资产引用提交屏障：2/2 Pass，真实 CollaborationSession 的 ACK 前不 GET，ACK 后读取；离线/拒绝不 GET 且队列状态保留。Pivot detail lifecycle：21/21 Pass，完整恢复和 region/hyperlink 拒绝无部分状态。
- Java 21 `mvn ... package`：实际 Maven 汇总 **321 tests，0 failure/error/skipped**，真实 H2 integration；不使用含过期报告的 XML 目录总数。
- 完整 `npm run test:unit`（本次 Pivot lifecycle 实现，提交前干净语义检查）：**1545 tests，1491 Pass，54 Fail**；上次 ed0f570 为 1543/1487/56。基线为 1525 tests，1442 Pass，83 Fail。不能未经逐项对照认定失败都来自基线；完整失败清单另见 sdk-product-unit-failures.md。
- GATE-01：Fail。CI 真实 SDK UAT 与完整 test:unit 独立执行，完整日志作为 artifact。ed0f570 的 Windows push run 37093921510：12/12 UAT Pass，完整单测 56 Fail；同 head 的 PR run 37093924549 在 DATA 新工作簿打开的默认 5 秒内未 ready（3 Pass / 1 Fail / 8 未执行），缺少失败 trace，原因未确认。34c44dbe 已补首次失败 trace/截图/console/pageerror/网络拒绝记录与递归 artifact，没有增大超时或增加 action retries，后续 Windows 结果待记录。
- BOUND-01：Fail。Web 仍有内部包业务 import；WorkbookSession 仍 8181 行并暴露给 Web，完整 WorkbookHandle/领域拆分未完成。认证、组合根、目录 service、尺寸规划/worker、六个数据 action 已实质迁移，不能据此宣称所有领域完成。
- BOUND-02：Pending。现有 graph gate 通过，不代表更严格的 Web-only-SDK 及绕过注入门禁已经实现。
- RT-02：Pending。subject owner 退休单测通过，但跨身份恢复 journal 的完整验收还没有完成。
- AUTH-03：Blocked。用户明确回复没有真实 OIDC issuer/provider，SDK 已迁入 OIDC 生命周期，但不能把本地认证测试作为 OIDC 验收。
- DOC-04：Blocked。用户确认没有可用桌面 Excel 环境；当前 Linux 环境也没有桌面 Microsoft Excel；没有声称做过 Excel 打开/重算/保存检查。
- 其他 Pending 父项按原始矩阵保留：完整结构/引用图、所有 Excel 领域和 formats、交错协作/history、输入/打印/对象等未完成。

### 下一项实施前的 UAT 设计

**DOC-01.b / REVIEW-01.a 空白格超链接往返（Pass，先制定再执行）**：构造真实 XLSX，只有 A1 有值、A2:A4 为 URL/email/sheet/name 超链接锚点，native dimension 仅 A1；导入需保留四个锚点并在显式 import 边界扩展 canonical extent；编辑 A1 后导出再导入仍保留四个 target，不能创建空格假值。将 hyperlink ref 改成 A1048577 或 XFE1 时必须 typed 拒绝且原输入 bytes 不变。已通过浏览器上传、打开、真实键盘编辑、保存、下载文件验证同一语义链；越界文件明确拒绝且不创建服务器工作簿。导入业务身份的跨表恢复又揭示了原生 sheetId 缺陷，按相同往返要求追加身份成功/拒绝检查。桌面 Excel 此项仍归 DOC-04 Blocked。

**DOC-01.c 首次图片导出（Pass，先制定再执行）**：使用真实 PNG 和对应 AssetRef，首次生成 XLSX；检查 drawing 根节点、图片 relationship、原始 media bytes、ContentTypes 与导入后的 canonical payload。现有空 drawing 的自闭合根、不同 prefix/default namespace 都必须成功，并保留未知元素与属性；错误根 namespace、多个根或未闭合 XML 必须返回带 part/recovery 的 typed error，原 snapshot、asset bytes 与 preserved package 不变。本项只覆盖首次生成和空 drawing 的添加；补充真实浏览器步骤：新建服务器工作簿→插入 PNG 浮动图片→Ctrl+S→服务端 drawing/payload 断言→刷新→目录导出副本→检查真实下载 ZIP 的图片 relationship/media 与重新导入结果，并检查 console/network。目录导出必须使用 SDK 同一个服务端 AssetStore；缺失资产、错误 metadata 或 hash 必须 typed 拒绝，不下载部分文件。真实 HTTP 资产上传用 application/octet-stream，X-Asset-Mime-Type 声明内容 MIME；未提供 width/height 时响应不得包含 null。H2 HTTP 测试需确认正常上传/原字节保存、错误 content-type 415、缺失 MIME 400、哈希不符 400、无授权写 403，无拒绝项落库。新增顺序检查：引用操作 ACK 前不得 GET 资产；ACK 后只读取一次；离线或引用提交拒绝时返回 AssetReferenceError，不发 GET，不绕过 subject-safe asset projection。已有图片的替换/删除/重复导出仍需单独验收，不据单测标记 OBJ-01 或 DOC-01 父项通过。桌面 Excel 仍 Blocked。

**CHART-01.a 原生 3D 饼图视角（Pass，codec 检查，先制定再执行）**：默认 3D 饼图在 canonical native writer/import chain 保持可编辑；将合法原生 chart/view3D 设置为自定义旋转，原生 capability 必须报告 preserve-only、原因可观察。未编辑保存需逐字节保留原 XLSX，编辑后再导出必须 NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED 拒绝且不修改原 native parts/snapshot。XML 中 view3D 必须位于 chart 下、与 plotArea 平级；codec 检查不代替真实浏览器/桌面 Excel 的 CHART-01 完整验收。

**PIVOT-01.a 明细表撤销的引用生命周期（Pass，CommandRuntime 成功/拒绝检查，先制定再执行）**：用真实 canonical Pivot 命令创建 block-backed 明细表，Undo 同一个事务必须依次撤销 region binding、专属 data source 和 sheet，不能以忽略引用检查实现删除；原 sheet/source/region 快照需完整恢复，Redo→Undo 仍相同。其他 sheet region/Pivot/table 或超链接引用明细时，撤销必须拒绝且 worksheet、source、region 与 history 无部分变化。此项先做 CommandRuntime 成功/拒绝检查，公式删除的 canonical #REF! 改写及其 history、真实浏览器和服务端完整 parity 尚需另行验收，PIVOT-01 父项保持 Pending。

### 合并与回滚

只有所有必需产品项 Pass、完整门禁 Pass、PR checks 对应待合并 head 验证后才 ready/merge；现在没有达到条件。

没有生产部署，验收只操作临时 H2。尺寸 null 和 outline null 是 canonical operation 语义变更，前后端必须同版本发布/回滚；出现新语义的 durable 日志后，单独回退代码不可安全读取，应使用兼容版本或恢复匹配的 operation/checkpoint 备份。原生 metadata v4 是另一个文件契约变更：旧版应用不能读取新生成的 v4 文件，回滚须保留当前 codec 或从原始文档/兼容备份恢复，不能只回退应用。二进制资产 HTTP MIME header/content-type 同样要求前后端同步发布/回滚，不提供 runtime 旧 header alias。旧 Web auth/service/尺寸 worker 等已删除，不保留兼容桥。schema upgrades 只在显式 migration/import boundary 执行。


## 对象模型与多工作簿：首轮验证记录（2026-10-03）

产品代码 head `5ce6111179f7fdbe687cbbb2a7fc60d73af4ecbf`。统一实现后 build/typecheck、boundaries 通过，SDK 成功/拒绝检查 36/36 Pass。完整 unit 1551 tests /1497 Pass /54 Fail；与上一 head 的失败标题逐条对比，新增 0、消失 0；GATE-01 仍 Fail。

首轮真实 Java/H2 UAT：6 Pass /1 Fail /6 未执行。OO-01/MWB-01/MWB-02 用例在加载 SDK 时即失败；`/packages/sdk/src/index.ts` 返回 404，业务场景没有执行，相关父项保持 Pending。原始 trace 同时证明 Web 正常消费的公开入口 `/@fs/.../packages/sdk/src/index.ts` 返回 200。首次失败证据保存在 `/tmp/sdk-oo-uat/first-attempt-test-results`，不以重新运行覆盖首次失败。

验收宿主修正边界：依据 package.json 的 `exports["."]` 使用 Node 的 `import.meta.resolve('@react-sheets/sdk')` 定位公开入口，转换为 Vite 的实际 filesystem module URL；保留真实 Java/H2、浏览器同源认证、无 route mocks、零 retries 和原断言。只调整验收消费入口，不修改冻结的产品实现或 Vite 配置；Linux/Windows 路径均在进入浏览器前归一为 URL。修正后执行真实 UAT 并分别记录结果。


真实授权补充 UAT 的首次独立执行（`7e0acdba118a26c9c009feae2356cd3fd29ba6fb`）：身份 A→B 的旧 Workbook/Cell 退休、重新打开读取已提交 43、active SDK dispose 退休旧 Cell 全部 Pass。五函数/撤权用例在 bootstrap 前置步骤返回 403，业务断言未执行。trace 的两个匿名 session 响应证明：bootstrap 使用第一个响应的 CSRF，却携带第二个响应写入的 session cookie；Web 和独立 SDK 的初始化竞争。没有输出 token/cookie，也不放宽服务器 CSRF。

验收宿主修正为 Vite 提供的独立 `e2e/support/sdk-consumer.html`，只启动被测公开 SDK，不挂载 Web 的另一个认证组合根。这符合预设“无编辑器 SDK consumer”前置条件；页面保留真实 origin/HttpOnly cookie/Worker。首次失败保留在 `/tmp/sdk-mwb-authority-uat/first-attempt-test-results`，产品代码仍冻结于 `5ce6111179f7fdbe687cbbb2a7fc60d73af4ecbf`。


独立 SDK consumer 的原始 filesystem HTML 尝试在 React dev preamble 校验时失败（2 Fail，业务未执行；head `425519666f8324506c11a1a8181aa2e97b163c2c`）。`/@fs` 静态 HTML 没有经过 Vite 的 HTML/React transform。宿主最终设为 Web 开发根中的 `sdk-consumer.uat.html`，由实际 Vite HTML 管线处理；不手工置入成功标志或省略 React 前置校验，不挂载 Web。生产构建仍只有既有 index.html 入口；该文件仅是开发服务器的验收入口。证据保存在 `/tmp/sdk-mwb-standalone-uat/first-attempt-test-results`，产品实现继续冻结。


## 本批最终逐条验收（产品实现未随验证修改）

产品实现固定于 `5ce6111179f7fdbe687cbbb2a7fc60d73af4ecbf`。后续提交只改变验收入口/宿主、预设场景及证据文档；`git diff 5ce61111 -- packages/sdk/src packages/spreadsheet-app/src ...` 的产品路径为空。基本对象 API 的 Pass 不代表完整公开 SDK 边界已完成；BOUND-01 仍 Fail。

| 细项 | 实际执行与结果 | 状态/证据 |
|---|---|---|
| OO-01.a | Worksheet ID/name/index、稳定 cell、显式 D8 对比 A1 选区、矩形惰性遍历、非法 A0/XFE1/Excel 越界、非法 scalar/formula 不改变 snapshot；Workbook 无 session/model/API/凭证 | Pass；36 SDK tests（新增 6 组）；基本对象入口父项 Pass |
| OO-01.b | typed 字面量 `=1+2`、明确 SUM 公式和计算值、style 保留；viewer 写/hidden-range 读 typed 拒绝；隐藏行列 canonical read；保护后 formulaHidden 不返回公式、结果仍为 3、snapshot 不变 | Pass；SDK tests + `/tmp/sdk-oo-remaining-read-uat.log` 的 3 项语义检查；完整 ACCESS-01 保持 Pending |
| MWB-01.a | 同 SDK 同时开两本、重复打开相同对象、关闭来源后目标仍可编辑；打开中的 ready Promise 共享且 close typed 拒绝 | Pass；真实 13 项产品 UAT + 补充语义检查 |
| MWB-01.b | 真正从用户 A 切到 B，旧 Cell RUNTIME_DISPOSED，新 Workbook 不是旧对象且读取已提交 43；active Workbook 下 dispose 后旧 Cell RUNTIME_DISPOSED | Pass；`sdk-workbook-acceptance.spec.ts` 第二项，真实 Java/H2、公开 SDK、独立宿主、无 HTTP mocks；完整跨身份未提交 journal 的 RT-02 仍 Pending |
| MWB-02.a/b | 实际 source 10/20 → SUM 30；source 改 40 并 flush/refresh → 50；保存/服务端 snapshot/重新打开/真实 XLSX 的公式、字面量与 source ID 一致 | Pass；本地与 Windows `124b2359` 13/13 产品 UAT，`sdk-two-workbooks.png` 明确显示 50；仅直接来源 |
| MWB-02.c | 真正普通用户 source viewer/target editor；SUM/AVERAGE/COUNT/MIN/MAX 初值 30/15/2/10/20；更新来源后 50/25/2/10/40，sourceRevision 增加 | Pass；`b9c24435` 的真实独立 SDK UAT，均在后续拒绝断言之前通过 |
| MWB-02.d 拒绝 | owner 真正删除 source ACL；external inputs 403、链接状态 denied、公式文本不变；COUNT 返回数字 0，而不是预设 #BLOCKED! | **Fail**；本地原始 trace `/tmp/sdk-mwb-vite-host-uat`；不弱化断言、不修改冻结的产品实现；MWB-02 父项 Fail |
| MWB-02.d 恢复 | 恢复 source viewer 后应重算为授权的版本 | Pending；前述断言失败，恢复步骤未执行 |
| OO-02/OO-03/MWB-03/MWB-04 | 所有 Excel/Aspose 对象与格式；Node/Java 宿主；完整依赖图/环/自动传播；跨簿原子写/history | Pending；不据少数函数或浏览器宿主通过宣称全部实现 |

当前完整 49 项父项：**11 Pass /3 Fail /2 Blocked /33 Pending**。原 42 项仍为 9 Pass /2 Fail /2 Blocked /29 Pending；新增 7 项为 2 Pass /1 Fail /4 Pending。SDK 36/36、build/typecheck、现有边界检查通过；完整 unit 为 1551 /1497 Pass /54 Fail，失败标题与上一产品实现逐条对照没有增减。

真实 UAT 按 head 分别记录，不混成一次绿色结果：`124b2359` 本地/Windows完整产品 13/13 Pass；补充独立宿主 `b9c24435` 本地 1 Pass /1 Fail，COUNT 业务拒绝问题。Windows `7e0acdba` run 37100228629 原产品 AUTH-01 的 viewer 登录前置等待超时（原因未确认），其余串行产品用例未执行；补充用例失败也保持可见。最终 Windows `b9c24435` run 37101187825/job 111140878728：**14 Pass /1 Fail**，COUNT 同样返回 0；完整 unit 仍 54 Fail。Java/build/SDK/calculation/boundaries 通过，真实 UAT 与完整 unit 门禁失败。不得隐藏前述失败或标记全部门禁通过。

后续修正所需的架构边界：授权/输入不可用属于求值前置故障，应与普通 Excel cell error 区分，并沿同一输入解析、函数/依赖计算与 Worker 结果链传播；COUNT 对合法源里的普通错误忽略行为应保留。不能用 COUNT 专用补丁、UI 改显示、保留旧缓存或放宽 403 来完成。需要先确定所有读取/派生路径、缓存退休/恢复与 Worker/服务端约束，再一次性实施完整修正批次。当前没有满足合并条件。


## C1 授权输入前置故障验收（2026-10-03）

产品 head `70c3920499b50ab6fa636662f00582ddc16d9e47`。单一输入边界观察实际消费的范围、数组、稀疏迭代、名称和派生值，传播带 reason/source 的输入 fault；普通 Excel 错误值保留函数语义。没有针对 COUNT 改写结果或削弱服务器授权。

干净 head 的真实 Java 21/H2/Chromium 补充验收 **2/2 Pass**（`/tmp/sdk-input-fault-browser-clean.log`；`/tmp/sdk-product-uat-7cIKtM/evidence`）。五函数用例实际完成初值、源更新、普通用户权限、撤权 403、五个 #BLOCKED!、恢复 ACL 和授权版本重算。subject 切换/active dispose 用例再次通过。MWB-02.c/d 当前全部 Pass，父项 MWB-02 改为 Pass；多级图与跨簿事务仍分别 Pending，不扩大本项结论。

新增 4 组成功/拒绝检查覆盖 12 种消费表达式（含 IFERROR/ISERROR/AGGREGATE）、三种输入故障、Worker 和派生公式、权限恢复、隐藏范围、audit trace、未执行 IF 分支及损坏元数据拒绝。calculation-domain **455/455 Pass**（主套件 450 +补充 5）；SDK **36/36 Pass**；build/typecheck 和现有 boundaries Pass。完整 unit **1555 /1501 Pass /54 Fail**，与先前 head 对比失败标题新增 0、消失 0。GATE-01 和 BOUND-01 继续 Fail。

当前完整 49 项父项为 **12 Pass /2 Fail /2 Blocked /33 Pending**。旧 COUNT 失败与恢复未执行记录是旧 head 的历史证据，完整保留。首次本批 browser 被 clean-source 前置检查拒绝，没有执行业务步骤；提交冻结实现后重新执行，没有绕过门禁。GitHub 当前没有返回该 head 的 PR workflow run，不把未执行的检查计为通过。PR #349 继续草稿，尚不满足合并条件。

## I1 身份上下文执行记录

产品 head `1100d699cc8f8d724fcf367dde3e3276666e4af3`：build/typecheck、现有 boundaries Pass；SDK 42/42 Pass；Java 327/327 Pass + package；完整真实 Java/H2/Chromium **15/15 Pass**，证据 `/tmp/sdk-identity-browser.log` 与 `/tmp/sdk-product-uat-GLepoX/evidence`。包括旧对象退休、当前身份重新打开、多工作簿 source 撤权和恢复，以及现有 UI、持久化和原生输出场景。

新增凭证检查覆盖服务端确认身份、并发单次获取与确认、同上下文续期、不变 subject 下换工作区、旧 HTTP/响应体拒绝、过期/401 清除身份、host CSRF/正式 prefix、来源不匹配与坏上下文。Java 以实际 RSA 签名/JWKS decoder 检查 issuer/audience/过期/错签名/缺 sid 拒绝，并验证不同 authority/scope 主体隔离、ACL 不能越过 scope、同 context 续期与显式可信身份映射。真实外部身份、ERP 路由和上下文恢复尚需 I2/后续验收，相关父项不升级为 Pass。

产品源在统一验证期间未修改。首轮 Java 只有正式 schema 已迁到 V14、迁移测试仍断言 13 的契约陈旧问题；同步版本并增加真实 v13→v14 旧数据保留和重复启动检查后，Java 全部通过。完整 unit 仍 1555/1501 Pass/54 Fail，失败标题与 C1 相同。49 父项保持 12 Pass/2 Fail/2 Blocked/33 Pending；现有 boundaries 的绿色结果仍不代替 BOUND-01。PR 继续草稿。

## C2 多簿图统一验证记录

本批采用实施前冻结的 MWB-03.a/b/c/d 细项，删除旧单 binding 输入端点与消费者，统一版本图、前置循环校验、来源变化订阅及递归计算。V15 仅增加 topology 事务门控行。64 簿、单节点 100000 输入和总图 1000000 输入为显式预算；流式 block 来源当前 UNSUPPORTED_FEATURE，完整 R1 仍待实现。

SDK 42/42、calculation-domain 457/457（452 主域 + 5 结构/引用域）、Java 331/331、build/typecheck 与 boundaries Pass。新增测试覆盖 A→B→C 的 20/1/20、叶子拒绝三式均 #BLOCKED!、恢复为 80/1/80、Worker 快照同结果、坏图/预算拒绝，以及并发相反绑定仅一个写入、另一操作无 history/版本写入。日志 `/tmp/sdk-graph-sdk.log`、`/tmp/sdk-graph-calculation-current-fixture.log`、`/tmp/sdk-graph-java.log`、`/tmp/sdk-graph-build-current-fixture.log`、`/tmp/sdk-graph-boundaries.log`。首次前端验收 fixture 调用了不存在的 Worker 快照方法，原始失败日志保留；按真实 exportCalculationSnapshot/fromCalculationSnapshot 契约同步后通过，产品实现未改。

全量 unit 1557/1503 Pass/54 Fail；与 I1 失败标题新增/减少均为 0（`/tmp/sdk-graph-all-unit.log`）。真实浏览器的自动传播、关闭来源对象、撤权/恢复、重开和 SDK 循环拒绝此刻尚未执行，MWB-03 保持 Pending，下一步在已提交的干净 head 执行。全计划与合并门禁仍未完成。

C2 首轮干净 head b8a9fbddffac618d63c6a285aee31e134f2fb675 的真实浏览器 **7 Pass/3 Fail/6 未执行**；三簿数值、服务器通知自动请求、撤权恢复、保存重开与循环前置拒绝已逐条通过，但意外图 GET 409 使 console/network 检查失败，不能报整项通过。原始日志 `/tmp/sdk-graph-browser.log`、trace `/tmp/sdk-graph-first-attempt-test-results`、后端证据 `/tmp/sdk-product-uat-nNd04n/evidence`。实际 409 为 stale ORM entity 升级锁时的乐观版本冲突；完成整体锁/ORM/拓扑入口审查后，统一更换为规范内部闭包锁集合、锁前 flush/clear、授权响应重新读取，并补齐 history restore 与 source 生命周期。

修正批次 Java **334/334 Pass**（`/tmp/sdk-graph-capture-java-current-fixture.log`）、typecheck/boundaries Pass。真实两事务旧 entity 后源提交 40 的读图验证通过；循环历史恢复零快照/版本变化、trash/restore/purge 无失效来源 snapshot、订阅生命周期检查通过。首轮新增来源写入 fixture 未携 canonical writeAuthority，服务端正确拒绝；按既有写入契约补全测试输入后通过，修正批次产品源保持冻结。浏览器待此批提交后执行，当前完整 C2 验收仍未通过。

e8bf2736763c2b74793ed0b4f9e2e7e63dd695be 的第二轮真实浏览器 **15/16 Pass**（`/tmp/sdk-graph-capture-browser.log`、`/tmp/sdk-product-uat-9Jh5qs/evidence`）。初轮图 409 全部消失；新增依赖生命周期 #REF!/恢复/永久删除断言全部通过。末尾 network 断言发现仍打开的来源自身在 WS 关闭后读取 access/snapshot 403，不能忽略。原始 trace `/tmp/sdk-graph-capture-attempt-test-results`；完成 lifecycle→Session disposal→公开对象退休→runtime/cache release 的整体契约批次。生命周期消息先于初始同步的 deferred messages 处理，晚到快照不发布；消息只含 unit/lifecycle，无来源数据，客户端不能发送。

对象退休批次统一验证：SDK **44/44**、calculation-domain **457/457**、Java **334/334**、build/typecheck 与 boundaries Pass。新增成功/拒绝检查包括 active 不退休、trash 退休 Cell/Range、其他 Workbook 可用、直接 Session dispose 退休、初始同步中 purged 即时退休、不同 unit 不被退休、迟到 revision 999 不发布、坏 lifecycle/客户端伪造拒绝。测试首次使用了未公开的内部同步函数入口，按测试内部真实路径与异步启动契约同步 fixture；产品未为测试新增公开入口。全量 unit **1559/1505 Pass/54 Fail**，与 C2 的失败标题无增减，完整日志 `/tmp/sdk-lifecycle-final-all-unit.log`。产品当前准备提交，完整真实浏览器尚待干净新 head 重验。

最终干净产品 head **5a938162e4326bf56e94c14e88f03b69da0e1d6b** 的完整真实 Java 21/H2/Chromium **16/16 Pass**（`/tmp/sdk-lifecycle-browser.log`、`/tmp/sdk-product-uat-uc0ERN/evidence`）。MWB-03.a/g 逐条完成：20/1/20 → source 40 自动 80/1/80；观察服务器事件触发的图请求先于 Cell.read；来源公开对象关闭后仍传播；撤权无 source snapshot、三式 #BLOCKED!；恢复、保存重开均 80/1/80；循环 SDK 前置拒绝且源 snapshot/revision 不变；来源 trash 后旧 Cell RUNTIME_DISPOSED、C 三式 #REF!；restore 新对象读 40、旧句柄继续退休；purge 自动 #REF!，无多余 access/snapshot 403。只有明确设计的循环绑定 422，console/network 其余全部干净。

Java 成功/拒绝与真实并发检查补齐 MWB-03.c/d/e/f：伪造循环 commit、并发相反绑定、循环 history restore、旧 ORM cache 后独立提交版本、坏图/预算与 Worker parity 均通过。MWB-03 的声明范围现为 Pass；64 簿/输入预算、迭代不支持及 block 来源待 R1 的限制保留，不扩大到完整 Aspose。原 49 父项现 **13 Pass/2 Fail/2 Blocked/32 Pending**；BOUND-01/GATE-01、54 项全量单测失败和真实 ERP/桌面 Excel 阻塞均保留。全部计划尚未完成，PR 不合并。

## O1.1 对象操作批次（浏览器前）

整批交付一个 canonical sheet.cells.commitMatrix、Range 读取/写入/样式/清空/填充/合并/复制值，以及 Worksheet 身份/结构/尺寸/可见性/pane 与 Workbook typed undo/redo。10000 单元格预算与显式 extent，数据/记录 owner、真实结构规划器缺失均明确拒绝。完整 O1/O2/T1 和原 49 项 Pending 不因子批交付改变。

首次 SDK 51/52 中唯一失败是测试入口拼写，保留 /tmp/sdk-objects-sdk.log；build 唯一闭包类型错误在 /tmp/sdk-objects-build.log。冻结差异 /tmp/sdk-objects-first-pass.patch，完整异步契约审查后统一修正 source/query 前置与调用输入捕获，不新增 DV 别名或放宽断言。最终 SDK **53/53 Pass**，build/typecheck 和 boundaries Pass；计算域 **457/457 Pass**。日志 /tmp/sdk-objects-final-sdk.log、/tmp/sdk-objects-final-build.log、/tmp/sdk-objects-final-boundaries.log、/tmp/sdk-objects-calculation.log。O1.1 实施前 a..h 仍等待真实浏览器逐条执行；单测中的本地结构成功不以假连接代替，明确验证 STRUCTURAL_PLANNER_OFFLINE 无写入。

### O1.1 首轮真实执行与完整所有权修正

产品 head 83261a156c45316e3c6dab60e8098ff5eaf9a047：19 项真实 browser **17 Pass/2 Fail**。矩阵/格式/历史/值复制步骤到保存均执行，关闭全部对象后的 open 失败；结构场景在真实 Java rename operation 201 后客户端不确认而失败。viewer 整块拒绝与公式错误值复制不写入的独立场景 Pass。保留 /tmp/sdk-objects-browser.log、/tmp/sdk-objects-first-browser-test-results、/tmp/sdk-product-uat-wc2OyW/evidence。没有因 HTTP 201 将 rename 验收计为成功，也没有重试写入或放宽 console/network。

完整链路审查：SDK 缺少根 lease，最后 child release 触发旧 catalog 退休；rename 手工分支改写 cell 公式却不输出它们的 owner deltas。修正 SDK 根 lifetime（dispose/身份变化仍退休），删除 200 余行重复 rename 代码，统一 planWorkbookFormulaRewrite，同时保留 preserve-only source 的拒绝。新增 root catalog 生存/身份退休与 cell/CF/DV/name facts、undo/redo、恶意 server fact 拒绝检查。产品修正后冻结，只有测试的分页 URL fixture 与 readonly fact 构造同步正式契约；未改产品别名或断言。

修正批次 SDK **56/56 Pass**（/tmp/sdk-objects-ownership-final-sdk.log），build/typecheck、boundaries Pass（/tmp/sdk-objects-ownership-final-build.log、/tmp/sdk-objects-ownership-final-boundaries.log），计算 **457/457 Pass**（/tmp/sdk-objects-ownership-calculation.log）；完整 unit **1570/1516 Pass/54 Fail**，失败标题增减均 0（/tmp/sdk-objects-ownership-unit.log）。下一干净提交重新执行整套 browser，O1.1 a..j 未执行成功的细项仍 Pending。GitHub 本次 head workflow runs/statuses 返回空，不宣告线上 CI 通过。

### O1.1 修正批次 k..o（实施前细项）

| 细项 | 成功验收 | 拒绝验收 | 初始状态 |
| --- | --- | --- | --- |
| O1.1-k | range.clear(contents/all) → undo/redo 后 SUM/COUNT 的值和输入恢复；copyValuesTo 取得恢复后的值；inline/Worker 使用同一输入 journal | viewer 清空/撤销拒绝且源和依赖值不变；manual 模式不擅自自动求值 | Pending |
| O1.1-l | forward/inverse 注册显式 calculation；输入、过滤/outline 行可见性、merge/table spill 几何、find.replaced 和 workbook restore 同步原 owner | 缺失声明、非法枚举、未知字段/错误 context 注册失败，不能通过名称补行为 | Pending |
| O1.1-m | frozen/split pane 的 activePane 缺失与显式值分别经过 snapshot/load/duplicate 精确保留且互不共享可变对象 | 非法 pane/额外字段拒绝，不归一化成可接受状态 | Pending |
| O1.1-n | 真实 Java/H2 删除工作表 → 撤销 → 重做 → 重开；完整 snapshot 与原 preimage 一致，成功 ACK | 对不相等的恢复候选保持 UNDO_RESULT_MISMATCH；不得弱化服务端 guards | Pending |
| O1.1-o | main/security 合并后的新 Jar + frontend 统一检查与真实 19 项 browser，记录 head/console/network | 不以旧 Jar、旧 head 或测试 mock 的成功抵消失败；真实 SSO/桌面 Excel 仍 Blocked | Pending |

第二次真实 UAT e435d678 是 17/19：catalog 最后 child 关闭后能重开；rename 的 canonical owner facts 可被服务端认可。新失败是 k/n，上述两项保持 Pending。只读 replay 文件 `/tmp/sdk-undo-expected-preimage.json` 与 `/tmp/sdk-undo-candidate.json` 的唯一差异为 pane.activePane；所有源代码在分析期间冻结。

首次 calculation 声明批次：SDK 57/58 Pass，新增自动 clear/undo/redo/copy 成功；Worker 手动例失败。完整 unit 1577/1522 Pass/55 Fail（新增同一 Worker 失败，旧 54 标题不变）。只读 probe `/tmp/sdk-clear-worker-proof.log` 证明 null inputs 已正确进入 journal，但 bootstrap pendingRoots 为 []，Worker completed 空结果，源值仍为 1；缺陷在 pendingCalculationRoots 从 live cells 反查地址而过滤已删除的根。脏根 canonical key 自身包含完整地址，导出与 collectAffectedFromRoots 改为同一无回退的 key→address 解码 owner；不再依赖 live cell 存在。新增 k 的手动 Worker 断言保留 1→0→1→0 与 3 次真实 protocol task，不改期望值。

本次批量元数据迁移误改两处动态 StructuralTransformResult（sheet reorder / sheetTable range facts），导致 typecheck 失败；纠正到原 effect 字段，不更改它们的语义。SDK Worker 测试依赖加入正式 devDependency。Java 首次 368 tests/36 context errors/2 skipped，原因是 target/classes 残留旧 V14/V15 SQL；源码只有 V16/V17，新验证使用 clean package 清理编译产物，不删除数据库/改迁移。修正批次完整源码落地后再次统一检查。

修正后 SDK 58/58、calculation 457/457、build/typecheck 和 boundaries Pass。完整 unit 1577/1524 Pass/53 Fail：旧 54 中两项 block-backed AutoFilter sort 通过，新一项结构测试在构造非法 ySplit=1.5 后仍要求 snapshot 成功，违反此批明确的 canonical snapshot 拒绝契约。仅更新 fixture：先保存合法 preimage；非法 pane 的 snapshot 在操作前后都必须拒绝、结构变换同样拒绝且 pane 不变；恢复合法 fixture pane 后比较整个 preimage。产品语义冻结，无断言删除或默认值修复。Maven clean 插件未缓存且当前 JDK 无网络代理 CA，保留原 target 到 `/tmp/sdk-objects-calculation-owner-old-target` 后进行全新 package（不删除数据库，不跳过测试）。

本批统一检查最终：SDK 58/58 Pass；calculation 452+5=457 Pass；build/typecheck、boundaries Pass；Java 368/366 Pass/0 Fail/2 Skipped，package 成功。完整 unit 1577/1525 Pass/52 Fail，与旧 54 标题比较无新增，解决两项 block-backed AutoFilter sort。日志 `/tmp/sdk-objects-calculation-owner-final-{sdk,build,boundaries,formula}.log`、`/tmp/sdk-objects-calculation-owner-final-unit-canonical-fixture.log`、`/tmp/sdk-objects-calculation-owner-clean-java.log`。Java Skipped 条件保留，不能宣告全部实机验收完成。远端 538f3e41 的 canonical-build 两个 check 为 FAILURE（run 37118468043 / 37118464866），不以本地通过覆盖 GitHub 状态；PR 仍草稿。

干净产品 head **0c741cee6bafb3474dccf70452de37d0e6307742**，重建后的 Java21/H2/Chromium 完整真实 SDK UAT **19/19 Pass**，零 route mocks、单 worker/retries=0。证据 `/tmp/sdk-objects-calculation-owner-browser.log`、`/tmp/sdk-product-uat-2Idror/evidence`。a..j 原有业务步骤全部执行成功；新增自动 clear/undo/redo/copy 与 manual Worker journal k、声明成功/拒绝 l、pane snapshot/load/clone/OOXML m 由相应领域检查通过。n 的真实 sheet remove→undo→save→reopen 和服务端严格 preimage 校验成功；其额外 redo→undo 子步骤留至下一轮 UAT，不虚报为全部完成。浏览器 console/network diagnostics 成功（预设拒绝案例按测试明确验证），Worksheet 结构/rename/fill/merge 的操作无异常 HTTP 响应。原 49 父项、T1/O2 全范围、full-unit/CI 等未完成门禁仍保持原实际状态，不据 19 条测试合并。

Java 两项 Skipped 为 QueryReadOnlyDialectIntegrationTest 的 PostgreSQL/MySQL 实机方言用例，缺 REMEDIATION_JDBC_CONFIG，记为外部方言环境 Blocked；H2 权威服务与本批真实 browser 已通过，不将它们算作方言实机证据。

### O2.1 实施前 UAT 细项

| ID | 成功流程 | 拒绝与恢复流程 | 状态 |
| --- | --- | --- | --- |
| O2.1-a | workbook/sheet 同名名称、绝对/相对 anchored 名称、公式修改/移除、undo/redo 与依赖计算；save/reopen 保留 anchor/comment/hidden | 无 scope/坏 ID/非法 anchor、旧 value 别名、viewer 修改拒绝，模型/历史不变 | Pending |
| O2.1-b | 单格与范围富文本保留文字/分段格式，=literal 保持文字；源 caller 修改不改变异步已捕获 intent；格式对象与 UI owner 同步 | text/runs 不匹配、NaN/非法格式、隐藏/只读/record、DV/spill/预算/extent、viewer 拒绝整个事务 | Pending |
| O2.1-c | worksheet/range 保护、locked/unlocked/allow flags、移除及 undo/redo；真实 Java 保存重开/native 保护往返 | 非 owner、异 sheet/out-of-extent 范围、坏 allow/unknown field、不存在的 remove、假 workbook-wide scope 均拒绝 | Pending |
| O2.1-d | 句柄稳定、snapshot immutable；身份切换/关闭使新名称/保护句柄不可调用；其他 Workbook 可继续使用 | 退休句柄不能改新上下文对象，不把新配置授权给任意资源 | Pending |
| O2.1-e | 完整真实 browser 继续19原流程，扩展 sheet.remove undo→redo→undo，新增名称/富文本/保护 save/reopen/Excel文件断言，检查 console/network | 客户端与直接 REST 的 invalid/unauthorized 均拒绝且版本/operation/snapshot 不变；实机 Excel仍 Blocked | Pending |

### O2.1 首轮冻结与统一纠正契约

首轮 SDK 62/60 Pass/2 Fail；build 仅 native UAT optional 名称字段的类型断言失败；calculation 457 Pass，boundaries Pass，Java 369/367 Pass/0 Fail/2 Skipped + package。完整 unit 1581/1527 Pass/54 Fail。原始源与日志保留 `/tmp/sdk-o21-first-pass.patch`、`/tmp/sdk-o21-first-{sdk,build,boundaries,unit,formula,java}.log`，没有在验证中反复改产品。

只读链路复核确认 PermissionService.getCapabilities 以 online=false 优先选择 LOCAL_CAPABILITIES，覆盖已确认 serverRole。全表保护的真实 affectedRanges=[] 不经过单格范围 ACL，因此 editor 错误获准。统一纠正：角色投影一旦存在，能力始终由它决定；连接状态仅决定无投影时 online 的 fail-close，不能将已知 editor/viewer 升级为本地 owner。Session 远程离线 mutation guard 和 Java requiredRole 不变。增加 owner/editor/viewer 的在线/离线命令及 mutation 重放成功/拒绝测试，并保留 SDK editor 拒绝及全模型不变断言。

两处测试契约同期更正：超 10000 格预算必须精确断言既有 UNSUPPORTED_FEATURE，不改产品预算；原生导入的可选名称字段先断言存在再收窄，不引入空数组 fallback。完成这一个纠正批后统一验证，再冻结产品 head 做真实 21 条 browser UAT。主题、模板、原生格式全量和 T1 等仍 Pending，不扩大 O2.1 的结论。

### O2.1 原生边界整链补全契约（落代码前）

权限纠正验证：permission 11 Pass；SDK editor 拒绝成功。剩余 rich-text 测试在 DV 前置步骤使用 CF 的 value1 字段，必须改为正式 DataValidationRule.formula1='10'，不改 DV 规则或放宽拒绝。完整 unit 1582/1529 Pass/53 Fail，新增仅此新 fixture 前置问题。

原生只读审查确认 comment 未写入原生 definedName attribute，anchor 未保存；保护原生 ID 在 import 边界重建，native allow flags 固定展开。这不是服务器 save/reopen 丢失，而是格式边界契约。新增 metadata v5 仅保存 anchored name 的身份、原公式指纹与 anchor，原生 formula/comment/hidden/scope 仍为权威，不保存第二份名称模型。3/4→5 仅原生 import migration 边界；Excel 修改原生名称或删除后不采用旧 anchor。坏/重复/越界/外 sheet anchor metadata 拒绝。增加真实文件 round-trip→同一引擎相对计算成功、旧 v4 import 与坏/stale metadata 拒绝/不覆盖测试。

原生 sheetProtection 的 allow booleans 按格式展开，ID 为 import 边界生成；UAT 比较全部保护行为字段，服务器 save/reopen 仍严格比较原模型。范围/工作簿保护及多个活动 sheet 规则当前无原生完整表示，export 明确 NATIVE_DOCUMENT_UNSUPPORTED，不能挑第一条或丢弃。inactive sheet rule 不产生原生保护；活动保护的 flags 与 passwordHash 仍以原生节点保存。相对名称 anchor extension 的 SDK 文件往返不等于已验证桌面 Excel 的执行能力，实机 Excel仍 Blocked。全部原生格式 X1/S1 仍 Pending。

原生补全统一验证首轮新语义成功：SDK 62 Pass；native 65/66 Pass，唯一旧 v3 fixture 仍从 version4 改写，已按正式 version5 生成旧包；build 两处仅测试给 readonly collection 重新赋值，改为测试 setup 对现有 collection 的 splice，不改产品语义。首轮日志 `/tmp/sdk-o21-native-{sdk,native,build,boundaries}.log` 保留。产品源仍冻结，统一迁移这三处测试契约后检查。

O2.1 冻结产品统一检查：SDK 62/62 Pass；permission 11/11 Pass；native 66/66 Pass；calculation 457 Pass（本批公式引擎未变）；boundaries Pass；Java 369/367 Pass/0 Fail/2 Skipped + package。完整 unit 1584/1532 Pass/52 Fail，与 O1 产品失败标题新增0/消失0；保留完整门禁 Fail。日志 `/tmp/sdk-o21-native-sdk.log`、`/tmp/sdk-o21-corrected-permission.log`、`/tmp/sdk-o21-final-{native,unit,build}.log`、`/tmp/sdk-o21-native-boundaries.log`、`/tmp/sdk-o21-first-java.log`。build 与真实 browser 完成后记录实际结论，不先填 Pass。GitHub 0f318900 的 canonical-build FAILURE 明确发生在完整 unit 检查（run37120591294 step11），不以局部绿色覆盖。

### O2.1 名称撤销整链纠正（代码前冻结）

产品 ca5ba64b 完整真实 browser 20/21 Pass。失败 `/tmp/sdk-o21-browser.log`、`/tmp/sdk-product-uat-moD4I1/evidence`、`/tmp/sdk-o21-first-browser-test-results`。真实 REST 409 UNDO_RESULT_MISMATCH，前9条规范操作通过。只读 H2 +实际 Java registry 重放 `/tmp/SDKO21Replay.java`、`/tmp/sdk-o21-replay.log`、`/tmp/sdk-o21-{expected-preimage,candidate,rejected-operation}.json` 证明唯一差异是 definedNameModels 顺序：删除中间 local Rate 后，name.set 只能尾部追加。

定义名称列表当前为规范 authored 顺序，不排序快照、不降低 Java preimage 守卫。新增正式 name.restore {model,position}，仅由已拥有的 undo 触发；Core 模型以单一 replaceDefinedNames 事务恢复缺失身份到有效原位置，所有正常 author name.set 仍保留既有位置。客户端 remove 的 inverse 捕获原位置；Java 同字段、范围、角色与候选验证。恢复不能覆盖已有身份，位置整数且0..当前长度，坏字段拒绝；直接非 undo REST 明确 RESTORE_REQUIRES_OWNED_UNDO。这是新增规范操作，无旧字段别名或第二个名称模型，存储 snapshot 版本不变。

O2.1-f 预验收：三项名称中移除中项，undo 完整快照与名字顺序一致，redo→undo 再次恢复；依赖重算、其他名称 anchor/comment 不变；恢复重复/越界/坏scope拒绝且模型不变；无 owned undo 的 direct REST restore 无任何revision/history写入。严格前置 guard 和现有21项UAT均保留。之后统一 checks、提交同一草稿PR、干净head完整真实验收。

名称恢复整批统一检查：SDK63/63 Pass；build/typecheck、boundaries Pass（45声明）；Java370/368 Pass/0 Fail/2 Skipped + package；完整unit1585/1533 Pass/52 Fail，无新失败。日志 `/tmp/sdk-o21-restore-{sdk,build,boundaries,java,unit}.log`。守卫与前置快照比较未修改，下一轮真实browser仍完整21项、retries0。

### O2.1 原生属性契约同步（代码前复核）

产品4b8a4e7b完整真实browser再次20/21 Pass；名称中项 undo 已实际通过 Java 严格preimage，后续富文本/保护的undo→redo、保存重开都执行到native export。失败为 NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED（非 HTTP 提交失败），证据 `/tmp/sdk-o21-restore-browser.log`、`/tmp/sdk-product-uat-AE97W0/evidence`、`/tmp/sdk-o21-restore-browser-test-results`。只读契约审查确认 definedName.comment 读写已实现，唯一源策略 contracts/native-document-structural-capability.json 的 child attributes 仍仅 name/localSheetId/hidden；源码生成的comment被当unknown-workbook-node，正常Save As拒绝。补全正式comment属性与全部generated消费者，不能绕过guard/改判定unknown。

新增 O2.1-g：带 comment/anchor 名称的native import→编辑workbook属性→实际artifact Save As→再import，comment/anchor全部保留；未知future属性仍触发严格UNCHANGED_SAVE_REQUIRED且未修改sourcebytes。完成这一个契约同步批后检查与干净head真实UAT。49父项/完整unit/SSO/桌面Excel状态仍不扩大。

属性契约同步统一检查：native66/66 Pass；build/typecheck、boundaries Pass；Java370/368 Pass/0 Fail/2 Skipped + package。日志 `/tmp/sdk-o21-comment-{native,build,boundaries,java}.log`。未知属性拒绝与sourcebytes不变测试通过，服务端guard未放宽。完整unit已记录52 Fail且未修改公式/应用运行时；干净提交后完整21项真实browser仍待执行。

O2.1 干净产品head **68aec758** 完整真实 Java21/H2/Chromium **21/21 Pass**，retries0、无HTTP mocks。成功场景执行全部最后断言：名称中项remove→undo顺序/值恢复、相对anchor/comment/hidden保存重开与实际xlsx再import；富文本literal等号、undo/redo、范围矩阵；保护锁定格拒绝/解锁格写入、保护remove→undo/redo、关闭旧句柄退休；viewer和非法REST保护/name.restore整体零revision写入。console/network diagnostics全通过。证据 `/tmp/sdk-o21-comment-browser.log`、`/tmp/sdk-product-uat-tcJlxV/evidence/sdk-o21.xlsx`。O2.1-a..g声明范围通过，Theme/完整格式/多种保护原生输出仍待各自实施；原49父项、全量unit52 Fail、ERP/SSO/桌面Excel/外部SQL Blocked仍保留。该head两项canonical-build当前in_progress（run37122552603/37122547520），未声明CI通过。

### F1.1 财务标量函数实施前契约

首批12个：PV/FV/PMT/NPER/IPMT/PPMT/ISPMT/EFFECT/NOMINAL/SLN/DOLLARDE/DOLLARFR。只扩展唯一executable registry、参数契约、既有标量数组broadcast与共享语料，不新增SDK/Java计算服务、函数名别名或空实现。金额流入正/流出负；period/type和Excel文档一致；EFFECT/NOMINAL npery、DOLLAR fraction截整数，负fraction不能截成负零后接受。TVM零利率使用解析极限，log1p/expm1避免微小利率相消；PV/PMT在大正利率/长周期采用折现表达式；不可求实数/非有限结果明确Excel错误。NPER不做隐式迭代；输入fault继承现有根求值边界，不能IFERROR吞掉。

Microsoft正式参数说明已复核：PMT、FV、NPER、IPMT、ISPMT、EFFECT、NOMINAL、DOLLARDE函数页面（support.microsoft.com/en-us/office）；本地只读记录 `/tmp/sdk-financial-*.txt`。黄金值为独立小整数现金流与手算分数，非实现输出生成。NPV的scalar/reference coercion、coupon/day-count/求根/折旧/累计尚需各自完整契约，未在这批发布；完整455目录仍Pending。

| 细项 | 成功验收 | 拒绝/边界验收 | 初始状态 |
| --- | --- | --- | --- |
| F1.1-a | 每个函数独立黄金值、参数契约、registry发现、inline和实际Worker任务入口结果相等 | 每个函数不足/过量参数、字符串坏数、普通错误，报对应Excel错误 | Pending |
| F1.1-b | 年金期初/期末、零/微小/负利率、长周期折现、IPMT+PPMT=PMT、NPER解析周期 | 坏type/per、无法求实数、零分母/溢出，不能返回NaN/Infinity或成功默认值 | Pending |
| F1.1-c | EFFECT/NOMINAL季度互逆、fraction截整与负金额、SLN现金流、二维broadcast | 非正利率/无有效分期、负/零fraction、错array形状与SLN零life明确拒绝 | Pending |
| F1.1-d | 隐藏数值/跨簿输入仍消费，来源更新重算，授权恢复inline/Worker同值 | 来源denied/unavailable/broken故障经全部12函数及IFERROR仍不可吞掉 | Pending |
| F1.1-e | public SDK矩阵写12公式→真实Java提交→保存关闭重开→实际xlsx导出/再import→SDK重算黄金值 | 原始无关cell不变，运行时对象无越层API，console/network干净 | Pending |

此批不把Java持久化快照中的formulaValue当新计算真相；导出缓存值投影另需唯一计算owner契约审查。原生公式保存/再import后SDK计算与桌面Excel实机缓存验收分别记录，不伪造缓存或宣称实机Pass。源码一次完成上述整批后统一checks，提交同一草稿PR，再干净head逐项真实验收。

F1.1 首轮统一检查：calculation主套544/543 Pass/1 Fail，新增失败为数组普通错误fixture直接setValue(FormulaError)，违背ScalarValue持久化与Worker快照契约；改为真实错误公式=#N/A，不接受非法snapshot、不放宽守卫。build仅验收代码错误：readonly readValues传可变类型、非公开trash方法和同一ScalarValue类型错误。产品financial源冻结不变。完整unit1677/1624 Pass/53 Fail，新增唯一为上述fixture；旧52失败保留。boundary Pass。日志 `/tmp/sdk-f11-first-{calculation,build,boundaries,unit}.log` 与首次源码 `/tmp/sdk-f11-first-pass.patch`、`/tmp/sdk-f11-first-financial*.ts`保留；统一修正验收调用契约后重验，不能算首轮通过。

F1.1 统一修正后：calculation **544+5=549/549 Pass**，全部12函数registry/语料/arity、零/微小/负rate与大正rate长期折现、期初/期末本金利息、标量二维broadcast/错形状、普通错误处理、隐藏读值与12函数三类跨簿fault/Worker恢复均通过。build/typecheck与boundary Pass。完整unit **1677/1625 Pass/52 Fail**，与O2失败清单新增0/消失0；不能算完整门禁通过。产品financial源与首轮一致，仅验收输入统一修正；日志 `/tmp/sdk-f11-final-{calculation,build,unit}.log`。实际registry148，其中12新增官方函数；官方455对齐仍缺308，F1整项Pending。真实SDK保存/重开/原生文件流程待干净提交后完整22项browser，无真实Excel环境仍Blocked。Java产品未变，沿用本轮已编译370/368 Pass/2 Skipped的Jar，不伪造服务端函数计算。

F1.1 干净产品head **e568053a** 完整真实Java21/H2/Chromium **22/22 Pass**（retries0，无HTTP mocks）。12个公共SDK公式矩阵真实Java提交→保存→close/open→原生xlsx导出→解析逐条公式原文→真实服务器import→公共SDK重算12黄金值，C1无关值77不变；独立Rates簿 .5→0 使PMT与IFERROR(PMT) -225→-125；真实trash使两式带sourceId的#REF!，restore恢复-125。全部console/network检查通过。F1.1-a..e声明范围Pass；实际文件 `/tmp/sdk-product-uat-XsjbOW/evidence/sdk-financial.xlsx`，日志 `/tmp/sdk-f11-browser.log`。完整F1/其他计划仍Pending，原49父项与52个全量unit Fail未扩大。68aec758两个canonical-build实际FAIL，已确认step11 complete frontend unit，不称CI通过。

### A1.1 规范验收输入同步（代码前整批审查）

审查真实入口后限定本批为已确认的验收契约，产品源保持e568053a不变：CellMatrix.get只规范化该单元格，forEach才显式整簿materialization，count/point读不递增内容revision；非法font的point、整读取与持久化均拒绝且原输入/内容版本不变。规则引用索引的structural query表示插入位置及之后的suffix，A7仍属于at5，精确A6/A7应使用正式range query，同时增加suffix断言而非减少约束。Remote replay fixture必须注册真实inverse handler，并明确注入真正的preflight failure，不能在注册不完整时宣称observer路径验收。buildOperation生成正式UUID，会话owner另外提供稳定会话identity；不能期待无来源的fixture-session常量。坏schema fixture携全必需字段，另验missing-session拒绝。Pivot已拥有worksheet identity的manifest必须给header+rowCount完整sourceRange；孤立sourceSheetId不是合法来源。

| 细项 | 成功验收 | 拒绝验收 | 初始状态 |
| --- | --- | --- | --- |
| A1.1-a | 单格/稀疏读取不整体水合、不改revision；显式forEach结果相同且materialization成功 | 坏font单格/whole/persistence拒绝；零callback/无部分水合、原输入与内容版本不变 | Pending |
| A1.1-b | CF引用A6→A7精确索引移动，结构suffix仍正确；清索引重建正确 | 重放participant故障保持已提交全内容，之后fail-stop；handler前置拒绝无内容改动 | Pending |
| A1.1-c | 完整OperationEnvelope所有公开字段与UUID合规；真实Pivot block值/row paths一致 | missing-session、wrong-schema、unknown-mutation分别精确拒绝；坏source pair/尺寸/owner在读block前拒绝 | Pending |

完成统一验收代码pass后一次检查相关套件、完整unit与typecheck/boundaries；不要把修正某些fixture概括为其他52问题均为fixture，也不移除或降低原成功/拒绝断言。所有其他门禁与功能继续逐条处理。

A1.1 首轮typecheck/boundaries Pass；完整unit1678/1636 Pass/42 Fail，旧52中10项消失、新标题0。最后同一Pivot成功fixture进入了真实读路径，暴露第二处契约：rowOrder [2,0,3,1] 下 overlay.rowIndex 是 immutable physical index，而原fixture用0却期待改首个展示行；源row paths同样必须反映真实physical+header偏移。完整只读审查确认block-source API文档、query.getPhysicalRow/getLoadedPhysicalRow、cell overlay owner与rowPathAt一致，产品源未改。整批验收输入修正为physical2覆盖首展示行、sourcepaths明确[3,1,4,2]；增加physical0只改第二展示行、之前source view不被改写与rowOrder不变断言，不能改源计算或用逻辑行假装来源地址。首轮 `/tmp/sdk-a11-{typecheck,boundaries,unit}.log`、`/tmp/sdk-a11-first-pass.patch`保留。

A1.1 最终统一typecheck/boundaries Pass；完整unit **1678/1637 Pass/41 Fail**，相对52消失11、新增0。原54历史清单保留，当前41已逐条更新至sdk-product-unit-failures.md。成功/拒绝与物理row path/view独立性全通过；产品源保持e568053a（其真实22/22证据有效），本批只调整验收契约。日志 `/tmp/sdk-a11-final-{typecheck,unit}.log`、`/tmp/sdk-a11-boundaries.log`，比对清单 `/tmp/sdk-a11-current-failures.json`。A1.1-a..c范围通过，不代表完整A1、CI或其他41错误已解决。

### O1.2 实施前细项（Pending）

- a：同表 cut/replace，内部、外部、绝对、跨表和区域外未移动引用；provenance/barcode/cache 正确。
- b：真实 Java ACK、一次 undo/redo 恢复全部源/目标和依赖，不增加 SDK 逆操作栈。
- c：错尺寸、跨表/簿、重叠、预算、offline、退休、权限/隐藏、公式组/不可逆引用明确拒绝，状态不变。
- d：真实浏览器 SDK、保存重开、实际 xlsx 导出/原生解析/服务器重导入、console/network；桌面 Excel Blocked。

O1.2 首轮00853d8e：完整22/23，新增a/b/c/d仍Fail（首次undo真实409；后续native未执行）。源代码及trace冻结，Java恢复事实owner修正前不升Pass。O1.2-e实施前Pending：目标覆盖单元格按真实preimage精确恢复，重复/篡改/无关/跨sheet恢复拒绝，完整快照guard不变。
