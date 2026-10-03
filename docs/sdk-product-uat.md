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
| OO-01 | SDK → Workbooks → Workbook → Worksheets → Cell/Range；真实显式寻址、typed 值/公式、保存与关闭 | 同一 canonical 数据/权限/计算/写链；不暴露可变 model/session/API；非法地址、失效、越权 typed 拒绝 | Pending |
| OO-02 | 按所有 Excel/Aspose 类对象逐能力调用、编辑、保存、重算、往返 | 完整领域对象与 read/edit/write/preserve capability，不能只有 Web 菜单可用 | Pending |
| OO-03 | 浏览器无编辑器、Node/Java 文件/Worker/输出宿主 | 宿主仅 I/O/线程，SDK 拥有语义与文档；缺失宿主明确拒绝 | Pending |
| MWB-01 | 一个 SDK 同时打开多个 workbook；重复打开；关闭一个、subject 切换、全部释放 | 按 unitId 唯一 owner、对象与租约；不影响其他 workbook，无旧身份读写 | Pending |
| MWB-02 | 真实跨 workbook SUM/范围/其他函数，source 更新→提交→refresh→target 重算与保存 | source ID/sheet ID/subject/revision/access 定义明确；撤权清缓存、#BLOCKED!、恢复重算 | Pending |
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
- `npm run test:sdk`：30/30 Pass；auth/identity/runtime/dimensions/data 成功与拒绝路径。
- protocol focused：33/33 Pass。
- `npm run check:boundaries`：Pass（包含 generated contracts/registry/stack/provenance/既有 acceptance matrix）；`npm run test:calculation-domain`：446 + 5 tests Pass。不能据此宣称 Web 已只依赖 SDK。
- `npm run test:native-codecs`：17/17 Pass；`npm run test:pointer-gesture`：7/7 Pass。
- OOXML suite：**63/63 Pass**；首次图片 DrawingML、合法/非法 preserved XML、业务/native sheet identity、自定义原生 3D 饼图保留与编辑拒绝。Camera typed error 与 Pivot OPC ContentTypes fixtures 已按真实契约修正。codec 通过不等于完整原生互通验收。
- 资产引用提交屏障：2/2 Pass，真实 CollaborationSession 的 ACK 前不 GET，ACK 后读取；离线/拒绝不 GET 且队列状态保留。Pivot detail lifecycle：21/21 Pass，完整恢复和 region/hyperlink 拒绝无部分状态。
- Java 21 `mvn ... package`：实际 Maven 汇总 **321 tests，0 failure/error/skipped**，真实 H2 integration；不使用含过期报告的 XML 目录总数。
- 完整 `npm run test:unit`（本次 Pivot lifecycle 实现，提交前干净语义检查）：**1545 tests，1491 Pass，54 Fail**；上次 ed0f570 为 1543/1487/56。基线为 1525 tests，1442 Pass，83 Fail。不能未经逐项对照认定失败都来自基线；完整失败清单另见 sdk-product-unit-failures.md。
- GATE-01：Fail。CI 真实 SDK UAT 与完整 test:unit 独立执行，完整日志作为 artifact。ed0f570 的 Windows push run 37093921510：12/12 UAT Pass，完整单测 56 Fail；同 head 的 PR run 37093924549 在 DATA 新工作簿打开的默认 5 秒内未 ready（3 Pass / 1 Fail / 8 未执行），缺少失败 trace，原因未确认。34c44dbe 已补首次失败 trace/截图/console/pageerror/网络拒绝记录与递归 artifact，没有增大超时或增加 action retries，后续 Windows 结果待记录。
- BOUND-01：Fail。Web 仍有内部包业务 import；WorkbookSession 仍 8121 行并暴露给 Web，完整 WorkbookHandle/领域拆分未完成。认证、组合根、目录 service、尺寸规划/worker、六个数据 action 已实质迁移，不能据此宣称所有领域完成。
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
