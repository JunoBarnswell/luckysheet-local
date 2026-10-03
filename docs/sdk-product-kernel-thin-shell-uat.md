# SDK 产品内核与 Web 薄壳：落代码前 UAT

本表在产品代码修改前建立。全部 Pending 表示尚未验收，禁止因为对象/方法存在就标记 Pass。执行入口必须是 `createSpreadsheetSdk()`；真实浏览器请求 Java/H2、检查 console/network，并检查保存/重开与真实导出文件。不得直接访问内部 Session 作为 SDK 验收。

基线 main：`3ff47ca6df8de4af49b809649831ea376179d24c`。证据必须记录执行时实际 head/build ID、日志、浏览器 trace、网络响应、native 文件，以及每条状态/失败原因。

## 按领域成功与拒绝细项

每个可写领域的成功项：执行真实动作 → 验证 canonical read 与 view → flush/Java ACK → undo/redo（适用）→ 第二客户端观测（适用）→ save/close/open → applicable native roundtrip。拒绝项捕获 SdkError code/operation/object/recovery，比较完整前像，验证无半写入/假成功。

| ID | Domain | 具体成功细项 | 具体拒绝细项 | 状态 | 本轮证据 |
|---|---|---|---|---|---|
| D01 | Auth | 登录、刷新、注销、guest 路由捕获及凭证切换 | token/CSRF 不出公开 snapshot；跨上下文异步响应与旧句柄退休 | Pending | — |
| D02 | Identity | 独立读取当前 verified identity 并订阅变化 | 未验证 JWT/profile 不授予权限；身份与管理用户职责分离 | Pass | U01：SDK identity/users contract；U02：真实公共入口同 verified context、只读 identity API。未验证和匿名不发布身份。 |
| D03 | Users | 真实管理员列表/创建/禁用/密码重置 | 非管理员 FORBIDDEN；错误只有 SdkError | Pass | U01：非法输入、403/503、旧上下文响应拒绝；U02：真实列表/创建/禁用/启用/重置密码及会话失效、guest FORBIDDEN。 |
| D04 | Workbook Center | 目录 create/list/open/import/export/share/trash/restore/purge | 非法 ID/名称、无权限与 retired catalog；公开面没有 resolve/retire | Pending | — |
| D05 | Workbook | 公开入口并发 open 两本、关闭/重新打开、dispose | 旧对象/任务退休；Web 永远不拿 Session/Resolution | Pending | — |
| D06 | Worksheet | 新增/重命名/移动/删除、axes/extent/freeze | 最后可见 sheet、非法 extent 与 protected sheet | Pending | — |
| D07 | Cell | 真实程序写入/计算读、hidden cell、rich text | viewer/hidden access/fault；公式隐藏与受限字段 | Pending | — |
| D08 | Range | 矩阵、格式/fill/merge、同 sheet 和跨 sheet move | 形状/预算/保护/冲突及重叠前像；全模型 unchanged | Pending | — |
| D09 | Selection | 地址、多区域、行列、全选、move/extend 与鼠标键盘 | hidden/frozen/merged/access bounds；选取/编辑/提交同一地址 | Pending | — |
| D10 | Editing | begin/draft/IME/caret/commit/cancel、公式引用与 autocomplete | 验证不通过保留 draft；viewer/权限切换拒绝；没有 UI 自修复 | Pending | — |
| D11 | Clipboard | copy/cut/paste 与 values/formulas/formats/transpose/links | 系统 clipboard 拒绝状态可见；跨 sheet cut 正式单事务 | Pending | — |
| D12 | Formula | 模式/重算、audit/evaluate/catalog、授权跨 book函数 | 源权限故障不可吞；manual pending 拒绝 stale output | Pending | — |
| D13 | Defined Name | workbook/sheet scope 与 anchor/comment/hidden，引用 move | 重复/非法名字与无权限；原生重开仍保留唯一 owner | Pending | — |
| D14 | Table | create/resize/rename/style/filter/total/header/columns/convert | 重叠/未知原生 table 拒绝；structural schema/refs 保持一致 | Pending | — |
| D15 | Data | 查询加载/刷新、sort/filter/text-to-columns/dedup/subtotal | 越权/过时代际/部分物化/未知列不回落；全部小计函数对齐 | Pending | — |
| D16 | Validation | 值列表/范围/公式/typed bounds 与公式输入，原生重开 | 客户端与 Java 相同算值拒绝非法输入；重复规则与 wrong preimage | Pending | — |
| D17 | Conditional Formatting | 增改删/优先级/范围/公式，原生重开 | 重复规则/未知字段/权限拒绝且 unchanged | Pending | — |
| D18 | Drawing | shape/connector/image/textbox/control/barcode/camera/group/z-order/align | asset 缺失/未知 native payload/对象保护；无假图层 | Pending | — |
| D19 | Chart | 创建/series/type/legend/labels/style/bounds、native重开 | 缺数据/无权限/未知 native chart 不可假成功 | Pending | — |
| D20 | Pivot | 创建/fields/layout/filter/aggregate/calculated/drilldown/refresh/controls | task version/权限 fence；结果不与另一模型或旧加载状态混用 | Pending | — |
| D21 | Sparkline | source/group/style/create/update/delete | 非法源/anchor/duplicate 与保护；命令逆向和重开 | Pending | — |
| D22 | Review | comment/thread/reply/resolve/note/hyperlink | commenter 与 viewer 限制；非法链接与 exact ownership | Pending | — |
| D23 | Range Access | 创建/更新/删除访问区域与 effective view | manager/tenant/context 与 read遮蔽；服务端是 authority | Pending | — |
| D24 | Permission | role+range+sheet/object protection 的 effective能力 | 旧 role/非法 owner 升级拒绝；offline 不提高权限 | Pending | — |
| D25 | Collaboration | 两真实客户端 presence/cursor/ACK/resync/pending/revision | 认证退休/socket关闭、远端 structural 与本地 pending 冲突 | Pending | — |
| D26 | History | undo/redo/revision/preview/restore/replay 状态 | 篡改/重复恢复/wrong preimage 拒绝；remote rebase 核心操作不大面积 invalidate | Pending | — |
| D27 | Document | save/saveAs/export、native capability/source artifact/compatibility | native 未知部分与宏保真；缺资产/未提交/过期身份拒绝 | Pending | — |
| D28 | Print | setup/area/titles/break/scale/preview/PDF | 无权限/非法范围；保存重开与导出真实页 | Pending | — |
| D29 | Viewport | scroll/zoom/freeze/resize/geometry 与覆盖层 | PaneMap 唯一坐标；隐藏轴与 freeze 四区无重叠 | Pending | — |
| D30 | Render Interaction | 真实 canvas pointer/keyboard/drawing/fill/drag/resize | hit/selection/edit/commit 地址相同；错误上游不被画布掩盖 | Pending | — |

## P0 结构完整性细项

| ID | 设置与操作 | 必须观测的断言 | 状态 | 本轮证据 |
|---|---|---|---|---|
| S01 | 源/目标 worksheet 各有值、公式、外部引用和目的地旧值；公开 Range.moveTo 跨 sheet | source 清除、目的地完整；absolute/relative/unmoved/其他 sheet引用正确；一次事务/ACK/历史项；undo恢复目的地旧值、redo；保存与原生重开 | Pending | — |
| S02 | cut/paste 同 workbook 不同 sheet，并包含 names/rules/chart/pivot/data/drawing anchors/table | 所有 participant 单事务和同一 structural patch；禁止 clear+write，失败完整前像 unchanged | Pending | — |
| S03 | cross-sheet range/partial-qualified endpoints、完整/局部/绝对引用，insert/delete/shift/move | TS/Java 同一 ownership/reference vectors；语义/资格保留；不在核心场景抛占位 UNSUPPORTED | Pending | — |
| S04 | worksheet column 在 table 内部插入/删除、移动/排序 table 区域 | schema/column ID、structured refs、filter/total/headers、data region 全部正确且逆向恢复 | Pending | — |
| S05 | A 本地已写并有 undo/redo/pending，B 远端 range move/shift/permutation/sheet/table | 本地历史、pending、selection、range owner 全部正式 transform；有效核心历史不整体 invalidate；真正冲突 typed错误 | Pending | — |
| S06 | DV 公式输入及公式界限/跨 sheet依赖/计算错误/权限源 | 客户端与服务器 evaluator同语义；Java独立算值并严格拒绝失败，不信客户端 proof、不漏数据 | Pending | — |
| S07 | collapsed outline+manual/filter hidden 共同存在，执行 row sort | 三种 visibility owner 独立、permutation同步，值读不因隐藏消失，undo与native重开正确 | Pending | — |
| S08 | 小计全部 Excel函数及空值/错误/嵌套subtotal/隐藏行 | sole formula engine与canonical Data API一致，undo/重开/原生公式正确 | Pending | — |

## 静态与公共契约细项

| ID | 操作 | 断言 | 状态 | 本轮证据 |
|---|---|---|---|---|
| B01 | 扫描 Web 全部非测试业务源码与 manifest | 八种内部 spreadsheet imports 全部0；仅SDK/UI system/React/ReactDOM 业务依赖 | Pending | — |
| B02 | 搜索 Web Session/Resolution/Command/Mutation/客户端/engine/patch 与业务fetch | 全部0；不存在相对路径越层或包装旧内部类型的绕过 | Pending | — |
| B03 | 检查 SDK核心公共入口及依赖闭包，在非React consumer导入并使用 | 无React依赖/自动加载；React只从 sdk/react导出 | Pass | U01：core-entry.test.ts 主动拒绝 React/ReactDOM resolve 后真实 import/初始化/拒绝/dispose；React 表面仅 sdk/react。 |
| B04 | 检查全部公共DTO、actions、失败类型；非法输入与真实HTTP拒绝 | 显式SDKDTO与actions，无 Omit/internal继承/resolve/markOpened/retire；所有抛出的错误均SdkError | Pending | — |
| B05 | Web交互读取/写入及原Session搜索 | WorkbookView单一只读发布，SDK对象和领域直接行为；WorkbookSession已职责分解并删除，非改名/转发 | Pending | — |
| B06 | 对boundary测试注入内部import、CommandDescriptor、relative escape、业务fetch | 自动拒绝各越层样例；合法SDK/host primitive通过，无广泛allowlist | Pending | — |
| B07 | 身份/guest与管理领域检查 | capability在SDK初始化捕获并清URL，公开state无token；identity/users分离且同上下文fence | Pass | U01：guest-capability/context retirement；U02：真实 capability 打开、URL 清除、第二 SDK、reload；公开身份和认证状态无 token。 |

## 第一组实现的额外逐项细项（设计后执行）

| ID | 操作与具体断言 | 状态 | 本轮证据 |
|---|---|---|---|
| I01 | 实际 Node SDK core 导入，主动拒绝任何 React/ReactDOM resolve；初始化、anonymous identity、users FORBIDDEN、dispose | Pass | U01：core-entry.test.ts；独立 Node 非 React 消费者实际执行。 |
| I02 | identity verified上下文与Auth同一来源、更新订阅、同身份refresh稳定snapshot、匿名/退休；管理方法只在users | Pass | U01：users/domain.test.ts；U02：同 context 与精确公开 identity/users keys。 |
| I03 | users真实管理成功及非管理员/非法参数/403/503/旧身份已返回response拒绝；CSRF/context仍由Auth唯一owner负责 | Pass | U01：auth/domain.test.ts + users/domain.test.ts；U02：真实管理员四种操作和 guest 拒绝。 |
| I04 | SDK初始化同步清fragment/query share，私有tab凭证重载、路由退出退休上下文、signout清除；snapshot不泄漏token | Pass | U01：guest-capability.test.ts；U02：真实 fragment 清除、私有 tab 重用和 reload。 |
| I05 | 真实Java/H2上独立公共SDK identity/admin调用；真实guest share打开、viewer读/写拒绝/管理员拒绝、URL清除、浏览器reload | Pass | U02：sdk-workbook-acceptance.spec.ts 的 I01-I05，Java/H2 真实请求，无 HTTP/计算 mock。 |

## 完整性矩阵（代码与行为双重证据）

| Domain | Public API | Canonical owner | Command chain | Permission | History | Collaboration | Persistence | Server authority | Web migrated |
|---|---|---|---|---|---|---|---|---|---|
| Auth | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Identity | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Users | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Workbook Center | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Workbook | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Worksheet | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Cell | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Range | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Selection | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Editing | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Clipboard | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Formula | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Defined Name | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Table | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Data | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Validation | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Conditional Formatting | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Drawing | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Chart | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Pivot | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Sparkline | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Review | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Range Access | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Permission | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Collaboration | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| History | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Document | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Print | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Viewport | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |
| Render Interaction | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending | Pending |

## 必跑门禁

| Gate | 命令/观察 | 状态 | 本轮证据 |
|---|---|---|---|
| typecheck/build | npm run build（包括 tsc） | Pass | U03；首组实现通过，后续产品修改须重新执行 |
| boundary/stack | npm run check:boundaries（含拒绝测试与真实全源码扫描） | Pending | — |
| contracts/mutation | generated TS/Java contracts 与 registry，无漂移 | Pass | U04；首组当前契约/45 mutation registry 无漂移 |
| SDK unit/contract/integration | npm run test:sdk，覆盖各正式领域成功+拒绝 | Pending | — |
| calculation/Worker | npm run test:calculation-domain，包含跨bookfault与TS/Java vectors | Pending | — |
| SDK real browser UAT | npm run test:sdk-uat，全createSpreadsheetSdk入口、console/network/native files | Pending | — |
| frontend unit | npm run test:unit，全量结果；不隐藏当前main已有40项失败 | Fail | U05：1681 tests，1641 Pass、40 Fail；仍需完成整改 |
| backend | Java21 Maven package/test，H2真实服务器和服务端authority | Pass | U06：375 tests，373 Pass、0 Fail、2 Skip；新增 WS 与 guest identity 合同 |
| OIDC/ERP SSO deployment | 用户已确认没有真实身份环境 | Blocked | 用户现有会话答复 |
| Desktop Excel interoperability | 用户已确认没有桌面Excel验收环境 | Blocked | 用户现有会话答复 |
| External PostgreSQL/MySQL | 当前无真实外部数据库验收环境 | Blocked | 环境事实；不得禁用TLS或重建真实库 |

## 首组执行记录（2026-10-03）

本组产品源码提交为 `9910ed01909bbda1e451576ad0f6491bd7adc4a8`（`fix(auth): separate workbook capabilities from verified identities`），浏览器执行时 worktree clean，source/build/backend ID 全部为此 SHA。基线远程 `main` 仍为 `3ff47ca6`。

| 证据 | 实际执行与结果 | 日志/产物 |
|---|---|---|
| U01 | `npm run test:sdk`，71/71 Pass | `/tmp/sdk-thin-shell-identity-sdk.log` |
| U02 | `npm run test:sdk-uat`，真实 Java21/H2 + Chromium，24/24 Pass；逐用例 console/page error/network 断言；真实保存、重开、跨工作簿授权计算和 native 导出/重导入 | `/tmp/sdk-thin-shell-identity-browser-verified.log`；`/tmp/sdk-product-uat-SJC2EQ/evidence`；provenance 见下文 |
| U03 | TypeScript + production build Pass | `/tmp/sdk-thin-shell-identity-build.log` |
| U04 | 当前 boundary/stack/contracts/mutation/e2e-artifacts/acceptance-matrix Pass；这仍是旧边界门禁，不能替代 B01/B02/B06 要求 | `/tmp/sdk-thin-shell-identity-boundaries.log` |
| U05 | 全量 frontend unit 40 Fail，失败名称与 main 基线一致；没有跳过或隐藏 | `/tmp/sdk-thin-shell-identity-unit.log` |
| U06 | Java21 Maven package：375 tests，0 Fail、0 Error、2 Skip | `/tmp/sdk-thin-shell-guest-contract-java.log` |

浏览器 provenance：runId `9910ed01909bbda1e451576ad0f6491bd7adc4a8-20261003141920607`；Node `v24.19.0`；Chromium `151.0.7922.34`；viewport `1440x960`；package-lock SHA256 `9199bcde46b91ac0b9f165f621bad8566476680db2b341162cc31fea76d1313c`；原始记录 `frontend-react/test-results/provenance.json`。

真实文件保留在 U02 evidence 目录：`native-hyperlinks-edited.xlsx`、`first-image.xlsx`、`sdk-range-cut.xlsx`、`sdk-financial.xlsx`、`sdk-o21.xlsx`，并有各流程截图和后端日志。本次采用 retain-on-failure trace，成功用例没有保留 trace ZIP；后续完整 UAT 需要开启 trace on，不能把已有截图称为 trace。

首次真实浏览器运行还发现两个真实链路缺陷：WebSocket 并发广播写入同一 transport 导致已提交 HTTP 动作返回失败；guest capability 被服务端错误标记为注册身份。分别由 `4779428c` 和 `9910ed01` 连贯修正，并在新增成功/拒绝合同测试后重新执行整组，当前 24/24 通过。

本记录只将已经逐项执行的 I01–I05、B03/B07、D02/D03 标记 Pass。SDK 全领域 gate、完整 browser gate 与其余领域仍为 Pending；WorkbookSession 删除、公开 DTO/错误统一、零 Web 内部依赖、结构 P0、全部小计/outline 和完整矩阵仍需实现。OIDC、desktop Excel、外部数据库继续 Blocked。
