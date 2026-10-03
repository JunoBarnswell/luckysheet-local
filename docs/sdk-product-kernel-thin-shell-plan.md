# SDK 产品内核与 Web 薄壳整改

## 本轮事实基线

- 本轮用户明确要求严格执行所附完整任务文档，并创建新的草稿 PR。
- 唯一源码基线：2026-10-03 通过 GitHub Connector 重新获取的 `main`，`3ff47ca6df8de4af49b809649831ea376179d24c`。
- 提交：`Merge pull request #349 from JunoBarnswell/codex/sdk-product-kernel`。
- 本地 `origin/main` 已 fetch 到同一 SHA；旧功能分支与该 main 的源码 tree 没有差异。本轮分支为 `codex/sdk-product-kernel-thin-shell`。
- [本轮已明确采纳的任务文档](sdk-product-kernel-thin-shell-request.md)定义交付范围；下面的架构安排不得降低其完成标准。
- 重新统计非测试 `apps/web/src` 文件：spreadsheet-app 47、core-model 60、command-runtime 17、render-engine 8、protocol 6、exchange-excel-ooxml 6、sheet-features 5、formula-engine 2。按每种包引用的唯一文件数统计，包含 inline import type。
- 真实 `WorkbookSession`：412,453 bytes，8,192 行；它必须被职责分解后删除。
- 保留 React/TypeScript + Java 21 + H2、canonical WorkbookRole、verified auth context、统一 Command/Mutation/Model、跨工作簿授权计算、未知 OOXML 与宏保真边界。

## 落代码前的完整架构审查

先逐个登记旧字段、方法、调用者和依赖，明确唯一 owner；读清模型、命令、服务端、历史、协作、导入导出之间的契约，再一次完成连贯实现。任何验证失败先冻结证据、审查完整链，再做连贯修正。不得按错误逐点放宽约束。

1. 公共对象 `Workbook` 仅保存 identity、lifecycle、领域组合、transaction context、view publication、资源释放。对象 API 和交互编辑器使用同一个 Workbook 和 canonical runtime。
2. 每个领域拥有真实状态和动作实现。跨领域通过明确的内部依赖调用；不得把旧 Session 改名、保留转发对象、生成大对象继承层或另建模型。
3. 生命周期负责启动/退休以及上下文、对象和异步任务代际；权限状态由服务端授权与原有 canonical protection/access resolver 管理。身份切换销毁旧工作簿、投影、任务、资产 URL、日志与连接。
4. `WorkbookView` 是唯一 reactive 发布 owner，缓存订阅 snapshot 的 identity，组合各领域只读投影。投影不得携带命令描述符、Mutation、协议客户端或可写模型；不得将内部 `UiSnapshot` 直接改名后暴露。
5. 公共 SDK DTO 显式定义，内部转换；目录 actions 显式 interface。公共 API 错误统一为 `SdkError`，保留内部错误 cause、operation、object、status 与 recovery。
6. Guest share capability 在 SDK Auth 初始化捕获并从浏览器 URL 清除；协议及 WS 仅读 SDK 内部源。`identity` 为当前 verified identity；管理用户归 `users`。
7. SDK core 依赖闭包不加载 React。React hooks/providers 仅在 `@react-sheets/sdk/react` 入口。`useWorkbook(sdk, unitId)` 由 SDK 完成 resolve/open，返回 Workbook/View/领域，不向 Web 返回内部 runtime。
8. 将命令构造、selection/keyboard/pointer、drawing hit test/chart/pivot projection、渲染调度和所有 Excel 事务迁入真实 SDK 领域；Web 只保留展示、路由、浏览器宿主原语。
9. `apps/web/package.json` 业务依赖收口为 SDK、UI system、React、ReactDOM；移除全部八种内部包引用及旧入口，重写 boundary gate 并加入拒绝样例测试。
10. 跨 sheet move/reference graph、table participant、structural history rebase、服务端 DV calculation、subtotal 与 row visibility 必须覆盖前端 canonical owner、Java authority、inverse/history、协同、持久化和 OOXML。禁止 copy-clear-write、保留核心 UNSUPPORTED 占位、客户端 proof 替代服务器计算或用 invalidate 隐藏错误。

## 领域状态 ownership

| 领域 | 唯一状态/资源 owner | 动作写链与投影责任 |
|---|---|---|
| Lifecycle | 工作簿身份、phase、启动/退休代际、资源释放 | 内部 resolution、运行时创建、所有任务与连接终止 |
| View | 订阅者、snapshot generation、readonly publication | 整合全部领域投影，保证 stable snapshot、权限遮蔽 |
| Selection | active sheet、分组 sheet、selection/gesture | selection canonical owner，合并单元格、隐藏轴、历史/远端变换 |
| Editing | cell edit、IME/caret、autocomplete tasks、overlay | 同一个 CellEditDomain；公式引用、输入解释、提交/取消 |
| Clipboard | 私有 clipboard、系统 formats/status | copy/cut/paste-special；cut 由正式 structural move |
| Formula | 计算、audit、external links、catalog | 继续使用 sole formula engine，保留授权 faults 和跨书图 |
| Cells/Range | canonical cell resolver、显式地址对象行为 | 授权读取、输入/格式/合并/fill/shift，禁止第二份数据 |
| Worksheet | identity/extent/visibility/axis/panes | 同一个 command、structural planner 与 model |
| Names | canonical defined names | scope/anchor/rename/reference graph 与 native roundtrip |
| Tables | table identity/schema/columns/filter/style | table structural transaction participant、structured refs |
| Data | 数据源、materialization、query tasks、排序/outline | 权限/代际/版本 fence；源块与投影不复制真相 |
| Validation/CF | canonical rule collections | 唯一命令 reducer，规则 preimage/范围/服务器判定 |
| Drawing/Chart | drawing selection、textbox placement、asset URL、chart element | 唯一 drawing model、insert IDs、bounds、payload/history |
| Pivot | task ports、registered sources、active tasks、field loads | 所有任务版本 fence；controls、layout、refresh、drilldown |
| Sparkline | canonical sparkline/group collections | 同一个 source/formula owner、commands/native |
| Review | canonical notes/comments/hyperlinks | viewer/commenter/editor/owner 各能力、review lifecycle |
| Permissions | verified role、effective access、protection | 服务端 authority 优先；range/object/cell permission |
| Collaboration | socket lifecycle、peers、revision、pending | contextual auth fence、resync、remote structural transform |
| History | undo/redo、revision list/preview、replay state | exact inverse preimages，远端 rebase、invalid 真错误可见 |
| Document | save state、native artifact、compatibility、checkpoint | flush/calc/auth fences，统一 save/export/native capture |
| Print | print setup/area/title/break 与 preview | canonical print model、PDF/export，不另建真相 |
| Viewport/Render | PaneMap、geometry、hit test、调度、floating projection | SDK 接收浏览器事件/尺寸，Web 不计算 Excel 语义 |
| UI Chrome | ribbon/panel/dialog/backstage/focus/notice | SDK 交互投影与显式动作；Web 仅展示 |

## 新 PR 与验收交付

每完成一个连贯功能，提交并推送到同一新草稿 PR。保留实现说明、受影响契约、删除的旧设计、迁移边界、验证证据、阻塞及 rollback。按照预先设计的 [逐项 UAT](sdk-product-kernel-thin-shell-uat.md)实际执行，逐项更新；Pending/Fail/Blocked 均不能计作 Pass。全量门禁未通过、结构或域仍有缺口时，继续保持 draft。

真实 OIDC/ERP SSO 与 desktop Excel 环境已由用户确认没有，本轮仍记录 Blocked；不再反复索要。Java PostgreSQL/MySQL 外部测试环境缺失也须单独注明，不能等同于产品成功或用本地假环境替代。

当前 main 已知上一轮完整前端单测存在 40 项失败，真实 SDK browser UAT 23/23；这是对应旧交付的历史证据。本轮会在完整实现后重新运行，不挪用旧结果声明新分支通过。
