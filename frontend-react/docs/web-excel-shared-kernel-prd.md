# Web Excel 共享内核重构与体验增强 PRD

批准时间：2026-09-07。审查基线：4fa80a2a879e6f5e0ac626f2fa45b6dd983a2cd4。实施分支：codex/web-excel-shared-kernel。

## 产品和架构决定

完全自研 Rust 共享内核，编译为浏览器 WASM 和服务端 native。Java 保留 OIDC/ACL/数据库/连接器/工作进程调度；React + Tailwind 保留视觉和交互。云端唯一持久化，删除 local-only/mirrored/offline 提交链。100万行×20列密集数据必须完整编辑，支持公式、结构操作、撤销、保存、重开。客户端正式基线是16GB、4核、SSD、64位Windows Chrome/Edge，服务器8核16GB SSD，RTT50ms/100Mbps。

范围覆盖编辑/公式/格式/筛选排序/透视表图/图表/文件/打印/权限/协作。VBA、DAX、Office外部插件不执行，文件中对应原生内容必须保留。只允许基础压缩/XML/密码学/字体/协议库，不以第三方表格、公式或分析引擎取代自研内核。

唯一链：UI intent → typed command → Rust WASM校验/草稿/PaneMap/视图 → Java鉴权/事务 → Rust native权威运算 → pages/operation/checkpoint/outbox原子提交 → committed revision/result patch → 同版本客户端副本。浏览器未提交结果必须可识别，不能冒充已保存。

页尺寸固定1024×32。类型标记、有效位图、数值向量、字符串字典、样式和公式ID按页存储。空白逻辑空间不物化。缓存未命中为DATA_PAGE_UNAVAILABLE，非空白。revision约束全部派生读；visibility单独投影。普通取消检查分区/批次边界保留缓存，异常/超时才销毁host。默认transferable，不要求SharedArrayBuffer。所有旧TS/Java语义链及消费者须在同PR删除，没有fallback、桥接、alias、doublewrite或第二事实源。

公共契约由kernel/core定义；宿主协议见kernel/host-protocol.md。manifest版本11、host协议1。v10只允许独立迁移工具读取，保留历史与未知原生内容；运行时拒绝旧版。切换冻结写入，完整迁移后原子切换，不双读双写。

## 审查事实

186个issue和148个PR全部关闭不表示完成；#334保留真实Excel/WPS Blocked且Actions运行记录0。六个公式案例真实复现：AND范围只取首值、ADDRESS模式2/3颠倒、OFFSET默认范围尺寸丢失、XLOOKUP忽略match/search、INDIRECT忽略R1C1。已有普通worksheet source注册缓存、四区FieldList、持久Worker和有界history应迁移正确语义。当前浏览器截图显示Hub假saved、Ribbon横滚、列头覆盖row1。

## 64项整改台账

已有智能体的实现汇总、局部测试证据和剩余接口风险统一收录在 [整合交接台账](shared-kernel-integration-handoff.md)。该文档是本 PRD 的整合输入，不代替原 64 项验收；局部测试通过不得提升整项状态。

本文件描述批准范围；机器可读状态在同目录web-excel-shared-kernel-status.json。not-started/implemented/verified/blocked必须按证据区分。注册、编译、synthetic self-roundtrip不算端到端验证。下面全部TODO初始未实施。

### A. 内核、存储与加载

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| A01 共享内核 | P0 | TypeScript/Java 重复工作簿语义 | Rust 同一领域内核编译 native/WASM；迁移全部消费者并删除旧算法 | WASM/native 语义差分一致；旧算法链删除 | #320,#321 |
| A02 页式单元格存储 | P0 | CellMatrix 嵌套 Map/逐格对象无法承载密集2000万格 | 1024×32列式页、页目录、有限缓存、字典/bitmap | 百万行完整编辑且内存受预算控制 | #321 |
| A03 批量事务与撤销 | P0 | 逐格 revision、clone和全量snapshot放大成本 | 页版本、copy-on-write、事务revision、增量undo | 单用户操作一次history；拒绝零部分写入 | #255,#332 |
| A04 结构变换 | P0 | 移动清除重排依赖逐格处理 | 统一range move/fill/clear与引用/对象变换 | 大区域可撤销、公式规则锚点一致 | #277 |
| A05 云端唯一持久化 | P0 | local-only/mirrored/内存保存多种语义 | 删除本地创建和离线提交；serverack唯一保存确认 | 刷新恢复已提交内容；断线停写且无假成功 | #331 |
| A06 任务调度 | P0 | 功能分别维护generation取消预算 | 统一任务状态、优先级、revision和取消检查点 | 旧结果不发布，取消有界、错误可恢复 | #329,#330 |
| A07 按需加载 | P1 | lazy面板仍依赖静态全量功能注册 | 按Hub/编辑器/分析/文件拆分装配 | Hub不加载完整计算/文件内核 | #320 |
| A08 页读取与预取 | P1 | occupied-cell/data-region可能全范围扫描 | PaneMap可见页请求、预取及空间索引 | 滚动选择局部成本；隐藏格仍可读 | #321 |

### F. 公式与数值语义

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| F01 AND/OR | P0 | 数组取首值且提前返回会遗漏错误 | 统一标量/数组/引用和错误传播 | TRUE/FALSE二维/文本/空白/错误与Excel一致 | #322 |
| F02 XLOOKUP | P0 | lazy分支遗漏match/search参数 | 完整参数策略和合法惰性求值 | 精确近似通配正反向二分均正确 | #322 |
| F03 引用函数 | P0 | ADDRESS模式反转、INDIRECT无R1C1、ROW/COLUMN引用不完整 | 统一A1/R1C1 AST与引用返回 | 六个回归案例、名称跨表及越界 | #322 |
| F04 OFFSET | P0 | 默认尺寸1且非法尺寸钳制 | 继承源范围尺寸并严格验证 | SUM(OFFSET(B1:B2,0,0))=30；非法无伪范围 | #322 |
| F05 增量依赖图 | P0 | 持久Worker仍有全量输入扫描/副本 | AST/点区间名称表索引与dirty闭包 | 单格编辑不重建全工作簿 | #322 |
| F06 大范围计算 | P1 | 整列引用与聚合materialize风险 | 页游标、共享公式模板、聚合索引 | 百万SUM/COUNTIFS无百万参数对象 | #322 |
| F07 现代函数 | P1 | LET/LAMBDA等缺失 | LET/LAMBDA/MAP/REDUCE/SCAN/BYROW/BYCOL/MAKEARRAY及动态数组 | 作用域、形状、错误、递归预算、spill | #322 |
| F08 跨域依赖 | P0 | spill/Table/Pivot/dynamic更新边界分散 | 一个依赖调度器和确定性计算上下文 | 结构更改/刷新/循环/stale均正确 | #322,#329 |
| F09 数字日期 | P1 | 格式类型与日期locale存在Partial项 | 统一1900/1904/Excel数值/格式AST，数字形文本保留文本 | 金额日期精度文本错误往返 | #323 |
| F10 能力目录 | P0 | runtime/autocomplete/OOXML能力报告分叉 | 单一函数定义生成各消费者 | 完整覆盖公开入口，不支持精确定位 | #322,#327 |

### S. 筛选、排序与规则

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| S01 筛选所有者 | P0 | 按列首匹配丢失多区域owner | 全链ownerId/range/column | A1:B10与A20:B30独立 | #324 |
| S02 筛选失败 | P0 | 解析异常转成全表隐藏 | typed failure中止事务/投影 | 损坏筛选不修改状态 | #324 |
| S03 百万行筛选 | P0 | UI/公式/菜单重复扫描 | owner/revision bitmap/列域增量索引 | 局部变更不全扫，百万筛选达标 | #324 |
| S04 筛选菜单 | P1 | 颜色/图标候选不统一其它列条件 | 排除自身条件的domain query、服务端分页搜索、虚拟列表 | 值日期颜色图标候选一致 | #87,#89 |
| S05 可见性链 | P0 | Chart/Print部分只读manual hiddenRows | 按原因统一visibility projection | 网格图表打印一致，SUBTOTAL9/109区分 | #324,#328 |
| S06 导入隐藏 | P0 | 公式缓存可能错误固化manual hidden | 导入明确来源；歧义保留并限制相关编辑 | 重算clearfilter正确，不猜缓存 | #85 |
| S07 排序元数据 | P0 | 对象式重排成本及关联复杂 | 页式稳定排序和引用metadata统一变换 | 计算值排序，批注链接规则锚点不脱离 | #254,#86 |
| S08 验证条件格式 | P1 | 规则读值缓存分散且成本风险 | 共享值/依赖/priority/stopIfTrue | 公式变更触发，非法规则原子拒绝 | #324 |

### P. 透视与分析

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| P01 列式计算 | P0 | Worker仍重建SourceRow对象 | rowId/dictionary/vector/bitmap唯一聚合 | 删除SourceRow主链 | #300,#329 |
| P02 一次聚合 | P0 | axisGroups/resultCells/resultNodes重复扫描 | 一次grouped hash聚合后稀疏转置 | 成本随input/group增长 | #300 |
| P03 稀疏投影 | P0 | 全量projection.cells及固定结果cap | 结果字典索引与viewport查询 | 未访问不物化；越界事前拒绝 | #300 |
| P04 源缓存 | P0 | worksheet缓存存在但block仍重读 | 页版本源缓存复用field/layout/filter | 同revision不重新获取全量 | #329 |
| P05 结果传输 | P1 | 普通对象树/重复provenance | 二进制结果页、路径池offset | 实际结果量成本，main无长任务 | #300 |
| P06 取消 | P0 | terminate丢sourcecache | 批次检查取消保留source | 及时取消且旧结果不覆盖 | #329 |
| P07 聚合语义 | P0 | 模型广但尚需实体验收 | 空白/错误/Count/Sum/Avg/分组/值过滤/TopN/汇总/ShowValuesAs统一 | Excel对照、文本不数值强转 | #90,#96,#105,#160 |
| P08 生命周期 | P0 | 定义目标关联多对象事务 | 一次定义+proof+collision+reference事务 | 一次history，失败无半成品 | #18,#101,#120 |
| P09 透视图查询 | P1 | series×leaf重复find | 稳定lookup及source/layout分层cache | 局部series更新、绘制不聚合 | #118,#329 |
| P10 明细控件 | P1 | 大明细/切片器/时间线边界分散 | 分页、typedmembers、显式连接、时间粒度 | 多选日期no-data联动权限 | #100,#122,#154,#159 |

### B. 文件、后端与协作

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| B01 权威导入 | P0 | browser snapshot/hash不能证明文件语义 | native从原bytes解析canonicalpages | 伪snapshot不可入库 | #331 |
| B02 artifact绑定 | P0 | PUT无workbookrevision绑定 | 唯一native save transaction，删除任意替换 | 错revision/hash/format保留原文件 | #327 |
| B03 格式能力 | P0 | edit/preserve/export混用 | 统一edit/preserve/export-only/unsupported | PDF/XPS不能冒充工作簿 | #305 |
| B04 OOXML保真 | P0 | 未知part机制缺真实producer验收 | Rust packagegraph/rels/nodes/unknown/macro preserve | xlsx/xlsm保真，不可安全编辑拒绝 | #305,#327 |
| B05 大文件 | P0 | 50MiB限/整bytes/全JSON | chunkupload、streamZIP/XML、字典分页与直接产页 | 2000万格、预算、取消无残留 | #327 |
| B06 查询流水线 | P0 | 分块容器仍全量物化全局算子 | Java批次connector+Rust投影/hashjoin/group/pivot/spillsort | boundedheap与首块延迟 | #330 |
| B07 查询绑定 | P0 | 结束时revision不代表读取时 | 起始revision/executiontoken/blockmanifest | 旧结果及取消不可load | #330 |
| B08 页checkpoint | P0 | tail有界但snapshot大对象 | 页checkpoint、增量log、幂等compaction | 单格不整表clone，gap拒绝 | #332 |
| B09 协作恢复 | P0 | PubSub不证明消息全达 | outbox+lastRevision对账+bounded补读 | 丢消息重连去重乱序正确 | #331 |
| B10 权限保存 | P0 | 模式和确认状态分散 | ACL控制commands/pages，serverack决定saved | viewer/commenter/protection无绕过 | #258,#331 |
| B11 数据库审计 | P1 | 真实三方言验证不足、无auditcursor | provider matrix及time/id游标 | 空库升级锁事务分页一致 | #333 |
| B12 对象打印 | P1 | 图形/评论/图片/打印跨链需完整验证 | 共享对象值可见性排版 | 公式筛选打印重复标题对象正确 | #326,#328,#329 |

### U. UI与交互

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| U01 网格列头 | P0 | 标题下移覆盖首行 | Rust PaneMap统一header/cell/select/editor | A1显命中编辑提交一致 | #316 |
| U02 Ribbon响应式 | P1 | 固定组宽横滚截断 | 唯一schema按组收纳不改command | 1024/1366/1920无遮挡/横滚 | #241,#304 |
| U03 设计tokens | P1 | globalfont覆盖与局部!important互盖 | Tailwind token/密度/共享组件 | 各状态排版一致 | #304 |
| U04 图标 | P1 | Fluent/Figma/selfSVG/文字符混用 | Fluent16/20/24 regular/filled语义映射 | 一致线宽光学尺寸状态 | #304 |
| U05 Hub与保存 | P0 | 假saved，模板优先文件 | cloudlogin/recent/search优先，模板入new | 各连接保存状态真实 | #286 |
| U06 公式编辑 | P1 | focus/引用/帮助待系统验收 | 统一EnterEscTabIMEPointF4completion | 键盘闭环焦点正确 | #301 |
| U07 透视面板 | P1 | 四区已存在，信息高级项待整理 | 保留四区、搜索拖放键盘重复值字段defer/error | 一次Apply，失败保留draft/lastvalid | #36,#37,#300 |
| U08 分析面板 | P1 | 大domain/条件/task状态需一致 | virtual日期层级摘要进度cancel/recovery | 无无限loading或瞬时错误 | #87,#329 |
| U09 图表对象 | P1 | 入口与内部标签体验欠缺 | context/properties/elementediting统一 | inserteditdeleteundoreopen闭环 | #311 |
| U10 a11y与语言 | P1 | CanvasAX及中英混用 | 可见语义grid、livecell、focus/keytips、locale | 键盘读屏操作全链 | #16,#22,#317 |
| U11 字体资源 | P1 | CJK font-display:block | UI字体分包与文档字体分离 | 首屏无空字，替代明确 | #295 |

### D. 交付

| TODO | 优先级 | 现状/影响 | 统一方案 | 验收 | 关联 |
|---|---|---|---|---|---|
| D01 PRD台账 | P0 | 历史文档完成与证据不一致 | 本PRD唯一执行台账，历史按被替代标记 | 每todo证据与状态唯一 | #333 |
| D02 可复现构建 | P0 | lock被忽略，脚本旧目录 | 固定Node/Rust/Java/Maven，lock与统一构建 | cleancheckout四端artifact | #333 |
| D03 CI | P0 | 无Actions | Rust/WASM/diff/frontend/backend/db/browser/perf | 失败阻交付，不continue-on-error | #333 |
| D04 Excelcorpus | P0 | 上一轮真实ExcelBlocked | 真实producer+hash+reopen差异 | 不能synthetic自验替代 | #305,#333 |
| D05 一次PR交付 | P0 | closed不等于达标 | 同codex分支64项全过后一个PR | 契约迁移证据状态整体rollback | #334 |

## UI规格

文档栏36px，Tab28px，Ribbon76px，公式栏32px；sidebar340px可调300–480px，四区独立scroll底部操作固定。UI13px，绿#107C41，正文#242424，面板#F5F5F5，边框#D1D1D1。Fluent16/20/24图标，主要目标至少28px。1366以下语义组收纳并支持折叠Ribbon，业务组件只用共享组件和Tailwind。先生成完整editor/filter/pivot/chart/hub/loading/error设计基线并固定golden，再逐像素对照；不改变既有真实业务能力。工作簿字体与UI字体分离，替代明确报告。

## 性能合同（目标，尚未实测）

| 场景 | P95目标 |
|---|---|
| server已存百万行首可编辑viewport | <=3s，不下载全量cells |
| scroll/select frame | <=16.7ms；feedback<=100ms |
| 单格servercommit | <=250ms，成本不随全历史增长 |
| 百万行常规filter / indexed member首屏 | <=2s / <=200ms |
| 百万行sort | <=5s，可取消UI可交互 |
| 百万行pivot，1万groups、<=100columns | first<=5s，layout reuse<=2s |
| cancel | <=200ms；不发布业务结果 |
| 单格<=1000公式依赖 | <=200ms，跨端commit另计 |
| 百万简单独立公式 | <=10s，可观察可取消 |
| browser内存峰值 | <=1.5GiB |
| native单任务 | 2GiB，受控spill |
| 2000万格import | upload完成后<=60s，progress间隔<=250ms |

文件预算：1GiB compressed/8GiB expanded/16GiB temporary，XMLnodes/cells/objects/rels独立预算，分配与提交前检查。禁止仅提高上限掩盖物化。

## 验收和交付

一次完整开发批次后集中验证：Rust native/WASM语义差分，frontend typecheck/unit/contracts/boundaries/build，backend Maven与H2/Postgres/MySQL migration/事务矩阵，真实OIDC身份、双客户端、console/network、1024/1366/1920与100/125/150%视觉、百万行dense/sparse/multisheet/高基数、Excel producer XLSX/XLSM导入→编辑→导出→真实Excel重开。失败/越权/缺页/hash/旧协议/stale/cancel/crash/budget必须零部分提交。SUBTOTAL9/109等按visibility原因保留不同语义。

所有64项须实现且适用验收通过。缺真实Excel或业务环境保持Blocked/尚未完全验收，不能计为通过。仅在ABI/协议/依赖会阻止继续开发时中间检查，不频繁全量编译。PR记录实现、删除路径、迁移、commit+工具链+证据、Blocked、整体rollback；应用和迁移前数据一起回滚，禁止单独回滚消费者。不直接main提交或推送。
