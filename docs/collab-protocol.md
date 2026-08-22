# LuckySheet WebSocket 协同协议

> Phase 4。文档化本仓库 **已有** `t / i / v / rv` 协议，并说明本轮补上的版本号、切 sheet 边界、断线重连。  
> **明确不实现 Univer Pro OT**（无 transform、无意图保留、无中心 operation log）。冲突策略是带时间戳的 last-write-wins + 失败重拉，不是 OT。

## 1. 连接

- 开关：`allowUpdate === true` 且配置了 `updateUrl`、`gridKey`。
- 入口：`frontend/src/controllers/server.js` → `openWebSocket()`。
- URL：`updateUrl + "?t=111&g=" + encodeURIComponent(gridKey)`；若 `updateUrl` 已带 `?` 则改用 `&`。
- 后端（Java Luckysheet Server）：`/luckysheet/websocket/luckysheet`，见 `backend/luckysheet/.../websocket/MyWebSocketHandler.java`。
- 载荷：除心跳外，客户端把 JSON `pako.gzip(..., {to:"string"})` 后发送；服务端解压再 `JSONObject.parseObject`。
- 心跳：连接成功后每 60s 发送明文 `"rub"`，服务端只记日志，不广播。

## 2. 服务端回包信封

`backend/README-zh.md` 与 `MyWebSocketHandler` 一致：

```json
{
  "createTime": 0,
  "data": "{}",
  "id": "7a",
  "returnMessage": "success",
  "status": "0",
  "type": 2,
  "username": "name",
  "message": ""
}
```

| type | 含义 |
|---|---|
| 0 | 连接成功（部分实现） |
| 1 | 只回给发送者：指令落库成功/失败；新建 sheet 时可能改写 `index` |
| 2 | 广播给其他用户：`data` 为原始指令 JSON 字符串 |
| 3 | 选区 / 光标（`t=mv`） |
| 4 | 批量指令（历史字段；`data` 可能是空串） |
| 5 / 6 | 显示 / 隐藏 loading |
| 999 | 用户断开（文案「用户退出」时前端摘掉该用户选区框） |

`status`：`"0"` 需要按 `data` 改本地；`"1"` 无意义/失败。

## 3. 客户端上行指令（`saveParam`）

公共字段：

| 字段 | 含义 |
|---|---|
| `t` | 操作类型 |
| `i` | 目标 sheet `index`（不是数组下标） |
| `v` | 载荷 |
| `ver` | 本客户端会话单调序号（本轮新增，后端可忽略） |
| `ts` | 客户端 `Date.now()`（本轮新增，LWW 用） |

### 3.1 单元格

| t | 附加字段 | v |
|---|---|---|
| `v` | `r`, `c` | 单格对象或 `null` |
| `rv` | `range: {row:[r1,r2], column:[c1,c2]}` | 二维数组，一次最多约 1000 格 |
| `rv_end` | 无 | `null`，表示本批 `rv` 结束（后端写库信号） |

`historyParam` 按列数切批：`timeR = floor(1000 / collen)`。

### 3.2 配置与通用字段

| t | 附加 | 说明 |
|---|---|---|
| `cg` | `k` | `config` 子项：`rowlen` / `columnlen` / `rowhidden` / `merge` / `borderInfo` 等 |
| `all` | `k` | sheet 根字段：`name` / `color` / `pivotTable` / `frozen` / `filter` / `images` / `dataVerification` / `hyperlink` / 条件格式等 |
| `fc` | `op`, `pos` | calcChain 增删 |
| `drc` / `arc` | `rc` | 删/增行列 |
| `f` / `fsc` / `fsr` | `op`, `pos` | 筛选更新 / 清除 / 恢复 |
| `c` | `cid`, `op` | 图表 add/xy/wh/update/del（遗留） |
| `na` | | 工作簿标题 |

### 3.3 Sheet 操作

| t | 说明 |
|---|---|
| `sha` | 新建 |
| `shc` | 复制 |
| `shd` | 删除 |
| `shr` | 重排序 |
| `shre` | 恢复删除 |
| `sh` | 隐藏/显示（`op`, `cur`） |
| `shs` | **激活当前 sheet**。`v` = 新 `currentSheetIndex` |

切 sheet 边界（#214 / 源码原 TODO）：

- 旧代码：`t==='shs'` 直接 `return`，后台不知道谁停留在哪一页。
- 现行为：客户端 **发送** `shs`，供 Java `Operation_shs` 更新文档激活状态。
- 其他客户端收到 type=2 的 `shs` **不得切换自己的当前视图**（`wsUpdateMsg` 对 `shs` 显式忽略）。这不是 OT，只是「激活状态入库、视图仍本地」。

## 4. 版本号与冲突（非 OT）

本轮在每条 `saveParam` 上附带 `ver` + `ts`。

| 规则 | 行为 |
|---|---|
| 发送 | `collabSeq++` 写入 `ver`，`ts=Date.now()` |
| 同格后写覆盖 | 对 `v`/`rv` 用 `i_r_c` 记录最后 `ts`；更旧的远程写丢弃 |
| 发送失败 type=1 / `returnMessage=error` | 不前进 ack；消息若仍在 `outboundQueue` 会在重连后重发 |
| 版本空洞 | **不**做操作变换，也不假装因果序；需要权威状态时走 `loadUrl` 重拉整表 |

没有：transform、intent、共享 version vector、中心 snapshot rebase。那是 Univer Pro 协作。本仓 Java 服务端仍是「加 Redis 锁后直接改块」。

## 5. 断线重连

| 事件 | 行为 |
|---|---|
| `onopen` | 清错误计数，启心跳，`flushOutboundQueue()` |
| 心跳 | 60s `"rub"` |
| `onerror` | 不立刻叠一条新 socket；交给 `onclose` |
| `onclose` code=1000 或主动 `closeWebSocket(true)` | 停重连 |
| 其他 close / 掉线 | 指数退避（1s 起，封顶 30s），最多 20 次，再提示刷新 |
| 离线发送 | `websocket` 非 OPEN 时入 `outboundQueue`（上限 200），避免 #213 直接丢指令 |
| `online` 事件 | 若已断开则立即重连 |

重连 **不会** 自动全量 `loadUrl`（避免覆盖本地未确认编辑）。队列里的 gzip 包按原序重发；若服务端已处理过，以服务端块数据为准。

## 6. 前端应用远程写（`wsUpdateMsg`）

- 目标 sheet 的 `file.data` 尚未加载时，`v`/`rv` 仍直接 return（历史行为）；未加载页依赖之后 `loadSheetUrl` / `buildGridData`。
- 仅当 `item.i == Store.currentSheetIndex` 才 `luckysheetrefreshgrid()`。
- 其他人的选区用 type=3 + `multipleRangeShow`；切走该 sheet 时隐藏其框。

## 7. 验收

- 两客户端同格先后编辑：后到的 `ts` 留下，先到的过期 `v` 不覆盖。
- 切 sheet：自己视图变，另一人视图不变；网络面板能看到 `shs`。
- 断网后再连：队列中的 `v`/`rv` 会再发出；不是静默丢。
- 协议仍是 `t/i/v/rv`，没有 OT 消息类型。
