# 导出保真度（LuckySheet → xlsx）

> Phase 4。本地导出走 Java `POST /luckysheet/luckyToXlsx`（luckysheet-lib），**不依赖 Univer Server / `@univerjs-pro` / Node**。

## 1. 当前链路

1. 前端插件 `frontend/src/expendPlugins/exportXlsx/plugin.js` 把 `luckysheet.toJson()` POST 到 `plugins[{name:'exportXlsx'}].config.url`（默认 `/luckysheet/luckyToXlsx`）。
2. Java `ExcelIoService` 用 luckysheet-lib 写 `.xlsx` 并返回二进制。
3. 失败时插件走真实回调：空 url、HTTP 非 2xx（含 `{ error }` JSON）、空/非表格响应、网络错误，不再静默成功。

## 2. 本轮已覆盖

| 类别 | 写出内容 |
|---|---|
| 值 | `v` / `m` |
| 公式 | `f` → Excel 公式；若同时有 `v` 则写入 `result`，打开时先显示缓存值 |
| 合并 | `config.merge`（`r/c/rs/cs`） |
| 字体 | `ff/fs/bl/it/fc/cl/un` |
| 填充 | `bg` |
| 对齐 | `ht/vt/tb/tr` |
| 数字格式 | `ct.fa`（非 General） |
| 边框 | `config.borderInfo`（range / cell） |
| 行列尺寸 | `config.rowlen/columnlen`，隐藏行/列 |
| 冻结 | `frozen` / `freezen` 的行列冻结（近似） |
| 数据源 | 优先 `sheet.data`；没有则从 `celldata` 重建 |

## 3. 与 Excel 完整 fidelity 的差距（明确不做或未对齐）

这些不是「漏写一行就能齐」，而是 LuckySheet 模型与 OOXML 不是同一套。

| 缺口 | 原因 |
|---|---|
| 图表 | ChartMix / ECharts 浮层，不是 Excel drawing chart；`chartMap` 仅附加 base64，服务端当前不写入 xlsx drawing |
| 图片 | `imageCtrl` 是工作表浮层，未写入 xl/media |
| 数据透视表 | 本仓是自研 `pivotDatas` 网格，不是 Excel PivotCache |
| 条件格式 / 数据验证 / 筛选 | JSON 字段在，导出未映射到 `cfRule` / `dataValidations` / `autoFilter` |
| 批注 / 超链接 | `ps` / `hyperlink` 未写入 comments / hyperlinks |
| 富文本 inline string | `ct.s` 多段样式未拆成 Excel rich text run |
| 斜线/复杂边框、主题色 | LuckySheet 边框枚举只映射到 ExcelJS 常用 style |
| 打印页设置 | 打印插件 Blocked，无页边距/页眉页脚/纸张 |
| 数组公式 / 动态数组 | 只写普通 `f` |
| 工作簿保护 / 工作表保护密码 | `config.authority` 不导出 |
| 协同 OT 历史 | 导出是快照，不含操作日志 |
| 像素 ↔ 磅/字符宽 | 行高 `px * 0.75`、列宽 `px / 8`，和 Excel 量测有系统误差 |

## 4. 验收

- 带公式、合并、背景色、加粗、边框的 sheet POST 到 `/luckysheet/luckyToXlsx`，用 Excel / WPS 打开后公式栏可见公式、合并区正确、颜色和加粗在。
- `config.url` 为空：对话框不发出请求，提示「未配置导出地址」。
- 服务不可达或返回 HTML/空 body：失败回调触发，loading 关闭，不下载坏文件。
