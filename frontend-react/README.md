# React Sheets

独立的 React + Canvas 电子表格。前端是 local-first PWA：未连接服务端时可打开本地工作簿、编辑单元格、计算公式、导入/导出 XLSX、打印 PDF、执行本地文件型查询和 What-if；协同、匿名共享、可信 ACL、审计与数据库查询只在唯一 Java `backend/` 连接后启用。

## 唯一链路

```text
UI intent
→ CommandDescriptor
→ CommandRuntime transaction
→ typed mutation + inverse
→ WorkbookModel / FormulaEngine
→ IndexedDB workspace checkpoint + operation journal
→ optional Java operation commit and WebSocket broadcast
```

没有 Node 服务、旧 Luckysheet 协议、`pro-features`、双模型或兼容桥。

## 开发

```powershell
npm install
npm run dev
```

Web：`http://127.0.0.1:4180/`

连接 Java 服务时，由部署环境提供 OIDC token、PostgreSQL 连接与可选 Redis。浏览器不保存数据库凭证；JSON、CSV、TSV、XLSX 文件型查询可在本地执行，数据库查询由后端执行。

## 验收

离线和在线验收步骤见 [acceptance.md](docs/acceptance.md)。构建或单测不能替代真实浏览器、重启恢复与双客户端协同验证。
