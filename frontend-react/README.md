# React Sheets

独立的 React + Canvas 电子表格。启动首先进入 `/workbooks` 文件中心；Catalog 管理多个工作簿资源，编辑器路由 `/workbooks/:unitId` 始终只挂载一个 `WorkbookSession`。未连接服务端时可创建/打开本地工作簿、编辑、导入/导出 XLSX；云端目录、空间/文件夹、共享、ACL、回收站和协同由唯一 Java `backend/` 提供。

## 唯一链路

```text
UI intent
→ CommandDescriptor
→ CommandRuntime transaction
→ typed mutation + inverse
→ WorkbookModel / FormulaEngine
→ Page-session memory workspace checkpoint + operation journal
→ REST `POST /api/workbooks/{unitId}/operations`
→ durable backend commit + WebSocket `revision.created` broadcast
```

没有 Node 服务、旧 Luckysheet 协议、双模型、原地 XLSX 替换或 Catalog/编辑器兼容桥。

## 开发

```powershell
npm install
npm run dev
```

Web：`http://127.0.0.1:4180/`

复制 `.env.example` 并由部署环境提供 OIDC public-client 配置。浏览器使用 Authorization Code + PKCE，不在工作簿或 URL 中存储 bearer token。本地工作簿只存在于当前页面内存会话，刷新或关闭页面后清空。开发时 Vite 将 `/api` 和 `/ws` 同源代理到 `REACT_SHEETS_API_ORIGIN`；生产环境也必须保持同源。

详细的文件中心模型、空间/文件夹、权限、导入导出和像素验收基线见 [workbook-hub-prd.md](docs/workbook-hub-prd.md)。

## 验收

离线和在线验收步骤见 [acceptance.md](docs/acceptance.md)。构建或单测不能替代真实浏览器、重启恢复与双客户端协同验证。
