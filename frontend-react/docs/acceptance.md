# React Sheets 验收准则

## 离线工作簿

在干净浏览器 profile 中只启动 Web，不启动 Java、PostgreSQL、Redis 或 OIDC。

1. 新建本地工作簿，双击、F2、Formula Bar 和中文输入法分别输入文本、数字、公式及首尾空格。
2. 验证 Enter、Shift+Enter、Tab、Escape、Delete、复制、剪切、HTML/TSV 粘贴、格式、行列、工作表、Undo/Redo。
3. 编辑公式依赖、动态数组、条件格式、数据验证、Table、Outline、Drawing、Chart、Pivot、Sparkline、Review、Print 和 Query 定义。
4. 等待 checkpoint 后硬刷新、关闭并重新打开；比较工作簿快照 checksum、公式文本和值、sheet order、对象、Review、Print 与 Query 定义。
5. 从 UI 导入真实 XLSX、编辑、导出、重新导入；导出 PDF 必须以 `%PDF-` 开头并包含 Unicode 文本。
6. Network 不得出现 API 或 WebSocket 请求；状态栏仅显示本地已保存、离线待同步、错误或计算中。

## 在线协同

使用 Java backend、PostgreSQL、OIDC 和多实例 Redis 配置，启动两个独立浏览器上下文。

1. owner 创建工作簿并授权 editor/commenter/viewer，创建可撤销、可过期的匿名共享链接。
2. 两端并发编辑同一单元格、插入/删除行列、移动 Drawing、修改 Pivot/Sparkline、断网后重连和协同 Undo。
3. 验证 server 生成 actor、revision、commit time 和 affected ranges；客户端伪造 actor、role、range 或 snapshot 必须被拒绝且 revision 不增加。
4. 验证匿名访客只能按共享角色交互；无共享凭据的匿名 REST/WebSocket 请求必须拒绝。
5. 验证 Redis 仅用于跨实例 broadcast/presence；关闭 Redis 的多实例部署必须 fail-closed，单实例不依赖 Redis。

## 通过条件

所有用户动作必须产生已注册 command/mutation，具有 inverse、权限、affected range、远端重放和重启恢复证据。任何兼容转发层、alias、版本后缀、直接 Workbook 写入、旧 Node/Java 服务或未注册 mutation 都是失败。
