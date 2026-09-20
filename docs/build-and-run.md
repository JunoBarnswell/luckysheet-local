# React Sheets 构建、运行与 Windows 安装

浏览器版由 `frontend-react` 生成 `dist/web`，Java 21 Spring Boot 服务托管前端静态资源、API 和 WebSocket。构建不会改写 `frontend-react` 或 `backend/src/main/resources`；前端产物只会暂存到 `backend/target/generated-web`，由 Maven 复制到 JAR 的 `static/`。

## 开发

先准备 Node 24、JDK 21 和 Maven 3.9。JDK 21 通过 `JAVA_HOME` 或 `-JavaHome` 选择；脚本会拒绝 Java 17、Node 20 或其他主版本。

```powershell
cd frontend-react
npm ci
npm run dev
```

Vite 地址是 `http://127.0.0.1:4180/`，`/api` 和 `/ws` 代理到 `http://127.0.0.1:8082`。Java 服务可在另一个终端运行：

```powershell
cd backend
mvn spring-boot:run
```

完成一次生产构建后，可用 `node scripts/dev-server.js` 在 `4181` 启动带 Java 代理的静态预览。端口可通过 `REACT_SHEETS_PREVIEW_PORT`、`REACT_SHEETS_BACKEND_PORT` 和 `REACT_SHEETS_WEB_ROOT` 覆盖。

## 构建

```powershell
.\scripts\build.ps1 -JavaHome $env:JAVA_HOME -SkipTests
```

脚本依次执行 `npm ci`、`npm run build`、`mvn clean`、复制 `frontend-react/dist/web`、`mvn package`，每个外部命令保留退出码。完整日志写到系统临时目录并只显示失败日志尾部；可用 `-LogRoot` 指定仓库外日志目录。默认会运行 Maven 测试，交付构建可显式传 `-SkipTests`。

## Windows 安装包

`installer/dependencies.json` 固定 Temurin JRE 21.0.12.1、WinSW 2.12.0 和 NSIS 3.12 的官方发布 URL 与 SHA-256。`build-installer.ps1` 首次使用时把这些依赖下载到 `%LOCALAPPDATA%\ReactSheets\build-cache`，先校验 SHA-256，再解压或使用；校验失败会停止构建，不会继续打包。当前过程只生成安装程序，不会安装服务或修改 Windows 服务状态。

```powershell
.\scripts\build-installer.ps1 -JavaHome $env:JAVA_HOME
```

产物位于 `backend/target/installer/ReactSheets-Setup-<version>.exe`。安装程序包含 Java 运行时、WinSW、JAR 和静态前端，并执行以下操作：

- 在 `%ProgramFiles%\React Sheets` 安装程序文件；
- 在 `%ProgramData%\ReactSheets` 创建独立的 `data`、`config`、`backups` 和 `logs` 目录；
- 注册 `ReactSheets` 自动启动服务，使用内置 JRE 21；
- WinSW 日志按大小和时间轮转，保留 14 个文件；
- 服务启动后轮询 `/health`，失败则中止安装；
- 升级先停服务，再调用备份脚本；现有配置不覆盖；卸载移除程序和服务但保留 ProgramData 数据。

安装程序默认监听 `127.0.0.1:8082`。局域网访问、HTTPS、认证和 `WEB_ALLOWED_ORIGINS` 应由部署配置明确开启，不由安装器猜测。

## 备份和离线恢复

H2 文件复制必须在服务停止后进行。备份脚本写入带 `manifest.json` 的 ZIP，清单包含每个文件的长度和 SHA-256：

```powershell
Stop-Service ReactSheets
.\scripts\backup-data.ps1
Start-Service ReactSheets
```

恢复要求显式 `-Force`，会在替换前保留 `data.before-restore-<timestamp>`，并在清单校验失败时停止，不创建空数据库掩盖损坏：

```powershell
Stop-Service ReactSheets
.\scripts\restore-backup.ps1 -BackupPath .\react-sheets-backup-20260921-120000.zip -Force
Start-Service ReactSheets
```

如果提供 `-ExpectedSha256`，脚本还会先校验整个备份包的 SHA-256。桌面 Excel 往返、5 人协作和 5 万行性能属于独立验收，不由构建成功代替。
