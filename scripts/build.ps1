$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot

Write-Host "==> Building Luckysheet frontend..."
Push-Location "$Root\frontend"
npm install
npm run build
Pop-Location

Write-Host "==> Copying frontend dist into backend static resources..."
$StaticDir = "$Root\backend\luckysheet\src\main\resources\static"
if (Test-Path $StaticDir) { Remove-Item -Recurse -Force $StaticDir }
New-Item -ItemType Directory -Path $StaticDir -Force | Out-Null
Copy-Item -Recurse -Force "$Root\frontend\dist\*" $StaticDir

Write-Host "==> Building Luckysheet backend (MySQL profile)..."
$MavenHome = "$Root\tools\apache-maven-3.9.9"
if (-not (Test-Path "$MavenHome\bin\mvn.cmd")) {
    throw "Maven not found at $MavenHome. Download apache-maven-3.9.9 to tools/ first."
}
$env:MAVEN_HOME = $MavenHome
$env:Path = "$MavenHome\bin;" + $env:Path
$env:MAVEN_OPTS = @(
    "--add-opens=jdk.compiler/com.sun.tools.javac.api=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.code=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.comp=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.file=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.main=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.model=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.parser=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.processing=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.tree=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.util=ALL-UNNAMED",
    "--add-opens=jdk.compiler/com.sun.tools.javac.jvm=ALL-UNNAMED"
) -join ' '

Push-Location "$Root\backend"
mvn clean package -Pmysql -DskipTests
Pop-Location

Write-Host "==> Done."
Write-Host "Frontend dist : $Root\frontend\dist"
Write-Host "Backend jar   : $Root\backend\luckysheet\target\web-lockysheet-mysql.jar"
