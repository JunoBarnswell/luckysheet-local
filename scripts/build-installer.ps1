[CmdletBinding()]
param(
    [string]$JavaHome,
    [string]$MavenCommand = 'mvn.cmd',
    [string]$OutputDirectory,
    [string]$LogRoot,
    [string]$NsisPath,
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$frontendRoot = Join-Path $repositoryRoot 'frontend-react'
$backendRoot = Join-Path $repositoryRoot 'backend'
$installerRoot = Join-Path $repositoryRoot 'installer'
$dependencyManifestPath = Join-Path $installerRoot 'dependencies.json'
$dependencyManifest = Get-Content -LiteralPath $dependencyManifestPath -Raw | ConvertFrom-Json
$version = ([string](Get-Content -LiteralPath (Join-Path $frontendRoot 'package.json') -Raw | ConvertFrom-Json).version)

if ([string]::IsNullOrWhiteSpace($LogRoot)) {
    $LogRoot = Join-Path ([IO.Path]::GetTempPath()) ('react-sheets-installer\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $backendRoot 'target\installer'
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$cacheRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReactSheets\build-cache'
$stageRoot = Join-Path $backendRoot 'target\installer-stage'

function Invoke-LoggedCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$Arguments,
        [Parameter(Mandatory)][string]$WorkingDirectory
    )
    $logPath = Join-Path $LogRoot (($Name -replace '[^A-Za-z0-9_.-]', '-') + '.log')
    Write-Host "==> $Name"
    Write-Host "    log: $logPath"
    Push-Location $WorkingDirectory
    try {
        & $FilePath @Arguments *> $logPath
        $exitCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }
    if ($exitCode -ne 0) {
        Get-Content -LiteralPath $logPath -Tail 60 -ErrorAction SilentlyContinue
        throw "$Name failed with exit code $exitCode. Full log: $logPath"
    }
}

function Get-VerifiedArtifact {
    param(
        [Parameter(Mandatory)]$Specification,
        [Parameter(Mandatory)][string]$CacheDirectory
    )
    New-Item -ItemType Directory -Path $CacheDirectory -Force | Out-Null
    $path = Join-Path $CacheDirectory ([string]$Specification.fileName)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        Write-Host "Downloading $($Specification.fileName) from the pinned release URL."
        Invoke-WebRequest -UseBasicParsing -Uri ([string]$Specification.url) -OutFile $path
    }
    $actualHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash.ToLowerInvariant()
    $expectedHash = ([string]$Specification.sha256).ToLowerInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "DEPENDENCY_CHECKSUM_MISMATCH: $($Specification.fileName), expected $expectedHash, actual $actualHash"
    }
    return $path
}

function Find-MakeNsis {
    param([string]$RequestedPath)
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) { $candidates.Add($RequestedPath) }
    if (-not [string]::IsNullOrWhiteSpace($env:MAKENSIS)) { $candidates.Add($env:MAKENSIS) }
    if (-not [string]::IsNullOrWhiteSpace($env:NSIS_HOME)) { $candidates.Add((Join-Path $env:NSIS_HOME 'makensis.exe')) }
    $pathCommand = Get-Command makensis.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $pathCommand) { $candidates.Add($pathCommand.Source) }
    foreach ($candidate in $candidates | Select-Object -Unique) {
        $resolved = if (Test-Path -LiteralPath $candidate -PathType Leaf) { [IO.Path]::GetFullPath($candidate) } elseif (Test-Path -LiteralPath (Join-Path $candidate 'makensis.exe') -PathType Leaf) { [IO.Path]::GetFullPath((Join-Path $candidate 'makensis.exe')) } else { $null }
        if ($null -eq $resolved) { continue }
        $output = & $resolved /VERSION 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0 -and $output -match '3\.12') { return $resolved }
    }

    $archive = Get-VerifiedArtifact -Specification $dependencyManifest.nsis -CacheDirectory $cacheRoot
    $extractRoot = Join-Path $cacheRoot ('nsis-' + [string]$dependencyManifest.nsis.version)
    $makeNsis = Get-ChildItem -LiteralPath $extractRoot -Filter 'makensis.exe' -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $makeNsis) {
        if (Test-Path -LiteralPath $extractRoot) { Remove-Item -LiteralPath $extractRoot -Recurse -Force }
        New-Item -ItemType Directory -Path $extractRoot -Force | Out-Null
        Expand-Archive -LiteralPath $archive -DestinationPath $extractRoot -Force
        $makeNsis = Get-ChildItem -LiteralPath $extractRoot -Filter 'makensis.exe' -File -Recurse | Select-Object -First 1
    }
    if ($null -eq $makeNsis) { throw 'NSIS_MAKENSIS_NOT_FOUND: the pinned NSIS archive has no makensis.exe.' }
    return $makeNsis.FullName
}

if (-not $SkipBuild) {
    $buildArguments = @('-SkipTests', '-MavenCommand', $MavenCommand)
    if (-not [string]::IsNullOrWhiteSpace($JavaHome)) { $buildArguments += @('-JavaHome', $JavaHome) }
    & (Join-Path $PSScriptRoot 'build.ps1') @buildArguments
    if (-not $?) { throw 'FRONTEND_BACKEND_BUILD_FAILED' }
}

$jar = Get-ChildItem -LiteralPath (Join-Path $backendRoot 'target') -Filter '*.jar' -File |
    Where-Object { $_.Name -notlike '*.original' } |
    Sort-Object Length -Descending |
    Select-Object -First 1
if ($null -eq $jar) { throw 'BACKEND_JAR_NOT_FOUND: run scripts/build.ps1 first.' }
$jarCommand = Get-Command jar.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -eq $jarCommand) { throw 'JAR_TOOL_NOT_FOUND: a JDK is required to inspect the packaged JAR.' }
$hasStaticIndex = & $jarCommand.Source tf $jar.FullName 2>$null | Select-String -SimpleMatch 'static/index.html'
if ($null -eq $hasStaticIndex) { throw "PACKAGED_WEB_MISSING: $($jar.FullName) does not contain static/index.html." }

if (Test-Path -LiteralPath $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $stageRoot 'app') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stageRoot 'runtime') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stageRoot 'service') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stageRoot 'config') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stageRoot 'tools') -Force | Out-Null
Copy-Item -LiteralPath $jar.FullName -Destination (Join-Path $stageRoot 'app\react-sheets.jar') -Force

$jreArchive = Get-VerifiedArtifact -Specification $dependencyManifest.javaRuntime -CacheDirectory $cacheRoot
$jreExtractRoot = Join-Path $cacheRoot ('jre-' + [string]$dependencyManifest.javaRuntime.version)
$jreJava = Get-ChildItem -LiteralPath $jreExtractRoot -Filter 'java.exe' -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Directory.Name -eq 'bin' } | Select-Object -First 1
if ($null -eq $jreJava) {
    if (Test-Path -LiteralPath $jreExtractRoot) { Remove-Item -LiteralPath $jreExtractRoot -Recurse -Force }
    New-Item -ItemType Directory -Path $jreExtractRoot -Force | Out-Null
    Expand-Archive -LiteralPath $jreArchive -DestinationPath $jreExtractRoot -Force
    $jreJava = Get-ChildItem -LiteralPath $jreExtractRoot -Filter 'java.exe' -File -Recurse | Where-Object { $_.Directory.Name -eq 'bin' } | Select-Object -First 1
}
if ($null -eq $jreJava) { throw 'JAVA_RUNTIME_NOT_FOUND: pinned JRE archive has no bin/java.exe.' }
$jreRoot = $jreJava.Directory.Parent.FullName
Copy-Item -Path (Join-Path $jreRoot '*') -Destination (Join-Path $stageRoot 'runtime') -Recurse -Force
$jreVersionText = & (Join-Path $stageRoot 'runtime\bin\java.exe') -version 2>&1 | Out-String
if ($LASTEXITCODE -ne 0 -or $jreVersionText -notmatch 'version\s+"21\.') { throw "JAVA_RUNTIME_VERSION_INVALID: $jreVersionText" }

$winSw = Get-VerifiedArtifact -Specification $dependencyManifest.winSw -CacheDirectory $cacheRoot
Copy-Item -LiteralPath $winSw -Destination (Join-Path $stageRoot 'service\ReactSheetsService.exe') -Force

$programData = [Environment]::GetFolderPath('CommonApplicationData')
$programDataForProperties = $programData.Replace('\', '/')
foreach ($templateName in @('application.properties', 'ReactSheetsService.xml')) {
    $source = Join-Path $installerRoot $templateName
    $destinationDirectory = if ($templateName.EndsWith('.xml')) { Join-Path $stageRoot 'service' } else { Join-Path $stageRoot 'config' }
    $destination = Join-Path $destinationDirectory $templateName
    $replacement = if ($templateName.EndsWith('.xml')) { $programData } else { $programDataForProperties }
    $content = (Get-Content -LiteralPath $source -Raw).Replace('@PROGRAMDATA@', $replacement)
    [IO.File]::WriteAllText($destination, $content, [Text.UTF8Encoding]::new($false))
}
foreach ($toolName in @('backup-data.ps1', 'restore-backup.ps1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $toolName) -Destination (Join-Path $stageRoot "tools\$toolName") -Force
}
Copy-Item -LiteralPath (Join-Path $installerRoot 'health-check.ps1') -Destination (Join-Path $stageRoot 'tools\health-check.ps1') -Force

$makeNsis = Find-MakeNsis -RequestedPath $NsisPath
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$nsisLog = Join-Path $LogRoot 'makensis.log'
$nsisArguments = @(
    '/NOCD',
    "/DAPP_VERSION=$version",
    "/DAPP_STAGE_DIR=$stageRoot",
    "/DOUT_DIR=$OutputDirectory",
    (Join-Path $installerRoot 'ReactSheets.nsi')
)
Push-Location $repositoryRoot
try {
    & $makeNsis @nsisArguments *> $nsisLog
    $nsisExitCode = $LASTEXITCODE
}
finally {
    Pop-Location
}
if ($nsisExitCode -ne 0) {
    Get-Content -LiteralPath $nsisLog -Tail 80 -ErrorAction SilentlyContinue
    throw "NSIS_BUILD_FAILED: exit code $nsisExitCode. Full log: $nsisLog"
}

$output = Join-Path $OutputDirectory "ReactSheets-Setup-$version.exe"
if (-not (Test-Path -LiteralPath $output -PathType Leaf)) { throw "INSTALLER_OUTPUT_MISSING: $output" }
$outputHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $output).Hash.ToLowerInvariant()
Write-Host "Installer created: $output"
Write-Host "SHA256: $outputHash"
Write-Host "Build logs: $LogRoot"
