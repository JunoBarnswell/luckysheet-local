[CmdletBinding()]
param(
    [string]$JavaHome,
    [string]$MavenCommand = 'mvn.cmd',
    [string]$LogRoot,
    [switch]$SkipTests,
    [switch]$SkipFrontend,
    [switch]$SkipBackend
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$frontendRoot = Join-Path $repositoryRoot 'frontend-react'
$backendRoot = Join-Path $repositoryRoot 'backend'
$webDistribution = Join-Path $frontendRoot 'dist\web'
$generatedWeb = Join-Path $backendRoot 'target\generated-web'

if ([string]::IsNullOrWhiteSpace($LogRoot)) {
    $LogRoot = Join-Path ([IO.Path]::GetTempPath()) ('react-sheets-build\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

function Resolve-CommandPath {
    param([Parameter(Mandatory)][string]$Name)

    $command = Get-Command $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $command) {
        throw "Required command '$Name' was not found on PATH."
    }
    return $command.Source
}

function Invoke-LoggedCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$Arguments,
        [Parameter(Mandatory)][string]$WorkingDirectory
    )

    $safeName = ($Name -replace '[^A-Za-z0-9_.-]', '-')
    $logPath = Join-Path $LogRoot ($safeName + '.log')
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
        Write-Host "    exit: $exitCode"
        Get-Content -LiteralPath $logPath -Tail 60 -ErrorAction SilentlyContinue
        throw "$Name failed with exit code $exitCode. Full log: $logPath"
    }

    Write-Host '    exit: 0'
}

function Get-VersionOutput {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    $output = & $FilePath @Arguments 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to query '$FilePath'."
    }
    return $output.Trim()
}

function Select-JavaHome {
    param([string]$RequestedHome)

    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($RequestedHome)) { $candidates.Add($RequestedHome) }
    if (-not [string]::IsNullOrWhiteSpace($env:JAVA_HOME)) { $candidates.Add($env:JAVA_HOME) }

    $pathJava = Get-Command java.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $pathJava) {
        $javaBin = Split-Path -Parent $pathJava.Source
        $candidates.Add((Split-Path -Parent $javaBin))
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        $jdkRoot = [IO.Path]::GetFullPath($candidate)
        $javaExe = Join-Path $jdkRoot 'bin\java.exe'
        $javacExe = Join-Path $jdkRoot 'bin\javac.exe'
        if (-not (Test-Path -LiteralPath $javaExe) -or -not (Test-Path -LiteralPath $javacExe)) { continue }

        $versionText = Get-VersionOutput -FilePath $javaExe -Arguments @('-version')
        $match = [regex]::Match($versionText, 'version\s+"(?<version>[^"\s]+)')
        if (-not $match.Success) { continue }
        $major = [int]($match.Groups['version'].Value.Split('.')[0])
        $javacVersion = Get-VersionOutput -FilePath $javacExe -Arguments @('-version')
        if ($major -eq 21 -and $javacVersion -match '\b21\.') {
            return $jdkRoot
        }
    }

    $observed = if ($pathJava) { Get-VersionOutput -FilePath $pathJava.Source -Arguments @('-version') } else { 'java-not-found' }
    throw "JDK 21 is required. Pass -JavaHome <JDK21 directory> or set JAVA_HOME. Observed: $observed"
}

if (-not (Test-Path -LiteralPath $frontendRoot -PathType Container)) { throw "Frontend root not found: $frontendRoot" }
if (-not (Test-Path -LiteralPath $backendRoot -PathType Container)) { throw "Backend root not found: $backendRoot" }

$nodeCommand = Resolve-CommandPath -Name 'node.exe'
$npmCommand = Resolve-CommandPath -Name 'npm.cmd'
$nodeVersion = Get-VersionOutput -FilePath $nodeCommand -Arguments @('--version')
$nodeMatch = [regex]::Match($nodeVersion, '^v(?<major>\d+)')
if (-not $nodeMatch.Success -or [int]$nodeMatch.Groups['major'].Value -ne 24) {
    throw "Node 24 is required. Observed: $nodeVersion"
}
Write-Host "Node: $nodeVersion"
Write-Host "npm:  $(Get-VersionOutput -FilePath $npmCommand -Arguments @('--version'))"

$selectedJavaHome = Select-JavaHome -RequestedHome $JavaHome
$env:JAVA_HOME = $selectedJavaHome
$env:Path = (Join-Path $selectedJavaHome 'bin') + [IO.Path]::PathSeparator + $env:Path
$javaCommand = Join-Path $selectedJavaHome 'bin\java.exe'
Write-Host "JDK:  $selectedJavaHome"
Write-Host "Java: $(Get-VersionOutput -FilePath $javaCommand -Arguments @('-version') | Select-Object -First 1)"

$mavenPath = if (Test-Path -LiteralPath $MavenCommand -PathType Leaf) {
    [IO.Path]::GetFullPath($MavenCommand)
} else {
    Resolve-CommandPath -Name $MavenCommand
}

if (-not $SkipFrontend) {
    if (-not (Test-Path -LiteralPath (Join-Path $frontendRoot 'package-lock.json') -PathType Leaf)) {
        throw "frontend-react/package-lock.json is required for npm ci."
    }
    Invoke-LoggedCommand -Name 'frontend-npm-ci' -FilePath $npmCommand -Arguments @('ci', '--no-audit', '--no-fund') -WorkingDirectory $frontendRoot
    Invoke-LoggedCommand -Name 'frontend-build' -FilePath $npmCommand -Arguments @('run', 'build') -WorkingDirectory $frontendRoot
}

if (-not (Test-Path -LiteralPath (Join-Path $webDistribution 'index.html') -PathType Leaf)) {
    throw "Frontend build output is missing: $webDistribution\index.html"
}

if (-not $SkipBackend) {
    Invoke-LoggedCommand -Name 'backend-clean' -FilePath $mavenPath -Arguments @('--batch-mode', '--no-transfer-progress', 'clean') -WorkingDirectory $backendRoot

    if (Test-Path -LiteralPath $generatedWeb) {
        Remove-Item -LiteralPath $generatedWeb -Recurse -Force
    }
    New-Item -ItemType Directory -Path $generatedWeb -Force | Out-Null
    Copy-Item -Path (Join-Path $webDistribution '*') -Destination $generatedWeb -Recurse -Force

    $packageArguments = [System.Collections.Generic.List[string]]::new()
    $packageArguments.Add('--batch-mode')
    $packageArguments.Add('--no-transfer-progress')
    $packageArguments.Add('package')
    if ($SkipTests) { $packageArguments.Add('-DskipTests') }
    Invoke-LoggedCommand -Name 'backend-package' -FilePath $mavenPath -Arguments $packageArguments.ToArray() -WorkingDirectory $backendRoot

    $jar = Get-ChildItem -LiteralPath (Join-Path $backendRoot 'target') -Filter '*.jar' -File |
        Where-Object { $_.Name -notlike '*.original' } |
        Sort-Object Length -Descending |
        Select-Object -First 1
    if ($null -eq $jar) { throw "Maven package completed without a runnable JAR in backend/target." }
    $jarCommand = Join-Path $selectedJavaHome 'bin\jar.exe'
    $hasStaticIndex = & $jarCommand tf $jar.FullName 2>$null | Select-String -SimpleMatch 'static/index.html'
    if ($null -eq $hasStaticIndex) { throw "Packaged JAR does not contain static/index.html: $($jar.FullName)" }
    Write-Host "Backend JAR: $($jar.FullName)"
}

Write-Host "Build completed. Logs: $LogRoot"
