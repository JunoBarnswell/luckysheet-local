$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$Root = Split-Path -Parent $PSScriptRoot
$Frontend = Join-Path $Root 'frontend-react'
$Cargo = Join-Path $Root 'Cargo.toml'
$Target = Join-Path $Root 'target'
$WebKernel = Join-Path $Frontend 'apps\web\public\kernel'
$ToolRoot = Join-Path $Root '.tools'
$LocalJdk = Get-ChildItem (Join-Path $ToolRoot 'jdk-21') -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
$JavaCommand = if ($LocalJdk) { Join-Path $LocalJdk.FullName 'bin\java.exe' } else { 'java' }
$MavenCommand = if (Test-Path (Join-Path $ToolRoot 'maven\apache-maven-3.9.9\bin\mvn.cmd')) { Join-Path $ToolRoot 'maven\apache-maven-3.9.9\bin\mvn.cmd' } else { 'mvn' }
if ($LocalJdk) { $env:JAVA_HOME = $LocalJdk.FullName; $env:Path = "$($LocalJdk.FullName)\bin;$env:Path" }

function Invoke-Checked([string]$Command, [string[]]$Arguments, [string]$WorkingDirectory = $Root) {
    Write-Host ("==> {0} {1}" -f $Command, ($Arguments -join ' '))
    Push-Location $WorkingDirectory
    try { & $Command @Arguments; if ($LASTEXITCODE -ne 0) { throw "Command failed ($LASTEXITCODE): $Command" } }
    finally { Pop-Location }
}

function Require-Version([string]$Command, [string]$ExpectedPattern, [string]$DisplayName) {
    $actual = (& $Command '--version' 2>&1 | Out-String).Trim()
    if ($actual -notmatch $ExpectedPattern) { throw "$DisplayName must match '$ExpectedPattern'; found '$actual'" }
    Write-Host "    $DisplayName $actual"
}

Require-Version 'node' '^v24\.18\.0' 'Node'
Require-Version 'rustc' '^rustc 1\.97\.1' 'Rust'
Require-Version $JavaCommand '^(openjdk|java) 21(?:\.|\s)' 'Java'
Require-Version $MavenCommand '^Apache Maven 3\.9\.9' 'Maven'
Invoke-Checked 'cargo' @('build', '--manifest-path', $Cargo, '-p', 'kernel-host', '--release')
Invoke-Checked 'cargo' @('build', '--manifest-path', $Cargo, '-p', 'kernel-host', '--target', 'wasm32-unknown-unknown', '--release')

$nativeName = if ($env:OS -eq 'Windows_NT') { 'workbook-kernel-host.exe' } else { 'workbook-kernel-host' }
$native = Join-Path $Target (Join-Path 'release' $nativeName)
$wasm = Join-Path $Target 'wasm32-unknown-unknown\release\kernel_host.wasm'
if (-not (Test-Path $native)) { throw "Rust native artifact not found: $native" }
if (-not (Test-Path $wasm)) { throw "Rust WASM artifact not found: $wasm" }
New-Item -ItemType Directory -Path $WebKernel -Force | Out-Null
Copy-Item -LiteralPath $wasm -Destination (Join-Path $WebKernel 'kernel_host.wasm') -Force
Invoke-Checked 'node' @((Join-Path $PSScriptRoot 'write-kernel-manifest.mjs'), $WebKernel)
Invoke-Checked 'npm' @('ci', '--ignore-scripts', '--no-audit', '--no-fund') $Frontend
Invoke-Checked 'npm' @('run', 'build') $Frontend
Write-Host "==> Build completed`n    native: $native`n    wasm:   $WebKernel\kernel_host.wasm"
