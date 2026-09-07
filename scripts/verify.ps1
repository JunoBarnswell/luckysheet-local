$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = Split-Path -Parent $PSScriptRoot
$Frontend = Join-Path $Root 'frontend-react'
$ToolRoot = Join-Path $Root '.tools'
$LocalJdk = Get-ChildItem (Join-Path $ToolRoot 'jdk-21') -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
$JavaCommand = if ($LocalJdk) { Join-Path $LocalJdk.FullName 'bin\java.exe' } else { 'java' }
$MavenCommand = if (Test-Path (Join-Path $ToolRoot 'maven\apache-maven-3.9.9\bin\mvn.cmd')) { Join-Path $ToolRoot 'maven\apache-maven-3.9.9\bin\mvn.cmd' } else { 'mvn' }
if ($LocalJdk) { $env:JAVA_HOME = $LocalJdk.FullName; $env:Path = "$($LocalJdk.FullName)\bin;$env:Path" }
function Invoke-Checked([string]$Command, [string[]]$Arguments, [string]$WorkingDirectory = $Root) { Push-Location $WorkingDirectory; try { & $Command @Arguments; if ($LASTEXITCODE -ne 0) { throw "Command failed ($LASTEXITCODE): $Command" } } finally { Pop-Location } }
Write-Host '==> Toolchain report'; node --version; rustc --version; & $JavaCommand --version 2>&1 | Select-Object -First 1; & $MavenCommand --version | Select-Object -First 1
Write-Host '==> One canonical build'; & (Join-Path $PSScriptRoot 'build.ps1')
Write-Host '==> Kernel artifact integrity'; Invoke-Checked 'node' @((Join-Path $PSScriptRoot 'verify-kernel-manifest.mjs'))
Write-Host '==> Frontend typecheck, contracts, boundaries, and unit suite'; Invoke-Checked 'npm' @('run', 'typecheck') $Frontend; Invoke-Checked 'npm' @('run', 'check:boundaries') $Frontend; Invoke-Checked 'npm' @('run', 'test:unit') $Frontend
Write-Host '==> Backend Maven verification'; Invoke-Checked $MavenCommand @('test') (Join-Path $Root 'backend')
Write-Host 'Verification completed. Native Excel corpus and browser acceptance require their explicitly configured environments.'
