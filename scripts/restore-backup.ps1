[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$BackupPath,
    [string]$DataRoot = (Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'ReactSheets\data'),
    [string]$ServiceName = 'ReactSheets',
    [string]$ExpectedSha256,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $Force) {
    throw 'RESTORE_CONFIRMATION_REQUIRED: pass -Force only after verifying the backup and stopping the service.'
}

$BackupPath = [IO.Path]::GetFullPath($BackupPath)
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) { throw "BACKUP_NOT_FOUND: $BackupPath" }
if ([IO.Path]::GetExtension($BackupPath).ToLowerInvariant() -ne '.zip') { throw "BACKUP_FORMAT_UNSUPPORTED: $BackupPath" }

$archiveHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $BackupPath).Hash.ToLowerInvariant()
if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256) -and $archiveHash -ne $ExpectedSha256.ToLowerInvariant()) {
    throw "BACKUP_CHECKSUM_MISMATCH: expected $ExpectedSha256, actual $archiveHash"
}

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($null -ne $service -and $service.Status -ne 'Stopped') {
    throw "SERVICE_MUST_BE_STOPPED: $ServiceName is $($service.Status)."
}

$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) ('react-sheets-restore-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null
$previousRoot = "$DataRoot.before-restore-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
$movedPrevious = $false
try {
    . (Join-Path $PSScriptRoot 'restore-archive.ps1')
    $manifestFiles = @(Expand-VerifiedBackup -BackupPath $BackupPath -StagingRoot $stagingRoot)
    $restoreBytes = [long]0
    foreach ($entry in $manifestFiles) { $restoreBytes += [long]$entry.length }
    $targetDrive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($DataRoot))
    if ($targetDrive.AvailableFreeSpace -lt $restoreBytes + 64MB) { throw 'BACKUP_TARGET_SPACE_INSUFFICIENT' }

    if (Test-Path -LiteralPath $DataRoot) {
        if (-not $PSCmdlet.ShouldProcess($DataRoot, "Move existing data to $previousRoot")) { return }
        Move-Item -LiteralPath $DataRoot -Destination $previousRoot
        $movedPrevious = $true
    }
    New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null
    foreach ($entry in $manifestFiles) {
        $relativePath = [string]$entry.path
        $sourceFile = Join-Path $stagingRoot $relativePath
        $targetFile = Join-Path $DataRoot $relativePath
        New-Item -ItemType Directory -Path (Split-Path -Parent $targetFile) -Force | Out-Null
        Copy-Item -LiteralPath $sourceFile -Destination $targetFile -Force
    }
    Write-Host "Restore completed: $DataRoot"
    Write-Host "Backup SHA256: $archiveHash"
    if ($movedPrevious) { Write-Host "Previous data retained: $previousRoot" }
}
catch {
    if ($movedPrevious -and (Test-Path -LiteralPath $previousRoot)) {
        if (Test-Path -LiteralPath $DataRoot) {
            Remove-Item -LiteralPath $DataRoot -Recurse -Force
        }
        Move-Item -LiteralPath $previousRoot -Destination $DataRoot
    }
    throw
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) { Remove-Item -LiteralPath $stagingRoot -Recurse -Force }
}
