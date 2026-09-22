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
    Expand-Archive -LiteralPath $BackupPath -DestinationPath $stagingRoot -Force
    $manifestPath = Join-Path $stagingRoot 'manifest.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw 'BACKUP_MANIFEST_MISSING' }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if ($manifest.format -ne 'react-sheets-backup' -or [int]$manifest.version -ne 1) {
        throw 'BACKUP_MANIFEST_UNSUPPORTED'
    }
    $manifestFiles = @($manifest.files)
    if ($manifestFiles.Count -eq 0) { throw 'BACKUP_MANIFEST_EMPTY' }

    foreach ($entry in $manifestFiles) {
        $relativePath = [string]$entry.path
        if ([string]::IsNullOrWhiteSpace($relativePath) -or [IO.Path]::IsPathRooted($relativePath)) {
            throw "BACKUP_MANIFEST_PATH_INVALID: $relativePath"
        }
        $filePath = [IO.Path]::GetFullPath((Join-Path $stagingRoot $relativePath))
        $stagingPrefix = "$([IO.Path]::GetFullPath($stagingRoot))$([IO.Path]::DirectorySeparatorChar)"
        if (-not $filePath.StartsWith($stagingPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw "BACKUP_MANIFEST_PATH_INVALID: $relativePath" }
        if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw "BACKUP_FILE_MISSING: $relativePath" }
        $actual = Get-FileHash -Algorithm SHA256 -LiteralPath $filePath
        if ($actual.Hash.ToLowerInvariant() -ne ([string]$entry.sha256).ToLowerInvariant()) { throw "BACKUP_FILE_CHECKSUM_MISMATCH: $relativePath" }
        if ((Get-Item -LiteralPath $filePath).Length -ne [int64]$entry.length) { throw "BACKUP_FILE_LENGTH_MISMATCH: $relativePath" }
    }
    if (-not ($manifestFiles.path -match '\.mv\.db$')) { throw 'BACKUP_DATABASE_MISSING' }

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
