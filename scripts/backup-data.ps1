[CmdletBinding()]
param(
    [string]$DataRoot = (Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'ReactSheets\data'),
    [string]$OutputDirectory = (Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'ReactSheets\backups'),
    [string]$ServiceName = 'ReactSheets',
    [string]$OutputPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-FullPath([string]$PathValue) {
    return [IO.Path]::GetFullPath($PathValue)
}

$DataRoot = Resolve-FullPath $DataRoot
$OutputDirectory = Resolve-FullPath $OutputDirectory
if (-not (Test-Path -LiteralPath $DataRoot -PathType Container)) {
    throw "BACKUP_SOURCE_NOT_FOUND: $DataRoot"
}

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($null -ne $service -and $service.Status -ne 'Stopped') {
    throw "SERVICE_MUST_BE_STOPPED: $ServiceName is $($service.Status). Stop the service before copying H2 files."
}

$sourceFiles = @(Get-ChildItem -LiteralPath $DataRoot -File -Recurse | Sort-Object FullName)
if ($sourceFiles.Count -eq 0) {
    throw "BACKUP_SOURCE_EMPTY: $DataRoot"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $OutputPath = Join-Path $OutputDirectory ('react-sheets-backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.zip')
} else {
    $OutputPath = Resolve-FullPath $OutputPath
    $OutputDirectory = Split-Path -Parent $OutputPath
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

if ((Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    throw "BACKUP_DESTINATION_EXISTS: $OutputPath. Pass -Force to replace it."
}

$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) ('react-sheets-backup-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null
try {
    $manifestFiles = [System.Collections.Generic.List[object]]::new()
    foreach ($sourceFile in $sourceFiles) {
        $relativePath = [IO.Path]::GetRelativePath($DataRoot, $sourceFile.FullName)
        $targetPath = Join-Path $stagingRoot $relativePath
        New-Item -ItemType Directory -Path (Split-Path -Parent $targetPath) -Force | Out-Null
        Copy-Item -LiteralPath $sourceFile.FullName -Destination $targetPath -Force
        $manifestFiles.Add([ordered]@{
                path = $relativePath.Replace('\', '/')
                length = $sourceFile.Length
                sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $targetPath).Hash.ToLowerInvariant()
            })
    }

    $manifest = [ordered]@{
        format = 'react-sheets-backup'
        version = 1
        createdAt = [DateTime]::UtcNow.ToString('o')
        databaseName = 'luckysheet_canonical'
        files = $manifestFiles
    }
    $manifestJson = $manifest | ConvertTo-Json -Depth 8
    Set-Content -LiteralPath (Join-Path $stagingRoot 'manifest.json') -Value $manifestJson -Encoding utf8NoBOM

    Compress-Archive -Path (Join-Path $stagingRoot '*') -DestinationPath $OutputPath -CompressionLevel Optimal -Force
    $archiveHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $OutputPath).Hash.ToLowerInvariant()
    Write-Host "Backup created: $OutputPath"
    Write-Host "SHA256: $archiveHash"
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}
