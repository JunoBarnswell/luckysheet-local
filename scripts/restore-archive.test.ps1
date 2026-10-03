$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'restore-archive.ps1')
Add-Type -AssemblyName System.IO.Compression
$root = Join-Path ([IO.Path]::GetTempPath()) ('restore-budget-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
function New-TestBackup($Name, $Extra = @{}, $DeclaredLength = 8, $Hash = $null) {
    $path = Join-Path $root "$Name.zip"
    $bytes = [Text.Encoding]::UTF8.GetBytes('database')
    if ([string]::IsNullOrWhiteSpace([string]$Hash)) { $Hash = [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($bytes)).Replace('-', '').ToLowerInvariant() }
    $manifest = @{ format = 'react-sheets-backup'; version = 1; files = @(@{path = 'db.mv.db'; length = $DeclaredLength; sha256 = $Hash}) } | ConvertTo-Json -Depth 8
    $stream = [IO.File]::Create($path); $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entries = @{'manifest.json' = [Text.Encoding]::UTF8.GetBytes($manifest); 'db.mv.db' = $bytes}
        
        foreach ($key in $entries.Keys) { $entry = $archive.CreateEntry($key); $output = $entry.Open(); try { $output.Write($entries[$key], 0, $entries[$key].Length) } finally { $output.Dispose() } }
        foreach ($key in $Extra.Keys) { $entry = $archive.CreateEntry($key); $output = $entry.Open(); try { $output.Write($Extra[$key], 0, $Extra[$key].Length) } finally { $output.Dispose() } }
    } finally { $archive.Dispose(); $stream.Dispose() }
    return $path
}
function Set-FalselySmallZipCount($Path, [switch]$AlsoShrinkDirectory) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    for ($index = $bytes.Length - 22; $index -ge [Math]::Max(0, $bytes.Length - 65557); $index--) {
        if ($bytes[$index] -eq 0x50 -and $bytes[$index + 1] -eq 0x4b -and $bytes[$index + 2] -eq 0x05 -and $bytes[$index + 3] -eq 0x06) {
            [BitConverter]::GetBytes([uint16]1).CopyTo($bytes, $index + 8)
            [BitConverter]::GetBytes([uint16]1).CopyTo($bytes, $index + 10)
            if ($AlsoShrinkDirectory) {
                $directoryOffset = [BitConverter]::ToUInt32($bytes, $index + 16)
                $firstHeaderLength = 46 + [BitConverter]::ToUInt16($bytes, $directoryOffset + 28) + [BitConverter]::ToUInt16($bytes, $directoryOffset + 30) + [BitConverter]::ToUInt16($bytes, $directoryOffset + 32)
                [BitConverter]::GetBytes([uint32]$firstHeaderLength).CopyTo($bytes, $index + 12)
            }
            [IO.File]::WriteAllBytes($Path, $bytes)
            return
        }
    }
    throw 'ZIP end directory not found in test fixture'
}
function Assert-Rejected($Path, $Code, $Options = @{}) {
    $staging = Join-Path $root ([guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $staging | Out-Null
    $caught = $false
    try { Expand-VerifiedBackup -BackupPath $Path -StagingRoot $staging @Options | Out-Null }
    catch { if ($_.Exception.Message -notlike "*$Code*") { throw }; $caught = $true }
    if (-not $caught) { throw "Expected rejection: $Code" }
    if (@(Get-ChildItem -LiteralPath $staging -File -Recurse).Count -ne 0) { throw 'Unexpected staging growth before admission' }
}
try {
    $valid = New-TestBackup 'valid'; $staging = Join-Path $root 'valid-stage'; New-Item -ItemType Directory -Path $staging | Out-Null
    $files = @(Expand-VerifiedBackup -BackupPath $valid -StagingRoot $staging -MaxEntryBytes 1000 -MaxExpandedBytes 1000)
    if ($files.Count -ne 1 -or [IO.File]::ReadAllText((Join-Path $staging 'db.mv.db')) -ne 'database') { throw 'Valid backup did not restore' }
    Assert-Rejected $valid 'BACKUP_ENTRY_COUNT_EXCEEDED' @{MaxEntries=1}
    $spoofedCount = New-TestBackup 'spoofed-count'
    Set-FalselySmallZipCount $spoofedCount
    Assert-Rejected $spoofedCount 'BACKUP_ENTRY_COUNT_EXCEEDED' @{MaxEntries=1}
    $spoofedCountAndSize = New-TestBackup 'spoofed-count-and-size'
    Set-FalselySmallZipCount $spoofedCountAndSize -AlsoShrinkDirectory
    Assert-Rejected $spoofedCountAndSize 'BACKUP_DIRECTORY_INVALID' @{MaxEntries=10}
    Assert-Rejected $valid 'BACKUP_EXTRACTION_BUDGET_EXCEEDED' @{MaxEntryBytes=7}
    Assert-Rejected $valid 'BACKUP_EXTRACTION_BUDGET_EXCEEDED' @{MaxExpandedBytes=10}
    Assert-Rejected (New-TestBackup 'bomb' @{'huge.bin'=[byte[]]::new(100000)}) 'BACKUP_EXTRACTION_BUDGET_EXCEEDED' @{MaxCompressionRatio=10}
    Assert-Rejected (New-TestBackup 'traversal' @{'../escape'=[byte[]]@(1)}) 'BACKUP_PATH_INVALID'
    Assert-Rejected (New-TestBackup 'duplicate' @{'DB.MV.DB'=[byte[]]@(1)}) 'BACKUP_DUPLICATE_PATH'
    Assert-Rejected (New-TestBackup 'undeclared' @{'extra.bin'=[byte[]]@(1)}) 'BACKUP_UNDECLARED_ENTRY'
    Assert-Rejected (New-TestBackup 'length' @{} 7) 'BACKUP_MANIFEST_METADATA_INVALID'
    $wrongHash = New-TestBackup 'hash' @{} 8 ('0' * 64)
    $hashStage = Join-Path $root 'hash-stage'; New-Item -ItemType Directory -Path $hashStage | Out-Null
    $hashRejected = $false
    try { Expand-VerifiedBackup -BackupPath $wrongHash -StagingRoot $hashStage | Out-Null }
    catch { if ($_.Exception.Message -notlike '*BACKUP_FILE_CHECKSUM_MISMATCH*') { throw }; $hashRejected = $true }
    if (-not $hashRejected) { throw 'Expected checksum rejection' }
    Write-Host 'PASS: valid restore, byte/count/ratio budgets, traversal, duplicates, undeclared entries, length and hash checks'
} finally { Remove-Item -LiteralPath $root -Recurse -Force }
