Set-StrictMode -Version Latest

# Admission and extraction are one boundary: no archive-controlled path is
# written until the directory and bounded manifest have both been validated.
function Expand-VerifiedBackup {
    param(
        [Parameter(Mandatory)][string]$BackupPath,
        [Parameter(Mandatory)][string]$StagingRoot,
        [long]$MaxArchiveBytes = 2GB,
        [long]$MaxExpandedBytes = 4GB,
        [long]$MaxEntryBytes = 2GB,
        [int]$MaxEntries = 10000,
        [int]$MaxCompressionRatio = 1000,
        [int]$MaxManifestBytes = 1MB
    )
    foreach ($limit in @($MaxArchiveBytes, $MaxExpandedBytes, $MaxEntryBytes, $MaxEntries, $MaxCompressionRatio, $MaxManifestBytes)) {
        if ($limit -le 0) { throw 'BACKUP_BUDGET_INVALID' }
    }
    if ((Get-Item -LiteralPath $BackupPath).Length -gt $MaxArchiveBytes) { throw 'BACKUP_ARCHIVE_BUDGET_EXCEEDED' }
    Add-Type -AssemblyName System.IO.Compression
    $archiveStream = [IO.File]::OpenRead($BackupPath)
    $archive = $null
    try {
        # Read only the fixed-size end records first. ZipArchive.Entries parses
        # and allocates the complete central directory, so it cannot be the
        # admission mechanism for an attacker-controlled entry count.
        $tailLength = [int][Math]::Min($archiveStream.Length, 65557)
        if ($tailLength -lt 22) { throw 'BACKUP_DIRECTORY_INVALID' }
        $tail = [byte[]]::new($tailLength)
        [void]$archiveStream.Seek(-$tailLength, [IO.SeekOrigin]::End)
        if ($archiveStream.Read($tail, 0, $tail.Length) -ne $tail.Length) { throw 'BACKUP_DIRECTORY_INVALID' }
        $eocd = -1
        for ($index = $tail.Length - 22; $index -ge [Math]::Max(0, $tail.Length - 65557); $index--) {
            if ($tail[$index] -eq 0x50 -and $tail[$index + 1] -eq 0x4b -and $tail[$index + 2] -eq 0x05 -and $tail[$index + 3] -eq 0x06) { $eocd = $index; break }
        }
        if ($eocd -lt 0) { throw 'BACKUP_DIRECTORY_INVALID' }
        if ([BitConverter]::ToUInt16($tail, $eocd + 4) -ne 0 -or [BitConverter]::ToUInt16($tail, $eocd + 6) -ne 0) { throw 'BACKUP_DIRECTORY_INVALID' }
        $entryCount = [long][BitConverter]::ToUInt16($tail, $eocd + 10)
        $directoryBytes = [long][BitConverter]::ToUInt32($tail, $eocd + 12)
        $directoryOffset = [long][BitConverter]::ToUInt32($tail, $eocd + 16)
        $authoritativeDirectoryEnd = [long]($archiveStream.Length - $tail.Length + $eocd)
        if ($entryCount -eq 0xFFFF) {
            $locator = $eocd - 20
            if ($locator -lt 0 -or $tail[$locator] -ne 0x50 -or $tail[$locator + 1] -ne 0x4b -or $tail[$locator + 2] -ne 0x06 -or $tail[$locator + 3] -ne 0x07) { throw 'BACKUP_DIRECTORY_INVALID' }
            $zip64Offset = [BitConverter]::ToInt64($tail, $locator + 8)
            if ($zip64Offset -lt 0 -or $zip64Offset + 56 -gt $archiveStream.Length) { throw 'BACKUP_DIRECTORY_INVALID' }
            [void]$archiveStream.Seek($zip64Offset, [IO.SeekOrigin]::Begin)
            $zip64 = [byte[]]::new(56)
            if ($archiveStream.Read($zip64, 0, $zip64.Length) -ne $zip64.Length -or $zip64[0] -ne 0x50 -or $zip64[1] -ne 0x4b -or $zip64[2] -ne 0x06 -or $zip64[3] -ne 0x06) { throw 'BACKUP_DIRECTORY_INVALID' }
            $entryCount = [BitConverter]::ToInt64($zip64, 32)
            $directoryBytes = [BitConverter]::ToInt64($zip64, 40)
            $directoryOffset = [BitConverter]::ToInt64($zip64, 48)
            $authoritativeDirectoryEnd = $zip64Offset
        }
        if ($entryCount -lt 0 -or $entryCount -gt $MaxEntries) { throw 'BACKUP_ENTRY_COUNT_EXCEEDED' }
        if ($directoryBytes -lt 0 -or $directoryOffset -lt 0 -or $directoryOffset + $directoryBytes -ne $authoritativeDirectoryEnd) { throw 'BACKUP_DIRECTORY_INVALID' }
        [void]$archiveStream.Seek($directoryOffset, [IO.SeekOrigin]::Begin)
        $directoryEnd = $directoryOffset + $directoryBytes
        $scannedEntries = [long]0
        $fixedHeader = [byte[]]::new(46)
        while ($archiveStream.Position -lt $directoryEnd) {
            if ($archiveStream.Position + $fixedHeader.Length -gt $directoryEnd -or $archiveStream.Read($fixedHeader, 0, $fixedHeader.Length) -ne $fixedHeader.Length) { throw 'BACKUP_DIRECTORY_INVALID' }
            if ($fixedHeader[0] -ne 0x50 -or $fixedHeader[1] -ne 0x4b -or $fixedHeader[2] -ne 0x01 -or $fixedHeader[3] -ne 0x02) { throw 'BACKUP_DIRECTORY_INVALID' }
            $nameLength = [long][BitConverter]::ToUInt16($fixedHeader, 28)
            $extraLength = [long][BitConverter]::ToUInt16($fixedHeader, 30)
            $commentLength = [long][BitConverter]::ToUInt16($fixedHeader, 32)
            $nextHeader = $archiveStream.Position + $nameLength + $extraLength + $commentLength
            if ($nextHeader -gt $directoryEnd) { throw 'BACKUP_DIRECTORY_INVALID' }
            $scannedEntries++
            if ($scannedEntries -gt $MaxEntries) { throw 'BACKUP_ENTRY_COUNT_EXCEEDED' }
            [void]$archiveStream.Seek($nextHeader, [IO.SeekOrigin]::Begin)
        }
        if ($scannedEntries -ne $entryCount -or $archiveStream.Position -ne $directoryEnd) { throw 'BACKUP_DIRECTORY_INVALID' }
        [void]$archiveStream.Seek(0, [IO.SeekOrigin]::Begin)
        $archive = [IO.Compression.ZipArchive]::new($archiveStream, [IO.Compression.ZipArchiveMode]::Read, $false)
        if ($archive.Entries.Count -ne $entryCount) { throw 'BACKUP_DIRECTORY_INVALID' }
        $entries = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
        $total = [long]0
        foreach ($entry in $archive.Entries) {
            $path = $entry.FullName
            if ([string]::IsNullOrWhiteSpace($path) -or $path.Contains('\') -or $path.Contains(':') -or $path.Contains([char]0) -or $path.StartsWith('/')) { throw "BACKUP_PATH_INVALID: $path" }
            $components = $path.TrimEnd('/').Split('/')
            foreach ($component in $components) {
                if ($component -in @('', '.', '..') -or $component.EndsWith('.') -or $component.EndsWith(' ') -or $component -match '[<>"|?*\x00-\x1f]' -or $component -match '^(?i:CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\.|$)') { throw "BACKUP_PATH_INVALID: $path" }
            }
            $key = $path.TrimEnd('/')
            if ($entries.ContainsKey($key)) { throw "BACKUP_DUPLICATE_PATH: $path" }
            if ((($entry.ExternalAttributes -shr 16) -band 0xF000) -eq 0xA000) { throw "BACKUP_LINK_UNSUPPORTED: $path" }
            if ($entry.Length -gt $MaxEntryBytes -or $entry.Length -gt $MaxExpandedBytes - $total -or ($entry.Length -gt 0 -and ($entry.CompressedLength -eq 0 -or $entry.Length / $entry.CompressedLength -gt $MaxCompressionRatio))) { throw "BACKUP_EXTRACTION_BUDGET_EXCEEDED: $path" }
            if ($path.EndsWith('/') -and $entry.Length -ne 0) { throw "BACKUP_DIRECTORY_INVALID: $path" }
            $total += $entry.Length
            $entries.Add($key, $entry)
        }
        foreach ($path in $entries.Keys) {
            $parent = $path
            while ($parent.Contains('/')) {
                $parent = $parent.Substring(0, $parent.LastIndexOf('/'))
                if ($entries.ContainsKey($parent) -and -not $entries[$parent].FullName.EndsWith('/')) { throw "BACKUP_PATH_COLLISION: $path" }
            }
        }
        if (-not $entries.ContainsKey('manifest.json')) { throw 'BACKUP_MANIFEST_MISSING' }
        $manifestEntry = $entries['manifest.json']
        if ($manifestEntry.Length -gt $MaxManifestBytes) { throw 'BACKUP_MANIFEST_BUDGET_EXCEEDED' }
        $manifestStream = $manifestEntry.Open()
        $manifestBuffer = [IO.MemoryStream]::new()
        try {
            $buffer = [byte[]]::new(8192)
            while (($read = $manifestStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                if ($manifestBuffer.Length + $read -gt $MaxManifestBytes -or $manifestBuffer.Length + $read -gt $manifestEntry.Length) { throw 'BACKUP_MANIFEST_BUDGET_EXCEEDED' }
                $manifestBuffer.Write($buffer, 0, $read)
            }
            if ($manifestBuffer.Length -ne $manifestEntry.Length) { throw 'BACKUP_MANIFEST_LENGTH_MISMATCH' }
            $manifest = [Text.UTF8Encoding]::new($false, $true).GetString($manifestBuffer.ToArray()).TrimStart([char]0xFEFF) | ConvertFrom-Json
        } finally { $manifestStream.Dispose(); $manifestBuffer.Dispose() }
        if ($manifest.format -ne 'react-sheets-backup' -or $manifest.version -ne 1) { throw 'BACKUP_MANIFEST_UNSUPPORTED' }
        $files = @($manifest.files)
        if ($files.Count -eq 0 -or $files.Count -gt $MaxEntries - 1) { throw 'BACKUP_MANIFEST_FILE_COUNT_INVALID' }
        $selected = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($file in $files) {
            $path = [string]$file.path
            if ($path -eq 'manifest.json' -or -not $entries.ContainsKey($path) -or $entries[$path].FullName.EndsWith('/') -or -not $selected.Add($path)) { throw "BACKUP_MANIFEST_PATH_INVALID: $path" }
            if ($file.length -isnot [long] -and $file.length -isnot [int]) { throw "BACKUP_MANIFEST_LENGTH_INVALID: $path" }
            if ([long]$file.length -lt 0 -or [long]$file.length -ne $entries[$path].Length -or [string]$file.sha256 -notmatch '^[a-fA-F0-9]{64}$') { throw "BACKUP_MANIFEST_METADATA_INVALID: $path" }
        }
        foreach ($entry in $archive.Entries) {
            if (-not $entry.FullName.EndsWith('/') -and $entry.FullName -ne 'manifest.json' -and -not $selected.Contains($entry.FullName)) { throw "BACKUP_UNDECLARED_ENTRY: $($entry.FullName)" }
        }
        if (-not ($files.path -match '\.mv\.db$')) { throw 'BACKUP_DATABASE_MISSING' }
        $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot([IO.Path]::GetFullPath($StagingRoot)))
        if ($drive.AvailableFreeSpace -lt $total + 64MB) { throw 'BACKUP_STAGING_SPACE_INSUFFICIENT' }
        foreach ($file in $files) {
            $entry = $entries[[string]$file.path]
            $target = Join-Path $StagingRoot $file.path
            New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
            $inputStream = $entry.Open()
            $outputStream = [IO.File]::Open($target, [IO.FileMode]::CreateNew)
            $hash = [Security.Cryptography.SHA256]::Create()
            try {
                $written = [long]0
                $buffer = [byte[]]::new(65536)
                while (($read = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    if ($written + $read -gt [long]$file.length) { throw "BACKUP_FILE_LENGTH_MISMATCH: $($file.path)" }
                    [void]$hash.TransformBlock($buffer, 0, $read, $null, 0)
                    $outputStream.Write($buffer, 0, $read)
                    $written += $read
                }
                [void]$hash.TransformFinalBlock([byte[]]::new(0), 0, 0)
                $actualHash = [BitConverter]::ToString($hash.Hash).Replace('-', '').ToLowerInvariant()
                if ($written -ne [long]$file.length) { throw "BACKUP_FILE_LENGTH_MISMATCH: $($file.path)" }
                if ($actualHash -ne ([string]$file.sha256).ToLowerInvariant()) { throw "BACKUP_FILE_CHECKSUM_MISMATCH: $($file.path)" }
            } finally { $inputStream.Dispose(); $outputStream.Dispose(); $hash.Dispose() }
        }
        return $files
    } finally {
        if ($null -ne $archive) { $archive.Dispose() } else { $archiveStream.Dispose() }
    }
}
