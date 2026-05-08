param(
    [int]$IntervalSeconds = 10,
    [int]$BackupIntervalSeconds = 3600,
    [int]$BackupKeepDays = 2,
    [int]$BackupMaxMB = 2048,
    [int]$MinFreeSpaceMB = 10240
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

$Root = Join-Path $env:LOCALAPPDATA 'OfficeLocalAutoSave'
$BackupRoot = Join-Path $Root 'Backups'
$LogPath = Join-Path $Root 'autosave.log'
$StatePath = Join-Path $Root 'state.json'
New-Item -ItemType Directory -Force -Path $Root, $BackupRoot | Out-Null
$BackupMaxBytes = [int64]$BackupMaxMB * 1MB
$MinFreeSpaceBytes = [int64]$MinFreeSpaceMB * 1MB

$State = @{}
if (Test-Path $StatePath) {
    try {
        $json = Get-Content $StatePath -Raw | ConvertFrom-Json
        foreach ($prop in $json.PSObject.Properties) {
            $State[$prop.Name] = $prop.Value
        }
    } catch { $State = @{} }
}

function Log($msg) {
    $line = "{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

function SafeName($s) {
    $safe = $s -replace '[\\/:*?"<>|\s]+', '_'
    if ($safe.Length -gt 180) { $safe = $safe.Substring(0, 180) }
    return $safe
}

function Ensure-Backup($kind, $fullName) {
    if (-not (Test-Path -LiteralPath $fullName)) { return }
    $dir = Join-Path $BackupRoot $kind
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $key = "$kind|$fullName"
    $now = Get-Date
    if ($State.ContainsKey($key)) {
        $last = [datetime]$State[$key]
        if (($now - $last).TotalSeconds -lt $BackupIntervalSeconds) { return }
    }
    $source = Get-Item -LiteralPath $fullName
    $sourceBytes = [int64]$source.Length
    if ($BackupMaxBytes -le 0 -or $sourceBytes -gt $BackupMaxBytes) {
        $State[$key] = $now.ToString('o')
        Log "$kind backup skipped quota file=$fullName size=$sourceBytes quota=$BackupMaxBytes"
        return
    }
    $freeBytes = Get-BackupDriveFreeBytes
    if (($freeBytes - $sourceBytes) -lt $MinFreeSpaceBytes) {
        $State[$key] = $now.ToString('o')
        Log "$kind backup skipped low_space file=$fullName free=$freeBytes min=$MinFreeSpaceBytes size=$sourceBytes"
        return
    }
    if (-not (Ensure-BackupQuota $sourceBytes)) {
        $State[$key] = $now.ToString('o')
        Log "$kind backup skipped quota_full file=$fullName size=$sourceBytes quota=$BackupMaxBytes"
        return
    }
    $backup = Join-Path $dir ("{0}.{1}.bak" -f (SafeName $fullName), ($now.ToString('yyyyMMdd-HHmmss')))
    try {
        Copy-Item -LiteralPath $fullName -Destination $backup -Force
    } catch {
        $State[$key] = $now.ToString('o')
        Log "$kind backup failed file=$fullName error=$($_.Exception.Message)"
        return
    }
    $State[$key] = $now.ToString('o')
    Log "$kind backup $backup"
}

function Cleanup-Backups {
    Cleanup-BackupsByAge
    Trim-BackupQuota 0 | Out-Null
}

function Cleanup-BackupsByAge {
    if (-not (Test-Path $BackupRoot)) { return }
    Get-ChildItem -Path $BackupRoot -Recurse -File | Where-Object {
        $_.LastWriteTime -lt (Get-Date).AddDays(-$BackupKeepDays)
    } | Remove-Item -Force
}

function Get-BackupFiles {
    if (-not (Test-Path -LiteralPath $BackupRoot)) { return @() }
    return @(Get-ChildItem -Path $BackupRoot -Recurse -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime)
}

function Get-BackupBytes {
    $total = [int64]0
    foreach ($file in Get-BackupFiles) {
        $total += [int64]$file.Length
    }
    return $total
}

function Get-BackupDriveFreeBytes {
    $qualifier = Split-Path -Path $BackupRoot -Qualifier
    if ([string]::IsNullOrWhiteSpace($qualifier)) { return [int64]::MaxValue }
    $driveName = $qualifier.TrimEnd(':')
    $drive = Get-PSDrive -Name $driveName -ErrorAction SilentlyContinue
    if ($null -eq $drive) { return [int64]::MaxValue }
    return [int64]$drive.Free
}

function Trim-BackupQuota($requiredBytes) {
    if ($BackupMaxBytes -le 0) { return $false }
    $files = Get-BackupFiles
    $total = [int64]0
    foreach ($file in $files) {
        $total += [int64]$file.Length
    }
    foreach ($file in $files) {
        if (($total + [int64]$requiredBytes) -le $BackupMaxBytes) { break }
        try {
            $size = [int64]$file.Length
            Remove-Item -LiteralPath $file.FullName -Force
            $total -= $size
            Log "backup quota removed $($file.FullName)"
        } catch {}
    }
    return (($total + [int64]$requiredBytes) -le $BackupMaxBytes)
}

function Ensure-BackupQuota($requiredBytes) {
    Cleanup-BackupsByAge
    return (Trim-BackupQuota $requiredBytes)
}

function Release-Com($obj) {
    if ($null -ne $obj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

function Save-Word {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    if ($null -eq $word) { return }
    foreach ($doc in @($word.Documents)) {
        try {
            if ($doc.Path -and -not $doc.ReadOnly -and -not $doc.Saved) {
                Ensure-Backup 'Word' $doc.FullName
                $doc.Save()
                Log "Word saved $($doc.FullName)"
            }
        } catch { Log "Word error $($_.Exception.Message)" }
    }
    Release-Com $word
}

function Save-Excel {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    if ($null -eq $excel) { return }
    foreach ($wb in @($excel.Workbooks)) {
        try {
            if ($wb.Path -and -not $wb.ReadOnly -and -not $wb.Saved) {
                Ensure-Backup 'Excel' $wb.FullName
                $wb.Save()
                Log "Excel saved $($wb.FullName)"
            }
        } catch { Log "Excel error $($_.Exception.Message)" }
    }
    Release-Com $excel
}

Log "watchdog started interval=${IntervalSeconds}s backup=${BackupIntervalSeconds}s keep=${BackupKeepDays}d quota=${BackupMaxMB}MB minfree=${MinFreeSpaceMB}MB"
while ($true) {
    try { Save-Word } catch {}
    try { Save-Excel } catch {}
    try { Cleanup-Backups } catch {}
    try { $State | ConvertTo-Json | Set-Content -Path $StatePath -Encoding UTF8 } catch {}
    Start-Sleep -Seconds $IntervalSeconds
}
