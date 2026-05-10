param(
    [int]$IntervalSeconds = 1,
    [int]$MinIdleSeconds = 3,
    [int]$ForegroundMinIdleSeconds = 10,
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

$IdleApiAvailable = $true
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class OfficeLocalAutoSaveInput
{
    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    [DllImport("user32.dll")]
    public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("kernel32.dll")]
    public static extern uint GetTickCount();
}
"@ -ErrorAction Stop
} catch {
    $IdleApiAvailable = $false
    Log "idle api unavailable error=$($_.Exception.Message)"
}

function Get-UserIdleSeconds {
    if (-not $IdleApiAvailable) { return 0 }
    try {
        $info = New-Object OfficeLocalAutoSaveInput+LASTINPUTINFO
        $info.cbSize = [uint32][Runtime.InteropServices.Marshal]::SizeOf($info)
        if (-not [OfficeLocalAutoSaveInput]::GetLastInputInfo([ref]$info)) { return 0 }
        $nowTick = [uint64][OfficeLocalAutoSaveInput]::GetTickCount()
        $lastTick = [uint64]$info.dwTime
        if ($nowTick -lt $lastTick) { $nowTick += [uint64]4294967296 }
        return [int](($nowTick - $lastTick) / 1000)
    } catch {
        return 0
    }
}

function Test-UserIdleForAutosave {
    return ((Get-UserIdleSeconds) -ge $MinIdleSeconds)
}

function Test-SafeToProbeOffice {
    $idleSeconds = Get-UserIdleSeconds
    if ($idleSeconds -lt $MinIdleSeconds) { return $false }
    $foregroundPid = Get-ForegroundProcessId
    if ($foregroundPid -eq 0) { return $true }
    try {
        $processName = (Get-Process -Id $foregroundPid -ErrorAction SilentlyContinue).ProcessName
        if (($processName -eq 'WINWORD' -or $processName -eq 'EXCEL') -and $idleSeconds -lt $ForegroundMinIdleSeconds) { return $false }
    } catch {}
    return $true
}

function Get-WindowProcessId($hwnd) {
    if ($null -eq $hwnd -or $hwnd -eq [IntPtr]::Zero) { return 0 }
    try {
        $pid = [uint32]0
        [void][OfficeLocalAutoSaveInput]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$pid)
        return [int]$pid
    } catch {
        return 0
    }
}

function Get-ForegroundProcessId {
    if (-not $IdleApiAvailable) { return 0 }
    try {
        return (Get-WindowProcessId ([OfficeLocalAutoSaveInput]::GetForegroundWindow()))
    } catch {
        return 0
    }
}

function Get-OfficeProcessId($app) {
    try {
        return (Get-WindowProcessId ([IntPtr]$app.Hwnd))
    } catch {
        return 0
    }
}

function Test-OfficeForeground($app) {
    $foregroundPid = Get-ForegroundProcessId
    $officePid = Get-OfficeProcessId $app
    return ($officePid -ne 0 -and $foregroundPid -eq $officePid)
}

function Test-OfficeSafeToSave($app) {
    $idleSeconds = Get-UserIdleSeconds
    if ($idleSeconds -lt $MinIdleSeconds) { return $false }
    if ((Test-OfficeForeground $app) -and $idleSeconds -lt $ForegroundMinIdleSeconds) { return $false }
    return $true
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
    if ($null -ne $obj) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
    }
}

function Save-Word {
    if (-not (Test-SafeToProbeOffice)) { return }
    try { $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { return }
    if ($null -eq $word) { return }
    try {
        if (-not (Test-OfficeSafeToSave $word)) { return }
        foreach ($doc in @($word.Documents)) {
            try {
                if (-not (Test-OfficeSafeToSave $word)) { return }
                if ($doc.Path -and -not $doc.ReadOnly -and -not $doc.Saved) {
                    $fullName = $doc.FullName
                    Ensure-Backup 'Word' $fullName
                    if (-not (Test-OfficeSafeToSave $word)) {
                        Log "Word save deferred active_input file=$fullName"
                        continue
                    }
                    $doc.Save()
                    Log "Word saved $fullName"
                }
            } catch { Log "Word error $($_.Exception.Message)" }
            finally { Release-Com $doc }
        }
    } finally {
        Release-Com $word
    }
}

function Save-Excel {
    if (-not (Test-SafeToProbeOffice)) { return }
    try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { return }
    if ($null -eq $excel) { return }
    try {
        if (-not (Test-OfficeSafeToSave $excel)) { return }
        try { if ($excel.Ready -eq $false) { return } } catch { return }
        foreach ($wb in @($excel.Workbooks)) {
            try {
                if (-not (Test-OfficeSafeToSave $excel)) { return }
                if ($wb.Path -and -not $wb.ReadOnly -and -not $wb.Saved) {
                    $fullName = $wb.FullName
                    Ensure-Backup 'Excel' $fullName
                    if (-not (Test-OfficeSafeToSave $excel)) {
                        Log "Excel save deferred active_input file=$fullName"
                        continue
                    }
                    try {
                        if ($excel.Ready -eq $false) {
                            Log "Excel save deferred busy file=$fullName"
                            continue
                        }
                    } catch {
                        Log "Excel save deferred busy file=$fullName"
                        continue
                    }
                    $wb.Save()
                    Log "Excel saved $fullName"
                }
            } catch { Log "Excel error $($_.Exception.Message)" }
            finally { Release-Com $wb }
        }
    } finally {
        Release-Com $excel
    }
}

$MaintenanceIntervalSeconds = 900
$LastMaintenance = [datetime]::MinValue
Log "watchdog started interval=${IntervalSeconds}s idle=${MinIdleSeconds}s foreground_idle=${ForegroundMinIdleSeconds}s backup=${BackupIntervalSeconds}s keep=${BackupKeepDays}d quota=${BackupMaxMB}MB minfree=${MinFreeSpaceMB}MB"
while ($true) {
    if (Test-SafeToProbeOffice) {
        try { Save-Word } catch {}
        try { Save-Excel } catch {}
        $now = Get-Date
        if (($now - $LastMaintenance).TotalSeconds -ge $MaintenanceIntervalSeconds) {
            try { Cleanup-Backups } catch {}
            $LastMaintenance = $now
        }
        try { $State | ConvertTo-Json | Set-Content -Path $StatePath -Encoding UTF8 } catch {}
    }
    Start-Sleep -Seconds $IntervalSeconds
}
