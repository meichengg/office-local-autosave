param(
    [int]$IntervalSeconds = 10,
    [int]$BackupIntervalSeconds = 3600,
    [int]$BackupKeepDays = 2,
    [int]$BackupMaxMB = 2048,
    [int]$MinFreeSpaceMB = 10240
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Name = 'OfficeLocalAutoSave'
$InstallDir = Join-Path $env:LOCALAPPDATA $Name
$ScriptSrc = Join-Path $PSScriptRoot 'OfficeLocalAutoSave.ps1'
$ScriptDst = Join-Path $InstallDir 'OfficeLocalAutoSave.ps1'
$RunKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Copy-Item -LiteralPath $ScriptSrc -Destination $ScriptDst -Force

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$args = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptDst`" -IntervalSeconds $IntervalSeconds -BackupIntervalSeconds $BackupIntervalSeconds -BackupKeepDays $BackupKeepDays -BackupMaxMB $BackupMaxMB -MinFreeSpaceMB $MinFreeSpaceMB"
$runCommand = "`"$ps`" $args"

Get-CimInstance Win32_Process -Filter "name='powershell.exe' or name='pwsh.exe'" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like '*OfficeLocalAutoSave.ps1*'
} | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
}

Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue
New-Item -Path $RunKey -Force | Out-Null
New-ItemProperty -Path $RunKey -Name $Name -Value $runCommand -PropertyType String -Force | Out-Null
Start-Process -FilePath $ps -ArgumentList $args -WindowStyle Hidden | Out-Null

Write-Host "Installed and started $Name via HKCU Run"
Write-Host "Script: $ScriptDst"
Write-Host "Log: $InstallDir\autosave.log"
