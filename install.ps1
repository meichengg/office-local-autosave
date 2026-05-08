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
$LauncherDst = Join-Path $InstallDir 'StartOfficeLocalAutoSave.vbs'
$RunKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Copy-Item -LiteralPath $ScriptSrc -Destination $ScriptDst -Force

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$wscript = Join-Path $env:SystemRoot 'System32\wscript.exe'
$args = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptDst`" -IntervalSeconds $IntervalSeconds -BackupIntervalSeconds $BackupIntervalSeconds -BackupKeepDays $BackupKeepDays -BackupMaxMB $BackupMaxMB -MinFreeSpaceMB $MinFreeSpaceMB"
$psCommand = "`"$ps`" $args"
$vbsCommand = $psCommand.Replace('"', '""')
$launcher = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "$vbsCommand", 0, False
"@
Set-Content -Path $LauncherDst -Value $launcher -Encoding ASCII
$runCommand = "`"$wscript`" //B `"$LauncherDst`""

Get-CimInstance Win32_Process -Filter "name='powershell.exe' or name='pwsh.exe' or name='wscript.exe'" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like '*OfficeLocalAutoSave.ps1*' -or $_.CommandLine -like '*StartOfficeLocalAutoSave.vbs*'
} | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
}

Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue
New-Item -Path $RunKey -Force | Out-Null
New-ItemProperty -Path $RunKey -Name $Name -Value $runCommand -PropertyType String -Force | Out-Null
Start-Process -FilePath $wscript -ArgumentList "//B `"$LauncherDst`"" -WindowStyle Hidden | Out-Null

Write-Host "Installed and started $Name via hidden HKCU Run launcher"
Write-Host "Script: $ScriptDst"
Write-Host "Launcher: $LauncherDst"
Write-Host "Log: $InstallDir\autosave.log"
