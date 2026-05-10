param(
    [int]$IntervalSeconds = 1,
    [int]$MinIdleSeconds = 3,
    [int]$ForegroundMinIdleSeconds = 10,
    [int]$BackupIntervalSeconds = 3600,
    [int]$BackupKeepDays = 2,
    [int]$BackupMaxMB = 2048,
    [int]$MinFreeSpaceMB = 10240,
    [switch]$SkipStart
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Name = 'OfficeLocalAutoSave'
$InstallDir = Join-Path $env:ProgramData $Name
$ScriptSrc = Join-Path $PSScriptRoot 'OfficeLocalAutoSave.ps1'
$ScriptDst = Join-Path $InstallDir 'OfficeLocalAutoSave.ps1'
$LauncherDst = Join-Path $InstallDir 'StartOfficeLocalAutoSave.vbs'
$RunKey = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'
$LegacyRunKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'

function Test-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-WatchdogLauncher($launcherPath) {
    $wscript = Join-Path $env:SystemRoot 'System32\wscript.exe'
    try {
        $shell = New-Object -ComObject Shell.Application
        $shell.ShellExecute($wscript, "//B `"$launcherPath`"", '', 'open', 0)
    } catch {
        Start-Process -FilePath $wscript -ArgumentList "//B `"$launcherPath`"" -WindowStyle Hidden | Out-Null
    }
}

function Remove-LegacyRunEntries {
    Remove-ItemProperty -Path $LegacyRunKey -Name $Name -ErrorAction SilentlyContinue
    Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Where-Object {
        $_.PSChildName -notlike '*_Classes'
    } | ForEach-Object {
        $runPath = Join-Path $_.PSPath 'Software\Microsoft\Windows\CurrentVersion\Run'
        Remove-ItemProperty -Path $runPath -Name $Name -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Admin)) {
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $self = $MyInvocation.MyCommand.Path
    $elevatedArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$self`" -IntervalSeconds $IntervalSeconds -MinIdleSeconds $MinIdleSeconds -ForegroundMinIdleSeconds $ForegroundMinIdleSeconds -BackupIntervalSeconds $BackupIntervalSeconds -BackupKeepDays $BackupKeepDays -BackupMaxMB $BackupMaxMB -MinFreeSpaceMB $MinFreeSpaceMB -SkipStart"
    try {
        $process = Start-Process -FilePath $ps -ArgumentList $elevatedArgs -Verb RunAs -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Remove-LegacyRunEntries
            $launcherPath = Join-Path (Join-Path $env:ProgramData $Name) 'StartOfficeLocalAutoSave.vbs'
            if (Test-Path -LiteralPath $launcherPath) {
                Start-WatchdogLauncher $launcherPath
            }
        }
        exit $process.ExitCode
    } catch {
        Write-Error "Administrator permission is required to install $Name for all Windows users."
        exit 1
    }
}

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Copy-Item -LiteralPath $ScriptSrc -Destination $ScriptDst -Force

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$wscript = Join-Path $env:SystemRoot 'System32\wscript.exe'
$args = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptDst`" -IntervalSeconds $IntervalSeconds -MinIdleSeconds $MinIdleSeconds -ForegroundMinIdleSeconds $ForegroundMinIdleSeconds -BackupIntervalSeconds $BackupIntervalSeconds -BackupKeepDays $BackupKeepDays -BackupMaxMB $BackupMaxMB -MinFreeSpaceMB $MinFreeSpaceMB"
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
Remove-LegacyRunEntries
New-Item -Path $RunKey -Force | Out-Null
New-ItemProperty -Path $RunKey -Name $Name -Value $runCommand -PropertyType String -Force | Out-Null
if (-not $SkipStart) {
    Start-WatchdogLauncher $LauncherDst
}

if ($SkipStart) {
    Write-Host "Installed $Name for all Windows users via hidden HKLM Run launcher"
} else {
    Write-Host "Installed and started $Name for all Windows users via hidden HKLM Run launcher"
}
Write-Host "Script: $ScriptDst"
Write-Host "Launcher: $LauncherDst"
Write-Host "Per-user log: %LOCALAPPDATA%\OfficeLocalAutoSave\autosave.log"
