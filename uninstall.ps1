$Name = 'OfficeLocalAutoSave'
$InstallDir = Join-Path $env:ProgramData $Name
$DataDir = Join-Path $env:LOCALAPPDATA $Name

function Test-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Remove-LegacyRunEntries {
    Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name $Name -ErrorAction SilentlyContinue
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
    try {
        $process = Start-Process -FilePath $ps -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$self`"" -Verb RunAs -Wait -PassThru
        exit $process.ExitCode
    } catch {
        Write-Error "Administrator permission is required to uninstall $Name for all Windows users."
        exit 1
    }
}

Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue
Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -Name $Name -ErrorAction SilentlyContinue
Remove-LegacyRunEntries
Get-CimInstance Win32_Process -Filter "name='powershell.exe' or name='pwsh.exe' or name='wscript.exe'" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like '*OfficeLocalAutoSave.ps1*' -or $_.CommandLine -like '*StartOfficeLocalAutoSave.vbs*'
} | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
}
Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Uninstalled $Name"
Write-Host "Per-user data kept at: $DataDir"
