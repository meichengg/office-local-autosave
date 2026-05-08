$Name = 'OfficeLocalAutoSave'
$InstallDir = Join-Path $env:LOCALAPPDATA $Name
Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue
Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name $Name -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "name='powershell.exe' or name='pwsh.exe' or name='wscript.exe'" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like '*OfficeLocalAutoSave.ps1*' -or $_.CommandLine -like '*StartOfficeLocalAutoSave.vbs*'
} | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
}
Remove-Item -LiteralPath (Join-Path $InstallDir 'StartOfficeLocalAutoSave.vbs') -Force -ErrorAction SilentlyContinue
Write-Host "Uninstalled $Name"
Write-Host "Data kept at: $env:LOCALAPPDATA\OfficeLocalAutoSave"
