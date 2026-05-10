param(
    [int]$WaitSeconds = 25,
    [switch]$SkipWord,
    [switch]$SkipExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$InstallRoot = Join-Path $env:ProgramData 'OfficeLocalAutoSave'
$DataRoot = Join-Path $env:LOCALAPPDATA 'OfficeLocalAutoSave'
$TestDir = Join-Path $env:TEMP 'OfficeLocalAutoSaveTest'
$CurrentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
New-Item -ItemType Directory -Force -Path $TestDir | Out-Null

function Assert($condition, $message) {
    if (-not $condition) { throw $message }
}

function Release-Com($obj) {
    if ($null -ne $obj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

function Get-WatchdogProcess {
    Get-CimInstance Win32_Process -Filter "name='powershell.exe' or name='pwsh.exe'" -ErrorAction SilentlyContinue | Where-Object {
        $_.SessionId -eq $CurrentSessionId -and $_.CommandLine -like '*OfficeLocalAutoSave.ps1*'
    }
}

function Test-WatchdogRunning {
    $runKey = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'
    $runValue = Get-ItemProperty -Path $runKey -Name 'OfficeLocalAutoSave' -ErrorAction SilentlyContinue
    Assert ($null -ne $runValue) 'HKLM Run entry OfficeLocalAutoSave not found. Run install.ps1 first.'
    Assert ($runValue.OfficeLocalAutoSave -like '*StartOfficeLocalAutoSave.vbs*') 'HKLM Run entry is old visible-console format. Run install.ps1 again.'

    $process = Get-WatchdogProcess
    if ($null -eq $process) {
        $scriptPath = Join-Path $InstallRoot 'OfficeLocalAutoSave.ps1'
        $launcherPath = Join-Path $InstallRoot 'StartOfficeLocalAutoSave.vbs'
        Assert (Test-Path -LiteralPath $scriptPath) "Watchdog script not found: $scriptPath"
        Assert (Test-Path -LiteralPath $launcherPath) "Hidden launcher not found: $launcherPath"
        $wscript = Join-Path $env:SystemRoot 'System32\wscript.exe'
        Start-Process -FilePath $wscript -ArgumentList "//B `"$launcherPath`"" -WindowStyle Hidden | Out-Null
        Start-Sleep -Seconds 2
        $process = Get-WatchdogProcess
    }

    Assert ($null -ne $process) 'OfficeLocalAutoSave watchdog process is not running. Run install.ps1 again.'
}

function Test-WordAutoSave {
    $path = Join-Path $TestDir 'word-autosave-test.docx'
    Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    $marker = 'WORD_AUTOSAVE_' + [guid]::NewGuid().ToString('N')
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $doc = $word.Documents.Add()
        $doc.Content.Text = 'initial'
        $doc.SaveAs2($path)
        $doc.Content.Text = $marker
        Write-Host "Waiting for Word autosave: $WaitSeconds seconds"
        Start-Sleep -Seconds $WaitSeconds
        Assert ($doc.Saved -eq $true) 'Word document is still dirty after wait; autosave did not run.'
        $doc.Close($false)
        $doc = $word.Documents.Open($path)
        Assert ($doc.Content.Text -like "*$marker*") 'Word reopened content does not contain marker.'
        Write-Host "Word autosave OK: $path"
    } finally {
        if ($null -ne $doc) { $doc.Close($false) }
        if ($null -ne $word) { $word.Quit() }
        Release-Com $doc
        Release-Com $word
    }
}

function Test-ExcelAutoSave {
    $path = Join-Path $TestDir 'excel-autosave-test.xlsx'
    Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    $marker = 'EXCEL_AUTOSAVE_' + [guid]::NewGuid().ToString('N')
    $excel = $null
    $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Range('A1').Value2 = 'initial'
        $wb.SaveAs($path)
        $ws.Range('A1').Value2 = $marker
        Write-Host "Waiting for Excel autosave: $WaitSeconds seconds"
        Start-Sleep -Seconds $WaitSeconds
        Assert ($wb.Saved -eq $true) 'Excel workbook is still dirty after wait; autosave did not run.'
        $wb.Close($false)
        $wb = $excel.Workbooks.Open($path)
        $ws = $wb.Worksheets.Item(1)
        Assert ($ws.Range('A1').Value2 -eq $marker) 'Excel reopened content does not contain marker.'
        Write-Host "Excel autosave OK: $path"
    } finally {
        if ($null -ne $wb) { $wb.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        Release-Com $wb
        Release-Com $excel
    }
}

Test-WatchdogRunning
if (-not $SkipWord) { Test-WordAutoSave }
if (-not $SkipExcel) { Test-ExcelAutoSave }
Write-Host "Verification completed. Log: $DataRoot\autosave.log"
