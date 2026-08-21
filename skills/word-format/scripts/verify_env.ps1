[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'windows_common.ps1')

$missing = $false
Write-Output '[Applications]'

try {
    $word = New-Object -ComObject Word.Application
    $wordVersion = $word.Version
    $word.DisplayAlerts = 0
    $word.Quit()
    Release-ComObject $word
    Write-Output ("  OK Microsoft Word (COM automation, version {0})" -f $wordVersion)
} catch {
    $missing = $true
    Write-Output '  MISSING Microsoft Word (desktop version with COM support)'
}

try {
    $browser = Get-BrowserInfo
    Write-Output ("  OK {0}: {1}" -f $browser.Name, $browser.Path)
} catch {
    $missing = $true
    Write-Output '  MISSING Google Chrome or Microsoft Edge'
}

Write-Output ''
Write-Output '[Runtime]'
if ($PSVersionTable.PSEdition -eq 'Desktop' -and $PSVersionTable.PSVersion.Major -ge 5) {
    Write-Output ("  OK Windows PowerShell {0}" -f $PSVersionTable.PSVersion)
} else {
    $missing = $true
    Write-Output '  MISSING Windows PowerShell 5.1 (run with powershell.exe, not pwsh.exe)'
}

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
    Write-Output '  OK Windows PDF renderer'
} catch {
    $missing = $true
    Write-Output '  MISSING Windows PDF renderer (Windows.Data.Pdf)'
}

Write-Output ''
if ($missing) {
    Write-Error 'Environment is incomplete. Install the missing applications and run this check again.'
    exit 1
}
Write-Output 'Environment is ready.'
