[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputHtml,
    [string]$OutputDocx,
    [string]$AppendTo,
    [ValidateSet('Auto', 'Chrome', 'Edge')][string]$Browser = 'Auto',
    [double]$RenderDelaySeconds = 2.5,
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'windows_common.ps1')

$htmlPath = Resolve-SkillPath $InputHtml
if (-not (Test-Path -LiteralPath $htmlPath -PathType Leaf)) {
    throw "Input HTML was not found: $htmlPath"
}

$equationOleCount = 0
$marker = Select-String -LiteralPath $htmlPath -Pattern 'word-format-skill: equation-ole-count=([0-9]+)' -AllMatches | Select-Object -First 1
if ($null -ne $marker) {
    $match = [regex]::Match($marker.Line, 'equation-ole-count=([0-9]+)')
    if ($match.Success) { $equationOleCount = [int]$match.Groups[1].Value }
}
if ($equationOleCount -gt 0 -and -not $AppendTo) {
    throw "This HTML came from a template containing $equationOleCount Equation.DSMT4 MathType/OLE objects. New-document mode would rasterize them; use -AppendTo with the original .docx to preserve editability."
}
if ($equationOleCount -gt 0 -and $AppendTo) {
    Write-Warning ("The source template contains {0} Equation.DSMT4 OLE objects. They remain editable in the copied template; any equations pasted from browser HTML are images." -f $equationOleCount)
}

if ($AppendTo -and -not $OutputDocx) {
    throw '-AppendTo requires -OutputDocx.'
}

$outputPath = $null
if ($OutputDocx) {
    $outputPath = Resolve-SkillPath $OutputDocx
    $outputDirectory = Split-Path -Parent $outputPath
    if ($outputDirectory) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
}

$appendPath = $null
if ($AppendTo) {
    $appendPath = Resolve-SkillPath $AppendTo
    if (-not (Test-Path -LiteralPath $appendPath -PathType Leaf)) {
        throw "Reference .docx was not found: $appendPath"
    }
    if ([StringComparer]::OrdinalIgnoreCase.Equals($appendPath, $outputPath)) {
        throw '-AppendTo and -OutputDocx must be different paths.'
    }
    Copy-Item -LiteralPath $appendPath -Destination $outputPath -Force
}

$browserInfo = Get-BrowserInfo $Browser
$htmlUri = ([Uri]$htmlPath).AbsoluteUri
$originalWindow = [WordFormatSkillNativeMethods]::GetForegroundWindow()
$browserProcess = $null
$word = $null
$document = $null
$range = $null
$afterPasteRange = $null
$profilePath = $null

try {
    Write-Output ("Browser: {0}" -f $browserInfo.Name)
    Write-Output ("Render:  {0}" -f $htmlPath)

    # Use an isolated profile so an existing browser window cannot receive clipboard operations.
    $profilePath = Join-Path ([IO.Path]::GetTempPath()) ('word-format-skill-browser-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $profilePath -Force | Out-Null
    $browserArgs = @(
        ('--user-data-dir="{0}"' -f $profilePath),
        '--remote-debugging-port=0',
        '--remote-allow-origins=*',
        '--no-first-run',
        '--no-default-browser-check',
        '--new-window',
        $htmlUri
    )
    $browserProcess = Start-Process -FilePath $browserInfo.Path -ArgumentList $browserArgs -PassThru
    $page = Wait-DevToolsPage -ProfilePath $profilePath
    Start-Sleep -Milliseconds ([Math]::Max(0, [int]($RenderDelaySeconds * 1000)))
    Copy-DevToolsPageToClipboard -WebSocketUrl $page.webSocketDebuggerUrl
    Start-Sleep -Milliseconds 300

    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    if ($appendPath) {
        Write-Output ("Mode:    append to template; output: {0}" -f $outputPath)
        $document = $word.Documents.Open($outputPath)
    } else {
        Write-Output 'Mode:    new document'
        $document = $word.Documents.Add()
    }

    Start-Sleep -Milliseconds 800
    $range = $document.Content
    $contentEndBeforePaste = [int]$range.End
    $range.Collapse(0) # wdCollapseEnd
    $range.Paste()
    Start-Sleep -Milliseconds 300
    $afterPasteRange = $document.Range()
    if ([int]$afterPasteRange.End -le $contentEndBeforePaste) {
        throw 'Word paste produced no content. Increase -RenderDelaySeconds and keep the desktop session focused.'
    }

    if ($appendPath) {
        $document.Save()
    } elseif ($outputPath) {
        # 16 = wdFormatDocumentDefault (.docx).
        $document.SaveAs2($outputPath, 16)
    } else {
        Write-Warning 'No -OutputDocx was supplied; the new Word document remains unsaved.'
    }

    if ($outputPath) { Write-Output ("Saved:   {0}" -f $outputPath) }
    if (-not $KeepOpen) {
        Release-ComObject $afterPasteRange
        $afterPasteRange = $null
        Release-ComObject $range
        $range = $null
        $document.Close(0)
        Release-ComObject $document
        $document = $null
        $word.Quit()
        Release-ComObject $word
        $word = $null
    }
    Write-Output 'Done.'
} finally {
    Release-ComObject $afterPasteRange
    Release-ComObject $range
    Release-ComObject $document
    Release-ComObject $word
    if (-not $KeepOpen -and $profilePath) {
        Stop-IsolatedBrowser -ProcessName $browserInfo.ProcessName -ProfilePath $profilePath
    }
    Bring-WindowToFront $originalWindow
}
