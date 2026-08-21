[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputDocx,
    [Parameter(Position = 1)][string]$OutputHtml
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'windows_common.ps1')

$inputPath = Resolve-SkillPath $InputDocx
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
    throw "Input .docx was not found: $inputPath"
}

if ([string]::IsNullOrWhiteSpace($OutputHtml)) {
    $cacheRoot = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [Environment]::GetFolderPath('LocalApplicationData') }
    $outputPath = Join-Path (Join-Path $cacheRoot 'word-format-skill') (([IO.Path]::GetFileNameWithoutExtension($inputPath)) + '.html')
} else {
    $outputPath = Resolve-SkillPath $OutputHtml
}

$outputDirectory = Split-Path -Parent $outputPath
if ($outputDirectory) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }

$equationOleCount = Get-EquationOleCount $inputPath
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath)
    # Word's filtered HTML format is 10 (wdFormatFilteredHTML).
    $document.SaveAs2($outputPath, 10)
    $document.Close(0)
    $document = $null
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
    }
    Release-ComObject $document
    Release-ComObject $word
}

if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    throw "Word did not produce the expected HTML file: $outputPath"
}

# Keep an invisible marker so the paste script can prevent accidental OLE loss
# in full-rewrite mode. The browser ignores HTML comments.
Add-Content -LiteralPath $outputPath -Encoding ASCII -Value ("<!-- word-format-skill: equation-ole-count={0} -->" -f $equationOleCount)
Write-Output ("Equation.DSMT4 OLE objects detected: {0}" -f $equationOleCount)

Write-Output $outputPath
