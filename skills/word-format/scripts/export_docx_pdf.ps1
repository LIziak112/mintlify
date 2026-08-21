[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$inputPath = [IO.Path]::GetFullPath($InputDocx)
$outputPath = [IO.Path]::GetFullPath($OutputPdf)
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "DOCX was not found: $inputPath" }
if ([IO.Path]::GetExtension($inputPath) -ne '.docx') { throw "InputDocx must end in .docx: $inputPath" }
if ([IO.Path]::GetExtension($outputPath) -ne '.pdf') { throw "OutputPdf must end in .pdf: $outputPath" }
if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
    if (-not $Force) { throw "PDF already exists. Pass -Force to replace it: $outputPath" }
    Remove-Item -LiteralPath $outputPath -Force
}
$outputDirectory = Split-Path -Parent $outputPath
if ($outputDirectory) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }

$tempPdf = Join-Path $outputDirectory ('.' + [IO.Path]::GetFileNameWithoutExtension($outputPath) + '.' + [Guid]::NewGuid().ToString('N') + '.pdf')
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    $pageCount = [int]$document.ComputeStatistics(2)
    $document.ExportAsFixedFormat($tempPdf, 17)
    $document.Close(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null
    $word.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    $word = $null
    Move-Item -LiteralPath $tempPdf -Destination $outputPath -Force
    [pscustomobject]@{ InputDocx = $inputPath; OutputPdf = $outputPath; PageCount = $pageCount }
} finally {
    if ($null -ne $document) { try { $document.Close(0) } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($null -ne $word) { try { $word.Quit(0) } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
    if (Test-Path -LiteralPath $tempPdf) { Remove-Item -LiteralPath $tempPdf -Force }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
