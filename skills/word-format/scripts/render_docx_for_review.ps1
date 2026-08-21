[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [ValidateRange(72, 600)][int]$Dpi = 144,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$inputPath = [IO.Path]::GetFullPath($InputDocx)
$outputPath = [IO.Path]::GetFullPath($OutputDirectory)
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "DOCX was not found: $inputPath" }

if (Test-Path -LiteralPath $outputPath) {
    $known = @(
        Get-ChildItem -LiteralPath $outputPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq 'document.pdf' -or $_.Name -eq 'review.html' -or $_.Name -eq 'render-report.json' -or $_.Name -like 'page-*.png' }
    )
    $unknown = @(Get-ChildItem -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue | Where-Object { $_ -notin $known })
    if (($known.Count -gt 0 -or $unknown.Count -gt 0) -and -not $Force) { throw "OutputDirectory is not empty. Pass -Force to replace this renderer's files: $outputPath" }
    if ($Force) { $known | Remove-Item -Force }
} else {
    New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
}

$pdfPath = Join-Path $outputPath 'document.pdf'
$export = & (Join-Path $PSScriptRoot 'export_docx_pdf.ps1') -InputDocx $inputPath -OutputPdf $pdfPath -Force:$Force
$images = & (Join-Path $PSScriptRoot 'pdf_to_page_images.ps1') -InputPdf $pdfPath -OutputDirectory $outputPath -Dpi $Dpi -Force:$Force
$pageFiles = @(Get-ChildItem -LiteralPath $outputPath -File -Filter 'page-*.png' | Sort-Object Name)
if ($pageFiles.Count -ne $export.PageCount -or $pageFiles.Count -ne $images.PageCount) {
    throw "Render page count mismatch: Word=$($export.PageCount), PDF=$($images.PageCount), PNG=$($pageFiles.Count)"
}
if (@($pageFiles | Where-Object { $_.Length -eq 0 }).Count -gt 0) { throw 'One or more rendered page images are empty.' }

$escapedTitle = [Net.WebUtility]::HtmlEncode([IO.Path]::GetFileName($inputPath))
$cards = foreach ($page in $pageFiles) {
    $name = [Net.WebUtility]::HtmlEncode($page.Name)
    "<figure><figcaption>$name</figcaption><a href=`"$name`"><img loading=`"lazy`" src=`"$name`" alt=`"$name`"></a></figure>"
}
$html = @"
<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>$escapedTitle — rendered pages</title>
<style>body{margin:24px;background:#ddd;font:14px system-ui;color:#222}header{position:sticky;top:0;background:#fffffff0;padding:12px;z-index:1}figure{margin:24px auto;max-width:1200px}figcaption{margin-bottom:6px}img{display:block;width:100%;height:auto;background:white;box-shadow:0 2px 12px #888}</style>
</head><body><header><strong>$escapedTitle</strong> · $($pageFiles.Count) pages · $Dpi DPI</header>$($cards -join "`n")</body></html>
"@
$reviewPath = Join-Path $outputPath 'review.html'
[IO.File]::WriteAllText($reviewPath, $html, [Text.UTF8Encoding]::new($false))

$report = [ordered]@{
    input_docx = $inputPath
    input_sha256 = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
    pdf = $pdfPath
    dpi = $Dpi
    page_count = $pageFiles.Count
    page_images = @($pageFiles | ForEach-Object { [ordered]@{ name = $_.Name; bytes = $_.Length } })
    review_html = $reviewPath
    status = 'PASS'
}
$reportPath = Join-Path $outputPath 'render-report.json'
[IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 4), [Text.UTF8Encoding]::new($false))

[pscustomobject]@{
    InputDocx = $inputPath
    OutputDirectory = $outputPath
    Pdf = $pdfPath
    PageCount = $pageFiles.Count
    Review = $reviewPath
    Report = $reportPath
    Status = 'PASS'
} | Format-List
