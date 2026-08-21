[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPdf,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [ValidateRange(72, 600)][int]$Dpi = 144,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Await-WinRtOperation {
    param([Parameter(Mandatory = $true)]$Operation, [Parameter(Mandatory = $true)][Type]$ResultType)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1
    } | Select-Object -First 1
    $task = $method.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    $task.GetAwaiter().GetResult()
}

function Await-WinRtAction {
    param([Parameter(Mandatory = $true)]$Action)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1
    } | Select-Object -First 1
    $task = $method.Invoke($null, @($Action))
    $task.GetAwaiter().GetResult()
}

$pdfPath = [IO.Path]::GetFullPath($InputPdf)
$outputPath = [IO.Path]::GetFullPath($OutputDirectory)
if (-not (Test-Path -LiteralPath $pdfPath -PathType Leaf)) { throw "PDF was not found: $pdfPath" }
if ([IO.Path]::GetExtension($pdfPath) -ne '.pdf') { throw "InputPdf must end in .pdf: $pdfPath" }

if (Test-Path -LiteralPath $outputPath) {
    $existing = @(Get-ChildItem -LiteralPath $outputPath -File -Filter 'page-*.png')
    if ($existing.Count -gt 0 -and -not $Force) { throw "Page images already exist. Pass -Force to replace them: $outputPath" }
    if ($Force) { $existing | Remove-Item -Force }
} else {
    New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.DataReader, Windows.Storage.Streams, ContentType = WindowsRuntime]

$file = Await-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath)) ([Windows.Storage.StorageFile])
$pdf = Await-WinRtOperation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
$scale = $Dpi / 96.0
$written = @()
try {
    for ($index = 0; $index -lt $pdf.PageCount; $index++) {
        $page = $pdf.GetPage($index)
        $stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
        $reader = $null
        try {
            $options = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
            $options.DestinationWidth = [uint32][Math]::Max(1, [Math]::Round($page.Dimensions.MediaBox.Width * $scale))
            $options.DestinationHeight = [uint32][Math]::Max(1, [Math]::Round($page.Dimensions.MediaBox.Height * $scale))
            [void](Await-WinRtAction ($page.RenderToStreamAsync($stream, $options)))
            $stream.Seek(0)
            $reader = [Windows.Storage.Streams.DataReader]::new($stream.GetInputStreamAt(0))
            [void](Await-WinRtOperation ($reader.LoadAsync([uint32]$stream.Size)) ([uint32]))
            $bytes = New-Object byte[] ([int]$stream.Size)
            $reader.ReadBytes($bytes)
            $target = Join-Path $outputPath ('page-{0:D4}.png' -f ($index + 1))
            [IO.File]::WriteAllBytes($target, $bytes)
            $written += $target
        } finally {
            if ($null -ne $reader) { $reader.Dispose() }
            $stream.Dispose()
            $page.Dispose()
        }
    }
} finally { }

[pscustomobject]@{
    InputPdf = $pdfPath
    OutputDirectory = $outputPath
    Dpi = $Dpi
    PageCount = $written.Count
    FirstImage = if ($written.Count) { $written[0] } else { $null }
    LastImage = if ($written.Count) { $written[-1] } else { $null }
}
