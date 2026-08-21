$ErrorActionPreference = 'Stop'

$source = Join-Path $PSScriptRoot 'renderer32\OlePreviewRenderer.cs'
$output = Join-Path $PSScriptRoot 'renderer32\OlePreviewRenderer.exe'
$compiler = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe'

& $compiler /nologo /platform:x86 /optimize+ "/out:$output" $source
if ($LASTEXITCODE -ne 0) {
    throw "32-bit renderer compilation failed with exit code $LASTEXITCODE"
}

Write-Output "built=$output"
