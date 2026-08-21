[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputDocx,
    [string]$StylePattern = '\u7ae0\u8282\u540d\u79f0|\u9898\u578b|\u6bcd\u9898|\u5b50\u9898|\u89e3\u6790\u52a0\u7c97'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'windows_common.ps1')

$inputPath = Resolve-SkillPath $InputDocx
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
    throw "Input .docx was not found: $inputPath"
}

Write-Output '[Matched styles]'
$styles = @(Get-DocxStyles $inputPath | Where-Object { $_.Name -match $StylePattern })
if ($styles.Count -eq 0) {
    Write-Output '  No matching styles found.'
} else {
    foreach ($style in $styles) {
        Write-Output ("  {0} (type={1}, custom={2})" -f $style.Name, $style.Type, $style.Custom)
    }
}

$equationCount = Get-EquationOleCount $inputPath
Write-Output ''
Write-Output '[Embedded equations]'
Write-Output ("  Equation.DSMT4: {0}" -f $equationCount)
if ($equationCount -gt 0) {
    Write-Output '  Preserve these objects by using -AppendTo. Browser HTML paste creates images, not editable MathType OLE objects.'
}
