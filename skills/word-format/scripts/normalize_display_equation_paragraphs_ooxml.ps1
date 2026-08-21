[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputDocx,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$inputPath = [IO.Path]::GetFullPath($InputDocx)
$outputPath = [IO.Path]::GetFullPath($OutputDocx)
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "DOCX was not found: $inputPath" }
if ([StringComparer]::OrdinalIgnoreCase.Equals($inputPath, $outputPath)) { throw 'OutputDocx must be different from InputDocx.' }
if ((Test-Path -LiteralPath $outputPath -PathType Leaf) -and -not $Force) { throw "Output already exists. Pass -Force to replace it: $outputPath" }
$outputDirectory = Split-Path -Parent $outputPath
if ($outputDirectory) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
$tmpPath = Join-Path $outputDirectory ('.' + [IO.Path]::GetFileName($outputPath) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')

function Read-XmlEntry {
    param([IO.Compression.ZipArchive]$Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { throw "DOCX entry not found: $Name" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try {
        $xml = [Xml.XmlDocument]::new()
        $xml.PreserveWhitespace = $true
        $xml.LoadXml($reader.ReadToEnd())
        return $xml
    } finally { $reader.Dispose() }
}

function Convert-XmlToBytes {
    param([Xml.XmlDocument]$Xml)
    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $false
    $memory = [IO.MemoryStream]::new()
    try {
        $writer = [Xml.XmlWriter]::Create($memory, $settings)
        try { $Xml.Save($writer); $writer.Flush() } finally { $writer.Dispose() }
        return $memory.ToArray()
    } finally { $memory.Dispose() }
}

function Copy-ZipEntry {
    param([IO.Compression.ZipArchiveEntry]$SourceEntry, [IO.Compression.ZipArchive]$DestinationZip)
    $destination = $DestinationZip.CreateEntry($SourceEntry.FullName, [IO.Compression.CompressionLevel]::Optimal)
    $input = $SourceEntry.Open(); $output = $destination.Open()
    try { $input.CopyTo($output) } finally { $output.Dispose(); $input.Dispose() }
}

function Get-ParagraphText {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    return (@($Paragraph.SelectNodes('.//w:t', $Ns) | ForEach-Object { $_.InnerText }) -join '')
}

function Test-FormulaOnlyParagraph {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    if (-not [String]::IsNullOrWhiteSpace((Get-ParagraphText $Paragraph $Ns))) { return $false }
    if ($null -eq $Paragraph.SelectSingleNode('.//w:object', $Ns)) { return $false }
    return $Paragraph.OuterXml -match 'Equation\.DSMT4'
}

function Normalize-FormulaParagraph {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlDocument]$Document, [Xml.XmlNamespaceManager]$Ns)
    $pPr = $Paragraph.SelectSingleNode('./w:pPr', $Ns)
    if ($null -eq $pPr) {
        $pPr = $Document.CreateElement('w', 'pPr', $wNs)
        [void]$Paragraph.PrependChild($pPr)
    }
    foreach ($tab in @($Paragraph.SelectNodes('.//w:tab', $Ns))) { [void]$tab.ParentNode.RemoveChild($tab) }
    foreach ($node in @($pPr.SelectNodes('./w:spacing|./w:ind|./w:tabs|./w:jc', $Ns))) {
        [void]$node.ParentNode.RemoveChild($node)
    }
    $spacing = $Document.CreateElement('w', 'spacing', $wNs)
    foreach ($pair in @(@('before','0'), @('after','0'), @('line','288'), @('lineRule','auto'))) {
        [void]$spacing.SetAttribute($pair[0], $wNs, $pair[1])
    }
    [void]$pPr.AppendChild($spacing)
    $indent = $Document.CreateElement('w', 'ind', $wNs)
    foreach ($pair in @(@('left','0'), @('right','0'), @('firstLine','0'), @('firstLineChars','0'))) {
        [void]$indent.SetAttribute($pair[0], $wNs, $pair[1])
    }
    [void]$pPr.AppendChild($indent)
    $alignment = $Document.CreateElement('w', 'jc', $wNs)
    [void]$alignment.SetAttribute('val', $wNs, 'left')
    [void]$pPr.AppendChild($alignment)
}

function Test-RemovableBlankParagraph {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    if (-not [String]::IsNullOrWhiteSpace((Get-ParagraphText $Paragraph $Ns))) { return $false }
    if ($null -ne $Paragraph.SelectSingleNode('.//w:object|.//w:pict|.//w:drawing|.//w:br|.//w:bookmarkStart|.//w:bookmarkEnd|./w:pPr/w:sectPr|.//w:fldChar|.//w:instrText', $Ns)) { return $false }
    return $true
}

function Get-ParagraphStyleId {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $style = $Paragraph.SelectSingleNode('./w:pPr/w:pStyle', $Ns)
    if ($null -eq $style) { return '' }
    return $style.GetAttribute('val', $wNs)
}

function Test-SpecialBodyGeometry {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $text = ((Get-ParagraphText $Paragraph $Ns) -replace '\s+', ' ').Trim()
    if ($text -match '^\u56fe\s*\d+') { return $true }
    if ($null -ne $Paragraph.SelectSingleNode('./w:pPr/w:numPr|.//w:tab', $Ns)) { return $true }
    $alignment = $Paragraph.SelectSingleNode('./w:pPr/w:jc', $Ns)
    if ($null -ne $alignment -and $alignment.GetAttribute('val', $wNs) -in @('center','right','both','distribute')) { return $true }
    if ([String]::IsNullOrWhiteSpace($text) -and
        $null -ne $Paragraph.SelectSingleNode('.//w:drawing[not(ancestor::w:object)]|.//w:pict[not(ancestor::w:object)]', $Ns)) {
        return $true
    }
    return $false
}

$inputZip = $null; $outputZip = $null
try {
    $inputZip = [IO.Compression.ZipFile]::OpenRead($inputPath)
    $document = Read-XmlEntry $inputZip 'word/document.xml'
    $ns = [Xml.XmlNamespaceManager]::new($document.NameTable); $ns.AddNamespace('w', $wNs)
    $normalized = 0; $removedTabs = 0; $removedBodyIndents = 0; $removedBlanks = 0
    foreach ($paragraph in @($document.SelectNodes('/w:document/w:body/w:p', $ns))) {
        if (Test-FormulaOnlyParagraph $paragraph $ns) {
            $removedTabs += @($paragraph.SelectNodes('.//w:tab', $ns)).Count
            Normalize-FormulaParagraph $paragraph $document $ns
            $normalized++
        }
    }
    foreach ($paragraph in @($document.SelectNodes('/w:document/w:body/w:p', $ns))) {
        if ((Get-ParagraphStyleId $paragraph $ns) -ne 'm1_left_body') { continue }
        if (Test-FormulaOnlyParagraph $paragraph $ns) { continue }
        if (Test-SpecialBodyGeometry $paragraph $ns) { continue }
        $indent = $paragraph.SelectSingleNode('./w:pPr/w:ind', $ns)
        if ($null -ne $indent) {
            [void]$indent.ParentNode.RemoveChild($indent)
            $removedBodyIndents++
        }
    }
    foreach ($paragraph in @($document.SelectNodes('/w:document/w:body/w:p', $ns))) {
        if (Test-RemovableBlankParagraph $paragraph $ns) {
            [void]$paragraph.ParentNode.RemoveChild($paragraph)
            $removedBlanks++
        }
    }
    $documentBytes = Convert-XmlToBytes $document
    $outputZip = [IO.Compression.ZipFile]::Open($tmpPath, [IO.Compression.ZipArchiveMode]::Create)
    foreach ($entry in $inputZip.Entries) {
        if ($entry.FullName -eq 'word/document.xml') {
            $destination = $outputZip.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
            $stream = $destination.Open()
            try { $stream.Write($documentBytes, 0, $documentBytes.Length) } finally { $stream.Dispose() }
        } else { Copy-ZipEntry $entry $outputZip }
    }
    $outputZip.Dispose(); $outputZip = $null
    $inputZip.Dispose(); $inputZip = $null
    Move-Item -LiteralPath $tmpPath -Destination $outputPath -Force
    Write-Output "Saved: $outputPath"
    Write-Output "Formula-only paragraphs normalized: $normalized"
    Write-Output "Positioning tabs removed: $removedTabs"
    Write-Output "Ordinary body direct indents removed: $removedBodyIndents"
    Write-Output "Removable blank paragraphs removed: $removedBlanks"
} finally {
    if ($null -ne $outputZip) { $outputZip.Dispose() }
    if ($null -ne $inputZip) { $inputZip.Dispose() }
    if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force }
}
