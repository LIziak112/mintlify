[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$TemplateDocx,
    [Parameter(Mandatory = $true)][string]$SourceDocx,
    [Parameter(Mandatory = $true)][string]$OutputDocx,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$templatePath = [IO.Path]::GetFullPath($TemplateDocx)
$sourcePath = [IO.Path]::GetFullPath($SourceDocx)
$outputPath = [IO.Path]::GetFullPath($OutputDocx)

foreach ($path in @($templatePath, $sourcePath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "DOCX was not found: $path" }
}
if ([StringComparer]::OrdinalIgnoreCase.Equals($templatePath, $sourcePath)) {
    throw 'TemplateDocx and SourceDocx must be different files.'
}
if ([StringComparer]::OrdinalIgnoreCase.Equals($sourcePath, $outputPath)) {
    throw 'OutputDocx must be different from SourceDocx.'
}
if ((Test-Path -LiteralPath $outputPath -PathType Leaf) -and -not $Force) {
    throw "Output already exists. Pass -Force to replace it: $outputPath"
}

$outputDirectory = Split-Path -Parent $outputPath
if ($outputDirectory) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
$tmpPath = Join-Path $outputDirectory ('.' + [IO.Path]::GetFileName($outputPath) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')

function New-XmlDocumentFromEntry {
    param([IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { throw "DOCX entry not found: $EntryName" }
    $stream = $entry.Open()
    try {
        $reader = [IO.StreamReader]::new($stream)
        try {
            $xml = [Xml.XmlDocument]::new()
            $xml.PreserveWhitespace = $true
            $xml.LoadXml($reader.ReadToEnd())
            return $xml
        } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
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

function Get-WVal {
    param([Xml.XmlElement]$Node)
    if ($null -eq $Node) { return $null }
    return $Node.GetAttribute('val', $wNs)
}

function Set-WVal {
    param([Xml.XmlElement]$Node, [string]$Value)
    if ($null -ne $Node) { [void]$Node.SetAttribute('val', $wNs, $Value) }
}

function Get-ParagraphText {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    return ((@($Paragraph.SelectNodes('.//w:t', $Ns) | ForEach-Object { $_.InnerText }) -join '') -replace '\s+', ' ').Trim()
}

function Test-HasObject {
    param([Xml.XmlElement]$Node, [Xml.XmlNamespaceManager]$Ns)
    return $null -ne $Node.SelectSingleNode('.//w:object|.//w:pict|.//w:drawing', $Ns)
}

function Test-EmptyParagraph {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    return ([String]::IsNullOrWhiteSpace((Get-ParagraphText $Paragraph $Ns)) -and -not (Test-HasObject $Paragraph $Ns))
}

function Get-StyleNameById {
    param(
        [Xml.XmlDocument]$Styles,
        [Xml.XmlNamespaceManager]$Ns,
        [string]$StyleId
    )
    if ([String]::IsNullOrWhiteSpace($StyleId)) { return '' }
    $style = $Styles.SelectSingleNode("/w:styles/w:style[@w:styleId='$StyleId']", $Ns)
    if ($null -eq $style) { return '' }
    $name = $style.SelectSingleNode('./w:name', $Ns)
    if ($null -eq $name) { return '' }
    return Get-WVal $name
}

function Test-DisplayEquationParagraph {
    param(
        [Xml.XmlElement]$Paragraph,
        [Xml.XmlNamespaceManager]$Ns,
        [string]$StyleName
    )
    if ($StyleName -eq 'MTDisplayEquation') { return $true }
    if (-not [String]::IsNullOrWhiteSpace((Get-ParagraphText $Paragraph $Ns))) { return $false }
    if ($null -eq $Paragraph.SelectSingleNode('./w:r[1]/w:tab', $Ns)) { return $false }
    foreach ($object in @($Paragraph.SelectNodes('.//w:object', $Ns))) {
        if ($object.OuterXml -match 'Equation\.DSMT4') { return $true }
    }
    return $false
}

function Test-SpecialBodyGeometry {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $text = Get-ParagraphText $Paragraph $Ns
    if ($text -match '^\u56fe\s*\d+') { return $true }
    if ($null -ne $Paragraph.SelectSingleNode('./w:pPr/w:numPr', $Ns)) { return $true }
    if ($null -ne $Paragraph.SelectSingleNode('.//w:tab', $Ns)) { return $true }
    $alignment = $Paragraph.SelectSingleNode('./w:pPr/w:jc', $Ns)
    if ($null -ne $alignment -and (Get-WVal $alignment) -in @('center','right','both','distribute')) { return $true }
    if ([String]::IsNullOrWhiteSpace($text) -and
        $null -ne $Paragraph.SelectSingleNode('.//w:drawing[not(ancestor::w:object)]|.//w:pict[not(ancestor::w:object)]', $Ns)) {
        return $true
    }
    return $false
}

function Clear-TemplateDirectIndent {
    param(
        [Xml.XmlElement]$Paragraph,
        [Xml.XmlNamespaceManager]$Ns,
        [string]$AssignedStyle,
        [string]$BodyStyle
    )
    $indent = $Paragraph.SelectSingleNode('./w:pPr/w:ind', $Ns)
    if ($null -eq $indent) { return $false }
    if ($AssignedStyle -eq $BodyStyle -and (Test-SpecialBodyGeometry $Paragraph $Ns)) { return $false }
    [void]$indent.ParentNode.RemoveChild($indent)
    return $true
}

function Normalize-DisplayEquationParagraph {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlDocument]$Document, [Xml.XmlNamespaceManager]$Ns)
    $pPr = $Paragraph.SelectSingleNode('./w:pPr', $Ns)
    if ($null -eq $pPr) {
        $pPr = $Document.CreateElement('w', 'pPr', $wNs)
        [void]$Paragraph.PrependChild($pPr)
    }
    foreach ($tab in @($Paragraph.SelectNodes('.//w:tab', $Ns))) {
        [void]$tab.ParentNode.RemoveChild($tab)
    }
    foreach ($spacing in @($pPr.SelectNodes('./w:spacing', $Ns))) {
        [void]$spacing.ParentNode.RemoveChild($spacing)
    }
    foreach ($ind in @($pPr.SelectNodes('./w:ind', $Ns))) {
        [void]$ind.ParentNode.RemoveChild($ind)
    }
    foreach ($tabs in @($pPr.SelectNodes('./w:tabs', $Ns))) {
        [void]$tabs.ParentNode.RemoveChild($tabs)
    }
    foreach ($jc in @($pPr.SelectNodes('./w:jc', $Ns))) {
        [void]$jc.ParentNode.RemoveChild($jc)
    }
    $spacing = $Document.CreateElement('w', 'spacing', $wNs)
    [void]$spacing.SetAttribute('before', $wNs, '0')
    [void]$spacing.SetAttribute('after', $wNs, '0')
    [void]$spacing.SetAttribute('line', $wNs, '288')
    [void]$spacing.SetAttribute('lineRule', $wNs, 'auto')
    [void]$pPr.AppendChild($spacing)
    $ind = $Document.CreateElement('w', 'ind', $wNs)
    [void]$ind.SetAttribute('left', $wNs, '0')
    [void]$ind.SetAttribute('right', $wNs, '0')
    [void]$ind.SetAttribute('firstLine', $wNs, '0')
    [void]$ind.SetAttribute('firstLineChars', $wNs, '0')
    [void]$pPr.AppendChild($ind)
    $jc = $Document.CreateElement('w', 'jc', $wNs)
    [void]$jc.SetAttribute('val', $wNs, 'left')
    [void]$pPr.AppendChild($jc)
}

function Test-RemovableBlankParagraph {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    if (-not (Test-EmptyParagraph $Paragraph $Ns)) { return $false }
    if ($null -ne $Paragraph.SelectSingleNode('.//w:br|.//w:bookmarkStart|.//w:bookmarkEnd|./w:pPr/w:sectPr', $Ns)) { return $false }
    if ($null -ne $Paragraph.SelectSingleNode('.//w:object|.//w:pict|.//w:drawing|.//w:fldChar|.//w:instrText', $Ns)) { return $false }
    return $true
}

function Remove-BlankParagraphsBefore {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $previous = $Paragraph.PreviousSibling
    while ($null -ne $previous -and $previous.LocalName -eq 'p' -and (Test-EmptyParagraph $previous $Ns)) {
        $remove = $previous
        $previous = $previous.PreviousSibling
        [void]$remove.ParentNode.RemoveChild($remove)
    }
}

function Add-ColumnBreak {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlDocument]$Document, [Xml.XmlNamespaceManager]$Ns)
    if ($null -ne $Paragraph.SelectSingleNode('.//w:br[@w:type="column"]', $Ns)) { return }
    $run = $Document.CreateElement('w', 'r', $wNs)
    $break = $Document.CreateElement('w', 'br', $wNs)
    [void]$break.SetAttribute('type', $wNs, 'column')
    [void]$run.AppendChild($break)
    $pPr = $Paragraph.SelectSingleNode('./w:pPr', $Ns)
    if ($null -ne $pPr) { [void]$Paragraph.InsertAfter($run, $pPr) }
    else { [void]$Paragraph.PrependChild($run) }
}

function Clear-StepDirectBold {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $runs = @($Paragraph.SelectNodes('./w:r', $Ns))
    if ($runs.Count -lt 2 -or -not (Test-HasObject $runs[0] $Ns)) { return $false }
    $firstText = $null
    foreach ($run in $runs) {
        $runText = ((@($run.SelectNodes('.//w:t', $Ns) | ForEach-Object { $_.InnerText }) -join ''))
        if ($runText) { $firstText = $runText; break }
    }
    if ($null -eq $firstText -or -not $firstText.StartsWith([char]0xFF1A)) { return $false }
    foreach ($run in $runs) {
        $rPr = $run.SelectSingleNode('./w:rPr', $Ns)
        if ($null -eq $rPr) { continue }
        foreach ($bold in @($rPr.SelectNodes('./w:b|./w:bCs', $Ns))) {
            [void]$bold.ParentNode.RemoveChild($bold)
        }
    }
    return $true
}

function Remove-DirectBold {
    param([Xml.XmlElement]$Run, [Xml.XmlNamespaceManager]$Ns)
    $rPr = $Run.SelectSingleNode('./w:rPr', $Ns)
    if ($null -eq $rPr) { return $false }
    $removed = $false
    foreach ($bold in @($rPr.SelectNodes('./w:b|./w:bCs', $Ns))) {
        [void]$bold.ParentNode.RemoveChild($bold)
        $removed = $true
    }
    return $removed
}

function Clear-AnalysisHeadDirectBold {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $prefixSeen = $false
    foreach ($run in @($Paragraph.SelectNodes('./w:r', $Ns))) {
        $runText = ((@($run.SelectNodes('.//w:t', $Ns) | ForEach-Object { $_.InnerText }) -join ''))
        if (-not $prefixSeen -and $runText -match '^\s*解析\s*[\uFF1A:]') {
            $prefixSeen = $true
            continue
        }
        if ($prefixSeen) { [void](Remove-DirectBold $run $Ns) }
    }
    return $prefixSeen
}

function Clear-AnalysisBodyDirectBold {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $changed = $false
    foreach ($run in @($Paragraph.SelectNodes('./w:r', $Ns))) {
        if (Remove-DirectBold $run $Ns) { $changed = $true }
    }
    return $changed
}

function Find-Style {
    param([Xml.XmlDocument]$Styles, [Xml.XmlNamespaceManager]$Ns, [string]$Id, [string]$Name)
    $style = $Styles.SelectSingleNode("/w:styles/w:style[@w:styleId='$Id']", $Ns)
    if ($null -eq $style -and $Name) {
        foreach ($candidate in $Styles.SelectNodes('/w:styles/w:style', $Ns)) {
            $nameNode = $candidate.SelectSingleNode('./w:name', $Ns)
            if ($null -ne $nameNode -and (Get-WVal $nameNode) -eq $Name) { return $candidate }
        }
    }
    return $style
}

function Copy-ZipEntry {
    param([IO.Compression.ZipArchiveEntry]$SourceEntry, [IO.Compression.ZipArchive]$DestinationZip)
    $destination = $DestinationZip.CreateEntry($SourceEntry.FullName, [IO.Compression.CompressionLevel]::Optimal)
    $input = $SourceEntry.Open(); $output = $destination.Open()
    try { $input.CopyTo($output) } finally { $output.Dispose(); $input.Dispose() }
}

function New-UnicodeString {
    param([int[]]$CodePoint)
    return -join ($CodePoint | ForEach-Object { [char]$_ })
}

$styleMap = [ordered]@{
    'af8' = @{ NewId = 'm1_chapter_name'; Name = (New-UnicodeString @(0x7AE0,0x8282,0x540D,0x79F0)) }
    'af9' = @{ NewId = 'm1_question_type'; Name = (New-UnicodeString @(0x9898,0x578B)) }
    'afa' = @{ NewId = 'm1_mother_no_gap'; Name = (New-UnicodeString @(0x6BCD,0x9898,0xFF08,0x4E0D,0x7A7A,0x884C,0xFF09)) }
    '11'  = @{ NewId = 'm1_label_body_bold'; Name = (New-UnicodeString @(0x6BCD,0x9898,0x6982,0x8FF0,0xFF0C,0x89E3,0x6790,0xFF0C,0x5B50,0x9898,0x0031)) }
    '12'  = @{ NewId = 'm1_subquestion'; Name = (New-UnicodeString @(0x5B50,0x9898,0x0031)) }
    '21'  = @{ NewId = 'm1_mother_gap'; Name = (New-UnicodeString @(0x6BCD,0x9898,0x0032,0xFF08,0x7A7A,0x4E24,0x884C,0xFF09)) }
    'afc' = @{ NewId = 'm1_left_body'; Name = (New-UnicodeString @(0x5DE6,0x9F50,0x6B63,0x6587)) }
    'afd' = @{ NewId = 'm1_analysis_bold'; Name = (New-UnicodeString @(0x89E3,0x6790,0x52A0,0x7C97)) }
}

$templateZip = $null; $sourceZip = $null; $outputZip = $null
try {
    $templateZip = [IO.Compression.ZipFile]::OpenRead($templatePath)
    $sourceZip = [IO.Compression.ZipFile]::OpenRead($sourcePath)
    $templateStyles = New-XmlDocumentFromEntry $templateZip 'word/styles.xml'
    $sourceStyles = New-XmlDocumentFromEntry $sourceZip 'word/styles.xml'
    $templateDoc = New-XmlDocumentFromEntry $templateZip 'word/document.xml'
    $sourceDoc = New-XmlDocumentFromEntry $sourceZip 'word/document.xml'

    $templateNs = [Xml.XmlNamespaceManager]::new($templateStyles.NameTable); $templateNs.AddNamespace('w', $wNs)
    $sourceStylesNs = [Xml.XmlNamespaceManager]::new($sourceStyles.NameTable); $sourceStylesNs.AddNamespace('w', $wNs)
    $sourceDocNs = [Xml.XmlNamespaceManager]::new($sourceDoc.NameTable); $sourceDocNs.AddNamespace('w', $wNs)
    $templateDocNs = [Xml.XmlNamespaceManager]::new($templateDoc.NameTable); $templateDocNs.AddNamespace('w', $wNs)

    $normalStyleId = '1'
    foreach ($style in $sourceStyles.SelectNodes('/w:styles/w:style', $sourceStylesNs)) {
        $nameNode = $style.SelectSingleNode('./w:name', $sourceStylesNs)
        if ($null -ne $nameNode -and (Get-WVal $nameNode) -eq 'Normal') { $normalStyleId = $style.GetAttribute('styleId', $wNs); break }
    }

    $sourceToNew = @{}
    $stylesRoot = $sourceStyles.SelectSingleNode('/w:styles', $sourceStylesNs)
    foreach ($sourceId in $styleMap.Keys) {
        $spec = $styleMap[$sourceId]
        $sourceStyle = Find-Style $templateStyles $templateNs $sourceId $spec.Name
        if ($null -eq $sourceStyle) { throw "Template style not found: $sourceId / $($spec.Name)" }
        $actualId = $sourceStyle.GetAttribute('styleId', $wNs)
        $sourceToNew[$actualId] = $spec.NewId
        foreach ($old in @($sourceStyles.SelectNodes("/w:styles/w:style[@w:styleId='$($spec.NewId)']", $sourceStylesNs))) { [void]$old.ParentNode.RemoveChild($old) }
        $clone = $sourceStyles.ImportNode($sourceStyle, $true)
        [void]$clone.SetAttribute('styleId', $wNs, $spec.NewId)
        [void]$stylesRoot.AppendChild($clone)
    }
    foreach ($clone in $sourceStyles.SelectNodes('/w:styles/w:style[starts-with(@w:styleId,"m1_")]', $sourceStylesNs)) {
        foreach ($link in $clone.SelectNodes('.//w:basedOn|.//w:next|.//w:link', $sourceStylesNs)) {
            $value = Get-WVal $link
            if ($sourceToNew.ContainsKey($value)) { Set-WVal $link $sourceToNew[$value] }
            elseif ($value -eq 'a') { Set-WVal $link $normalStyleId }
        }
    }

    $chapterStyle = $styleMap['af8'].NewId; $topicStyle = $styleMap['af9'].NewId
    $motherNoGap = $styleMap['afa'].NewId; $labelBold = $styleMap['11'].NewId
    $subStyle = $styleMap['12'].NewId; $motherGap = $styleMap['21'].NewId
    $bodyStyle = $styleMap['afc'].NewId; $analysisBold = $styleMap['afd'].NewId
    $motherSeen = $false
    $inAnalysis = $false
    $paragraphs = $sourceDoc.SelectNodes('/w:document/w:body/w:p', $sourceDocNs)
    foreach ($paragraph in @($paragraphs)) {
        $styleNode = $paragraph.SelectSingleNode('./w:pPr/w:pStyle', $sourceDocNs)
        $oldStyle = if ($null -ne $styleNode) { Get-WVal $styleNode } else { '' }
        $oldStyleName = Get-StyleNameById $sourceStyles $sourceStylesNs $oldStyle
        $text = Get-ParagraphText $paragraph $sourceDocNs
        $hasObject = Test-HasObject $paragraph $sourceDocNs
        $empty = [String]::IsNullOrWhiteSpace($text) -and -not $hasObject
        $isDisplayEquation = Test-DisplayEquationParagraph $paragraph $sourceDocNs $oldStyleName
        $newStyle = $null
        if ($isDisplayEquation) {
            Normalize-DisplayEquationParagraph $paragraph $sourceDoc $sourceDocNs
            $newStyle = $null
        }
        elseif ($oldStyle -eq '2' -or $text -match '^M4\s*') { $newStyle = $chapterStyle }
        elseif ($oldStyle -eq '3' -or $text -match '^(\u9898\u578b|\u68c0\u6d4b\u9898\u7b54\u6848|\u53cd\u9988\u901a\u9053|\u9898\u518c\u8c03\u6574)') { $newStyle = $topicStyle; $motherSeen = $false }
        elseif ($oldStyle -eq '4' -or $text -match '^\u6bcd\u9898\d+\s*[\uFF1A:]') { $newStyle = if ($motherSeen) { $motherGap } else { $motherNoGap }; $motherSeen = $true }
        elseif ($oldStyle -in @('29','67') -or $text -match '^(\u6bcd\u9898\s*[\uFF1A:]|\u5b50\u9898\d+\s*[\uFF1A:])') { $newStyle = $subStyle }
        elseif (-not $empty -and $text -match '^\u6bcd\u9898\u6982\u8ff0\s*[\uFF1A:]$') { $newStyle = $labelBold }
        elseif (-not $empty -and $text -match '^\u89e3\u6790\s*[\uFF1A:]\s*[\.\u3002]*$') { $newStyle = $analysisBold }
        elseif (-not $empty) { $newStyle = $bodyStyle }
        if ($null -ne $newStyle) {
            $pPr = $paragraph.SelectSingleNode('./w:pPr', $sourceDocNs)
            if ($null -eq $pPr) { $pPr = $sourceDoc.CreateElement('w','pPr',$wNs); [void]$paragraph.PrependChild($pPr) }
            if ($null -eq $styleNode) { $styleNode = $sourceDoc.CreateElement('w','pStyle',$wNs); [void]$pPr.PrependChild($styleNode) }
            Set-WVal $styleNode $newStyle
            $spacing = $paragraph.SelectSingleNode('./w:pPr/w:spacing', $sourceDocNs)
            if ($null -ne $spacing) { [void]$spacing.ParentNode.RemoveChild($spacing) }
            [void](Clear-TemplateDirectIndent $paragraph $sourceDocNs $newStyle $bodyStyle)
        }

        $isAnalysisHead = $text -match '^\s*解析\s*[\uFF1A:]'
        $isAnalysisBoundary = $text -match '^(M\d+\s*|题型|母题|子题|检测题|反馈通道|题册调整)'
        if ($isAnalysisHead) {
            [void](Clear-AnalysisHeadDirectBold $paragraph $sourceDocNs)
            $inAnalysis = $true
        } elseif ($inAnalysis) {
            if ($isAnalysisBoundary) { $inAnalysis = $false }
            else { [void](Clear-AnalysisBodyDirectBold $paragraph $sourceDocNs) }
        }
    }

    # The exercise template is a continuous two-column flow. Source documents
    # sometimes put a page break inside a heading run; deleting only empty
    # page-break paragraphs leaves those hidden breaks behind. Remove every
    # explicit page-break node while retaining the surrounding paragraph/text.
    foreach ($pageBreak in @($sourceDoc.SelectNodes('//w:br[@w:type="page"]', $sourceDocNs))) {
        [void]$pageBreak.ParentNode.RemoveChild($pageBreak)
    }
    foreach ($paragraph in @($paragraphs)) {
        if (Test-RemovableBlankParagraph $paragraph $sourceDocNs) { [void]$paragraph.ParentNode.RemoveChild($paragraph) }
    }
    $paragraphs = $sourceDoc.SelectNodes('/w:document/w:body/w:p', $sourceDocNs)
    $firstTopic = $false
    foreach ($paragraph in @($paragraphs)) {
        $text = Get-ParagraphText $paragraph $sourceDocNs
        $isTopic = $text -match '^\u9898\u578b'
        $isDetection = $text -match '^\u68c0\u6d4b\u9898\s*$'
        $isAnswer = $text -match '^\u68c0\u6d4b\u9898\u7b54\u6848'
        $isFeedback = $text -match '^\u53cd\u9988\u901a\u9053'
        if ($isTopic) {
            if ($firstTopic) { Remove-BlankParagraphsBefore $paragraph $sourceDocNs; Add-ColumnBreak $paragraph $sourceDoc $sourceDocNs }
            $firstTopic = $true
        } elseif ($isDetection -or $isAnswer -or $isFeedback) {
            Remove-BlankParagraphsBefore $paragraph $sourceDocNs
            Add-ColumnBreak $paragraph $sourceDoc $sourceDocNs
            if ($isDetection) {
                $pPr = $paragraph.SelectSingleNode('./w:pPr', $sourceDocNs)
                if ($null -eq $pPr) { $pPr = $sourceDoc.CreateElement('w','pPr',$wNs); [void]$paragraph.PrependChild($pPr) }
                $pStyle = $paragraph.SelectSingleNode('./w:pPr/w:pStyle', $sourceDocNs)
                if ($null -eq $pStyle) { $pStyle = $sourceDoc.CreateElement('w','pStyle',$wNs); [void]$pPr.PrependChild($pStyle) }
                Set-WVal $pStyle $topicStyle
            }
        }
        [void](Clear-StepDirectBold $paragraph $sourceDocNs)
    }

    $templateSect = $templateDoc.SelectSingleNode('/w:document/w:body/w:sectPr', $templateDocNs)
    $sourceSect = $sourceDoc.SelectSingleNode('/w:document/w:body/w:sectPr', $sourceDocNs)
    if ($null -eq $templateSect -or $null -eq $sourceSect) { throw 'Document section properties were not found.' }
    foreach ($tag in @('pgSz','pgMar','cols','docGrid')) {
        $from = $templateSect.SelectSingleNode("./w:$tag", $templateDocNs)
        $to = $sourceSect.SelectSingleNode("./w:$tag", $sourceDocNs)
        if ($null -ne $from) {
            $clone = $sourceDoc.ImportNode($from, $true)
            if ($null -ne $to) { [void]$sourceSect.ReplaceChild($clone, $to) } else { [void]$sourceSect.AppendChild($clone) }
        }
    }

    $stylesBytes = Convert-XmlToBytes $sourceStyles
    $documentBytes = Convert-XmlToBytes $sourceDoc
    $outputZip = [IO.Compression.ZipFile]::Open($tmpPath, [IO.Compression.ZipArchiveMode]::Create)
    foreach ($entry in $sourceZip.Entries) {
        if ($entry.FullName -eq 'word/styles.xml' -or $entry.FullName -eq 'word/document.xml') {
            $bytes = if ($entry.FullName -eq 'word/styles.xml') { $stylesBytes } else { $documentBytes }
            $destination = $outputZip.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
            $stream = $destination.Open(); try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
        } else { Copy-ZipEntry $entry $outputZip }
    }
    $outputZip.Dispose(); $outputZip = $null
    $sourceZip.Dispose(); $sourceZip = $null
    $templateZip.Dispose(); $templateZip = $null
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
    Move-Item -LiteralPath $tmpPath -Destination $outputPath -Force

    $checkZip = [IO.Compression.ZipFile]::OpenRead($outputPath)
    try {
        $entry = $checkZip.GetEntry('word/document.xml'); $reader = [IO.StreamReader]::new($entry.Open())
        try { $outputXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $checkZip.Dispose() }
    $oleCount = ([regex]::Matches($outputXml, 'Equation\.DSMT4')).Count
    Write-Output ("Saved: {0}" -f $outputPath)
    Write-Output ("Equation.DSMT4 markers retained: {0}" -f $oleCount)
} finally {
    if ($null -ne $outputZip) { $outputZip.Dispose() }
    if ($null -ne $sourceZip) { $sourceZip.Dispose() }
    if ($null -ne $templateZip) { $templateZip.Dispose() }
    if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue }
}
