[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$ReferenceDocx
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$docxPath = [IO.Path]::GetFullPath($Docx)
if (-not (Test-Path -LiteralPath $docxPath -PathType Leaf)) { throw "DOCX was not found: $docxPath" }
$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function Open-DocxData {
    param([string]$Path)
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if ($null -eq $entry) { throw "word/document.xml not found: $Path" }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $xml = [Xml.XmlDocument]::new(); $xml.LoadXml($text)
        $stylesEntry = $zip.GetEntry('word/styles.xml')
        if ($null -eq $stylesEntry) { throw "word/styles.xml not found: $Path" }
        $stylesReader = [IO.StreamReader]::new($stylesEntry.Open())
        try { $stylesText = $stylesReader.ReadToEnd() } finally { $stylesReader.Dispose() }
        $stylesXml = [Xml.XmlDocument]::new(); $stylesXml.LoadXml($stylesText)
        $stylesNs = [Xml.XmlNamespaceManager]::new($stylesXml.NameTable); $stylesNs.AddNamespace('w', $wNs)
        $displayEquationStyleIds = @()
        foreach ($style in @($stylesXml.SelectNodes('/w:styles/w:style', $stylesNs))) {
            $name = $style.SelectSingleNode('./w:name', $stylesNs)
            if ($null -ne $name -and $name.GetAttribute('val', $wNs) -eq 'MTDisplayEquation') {
                $displayEquationStyleIds += $style.GetAttribute('styleId', $wNs)
            }
        }
        [pscustomobject]@{
            Xml = $xml
            Text = $text
            DisplayEquationStyleIds = $displayEquationStyleIds
            EquationMarkers = ([regex]::Matches($text, 'Equation\.DSMT4')).Count
            OleParts = @($zip.Entries | Where-Object { $_.FullName -like 'word/embeddings/*' }).Count
            MediaParts = @($zip.Entries | Where-Object { $_.FullName -like 'word/media/*' }).Count
        }
    } finally { $zip.Dispose() }
}

function Get-ParagraphText {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    return ((@($Paragraph.SelectNodes('.//w:t', $Ns) | ForEach-Object { $_.InnerText }) -join '') -replace '\s+', ' ').Trim()
}

function Test-Object {
    param([Xml.XmlElement]$Node, [Xml.XmlNamespaceManager]$Ns)
    return $null -ne $Node.SelectSingleNode('.//w:object|.//w:pict|.//w:drawing', $Ns)
}

function Get-ParagraphStyleId {
    param([Xml.XmlElement]$Paragraph, [Xml.XmlNamespaceManager]$Ns)
    $style = $Paragraph.SelectSingleNode('./w:pPr/w:pStyle', $Ns)
    if ($null -eq $style) { return '' }
    return $style.GetAttribute('val', $wNs)
}

function Test-MathTypeDisplayEquation {
    param(
        [Xml.XmlElement]$Paragraph,
        [Xml.XmlNamespaceManager]$Ns,
        [string[]]$DisplayStyleIds
    )
    $styleId = Get-ParagraphStyleId $Paragraph $Ns
    if ($styleId -and $styleId -in $DisplayStyleIds) { return $true }
    if (-not [String]::IsNullOrWhiteSpace((Get-ParagraphText $Paragraph $Ns))) { return $false }
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
    if ($null -ne $alignment -and $alignment.GetAttribute('val', $wNs) -in @('center','right','both','distribute')) { return $true }
    if ([String]::IsNullOrWhiteSpace($text) -and
        $null -ne $Paragraph.SelectSingleNode('.//w:drawing[not(ancestor::w:object)]|.//w:pict[not(ancestor::w:object)]', $Ns)) {
        return $true
    }
    return $false
}

$data = Open-DocxData $docxPath
$ns = [Xml.XmlNamespaceManager]::new($data.Xml.NameTable); $ns.AddNamespace('w', $wNs)
$paragraphs = @($data.Xml.SelectNodes('//w:body//w:p', $ns))
$breaks = @($data.Xml.SelectNodes('//w:body//w:br', $ns))
$columnBreaks = @($breaks | Where-Object { $_.GetAttribute('type', $wNs) -eq 'column' }).Count
$pageBreaks = @($breaks | Where-Object { $_.GetAttribute('type', $wNs) -eq 'page' }).Count
$pageBreakBefore = @($data.Xml.SelectNodes('//w:body//w:pageBreakBefore', $ns) | Where-Object {
    $_.GetAttribute('val', $wNs) -ne '0'
}).Count
$topics = @($paragraphs | Where-Object { (Get-ParagraphText $_ $ns) -match '^\u9898\u578b' }).Count
$detections = @($paragraphs | Where-Object { (Get-ParagraphText $_ $ns) -match '^\u68C0\u6D4B\u9898\s*$' }).Count
$answers = @($paragraphs | Where-Object { (Get-ParagraphText $_ $ns) -match '^\u68C0\u6D4B\u9898\u7B54\u6848' }).Count
$feedback = @($paragraphs | Where-Object { (Get-ParagraphText $_ $ns) -match '^\u53CD\u9988\u901A\u9053' }).Count
$displayEquationParagraphs = @($paragraphs | Where-Object {
    Test-MathTypeDisplayEquation $_ $ns $data.DisplayEquationStyleIds
})
$displayEquationLeadingTabs = @($displayEquationParagraphs | Where-Object {
    $null -ne $_.SelectSingleNode('.//w:tab', $ns)
})
$displayEquationNonZeroSpacing = @($displayEquationParagraphs | Where-Object {
    $spacing = $_.SelectSingleNode('./w:pPr/w:spacing', $ns)
    $null -eq $spacing -or
    $spacing.GetAttribute('before', $wNs) -notin @('', '0') -or
    $spacing.GetAttribute('after', $wNs) -notin @('', '0')
})
$displayEquationNotLeftAligned = @($displayEquationParagraphs | Where-Object {
    $alignment = $_.SelectSingleNode('./w:pPr/w:jc', $ns)
    $null -eq $alignment -or $alignment.GetAttribute('val', $wNs) -ne 'left'
})
$formulaOnlyBodyLeadingTabs = @($paragraphs | Where-Object {
    (Get-ParagraphStyleId $_ $ns) -eq 'm1_left_body' -and
    [String]::IsNullOrWhiteSpace((Get-ParagraphText $_ $ns)) -and
    $null -ne $_.SelectSingleNode('./w:r[1]/w:tab', $ns) -and
    $_.OuterXml -match 'Equation\.DSMT4'
})
$unexpectedBodyIndents = @($paragraphs | Where-Object {
    (Get-ParagraphStyleId $_ $ns) -eq 'm1_left_body' -and
    $null -ne $_.SelectSingleNode('./w:pPr/w:ind', $ns) -and
    -not (Test-MathTypeDisplayEquation $_ $ns $data.DisplayEquationStyleIds) -and
    -not (Test-SpecialBodyGeometry $_ $ns)
})
$removableBlankParagraphs = @($paragraphs | Where-Object {
    [String]::IsNullOrWhiteSpace((Get-ParagraphText $_ $ns)) -and
    $null -eq $_.SelectSingleNode('.//w:object|.//w:pict|.//w:drawing|.//w:br|.//w:bookmarkStart|.//w:bookmarkEnd|./w:pPr/w:sectPr|.//w:fldChar|.//w:instrText', $ns)
})

$stepParagraphs = @()
$boldStepParagraphs = @()
foreach ($paragraph in $paragraphs) {
    $runs = @($paragraph.SelectNodes('./w:r', $ns))
    if ($runs.Count -lt 2 -or -not (Test-Object $runs[0] $ns)) { continue }
    $firstText = $null
    foreach ($run in $runs) {
        $runText = ((@($run.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''))
        if ($runText) { $firstText = $runText; break }
    }
    if ($null -eq $firstText -or -not $firstText.StartsWith([char]0xFF1A)) { continue }
    $stepParagraphs += $paragraph
    $boldCount = 0
    foreach ($run in $runs) { $boldCount += @($run.SelectNodes('./w:rPr/w:b|./w:rPr/w:bCs', $ns)).Count }
    if ($boldCount -gt 0) { $boldStepParagraphs += $paragraph }
}

$expectedColumnBreaks = [Math]::Max(0, $topics - 1) + $detections + $answers + $feedback
$checks = [ordered]@{
    'DOCX exists' = $true
    'Equation marker count' = ($data.EquationMarkers -gt 0)
    'No page breaks' = ($pageBreaks -eq 0 -and $pageBreakBefore -eq 0)
    'Expected column break count' = ($columnBreaks -eq $expectedColumnBreaks)
    'Step colon and suffix not directly bold' = ($boldStepParagraphs.Count -eq 0)
    'No formula-only body paragraph begins with a tab' = ($formulaOnlyBodyLeadingTabs.Count -eq 0)
    'No unexpected direct indent on template body paragraphs' = ($unexpectedBodyIndents.Count -eq 0)
    'Display equations contain no positioning tabs' = ($displayEquationLeadingTabs.Count -eq 0)
    'Display equations have zero before and after spacing' = ($displayEquationNonZeroSpacing.Count -eq 0)
    'Display equations are explicitly left aligned' = ($displayEquationNotLeftAligned.Count -eq 0)
    'No removable blank paragraphs remain' = ($removableBlankParagraphs.Count -eq 0)
}

if ($ReferenceDocx) {
    $referencePath = [IO.Path]::GetFullPath($ReferenceDocx)
    if (-not (Test-Path -LiteralPath $referencePath -PathType Leaf)) { throw "Reference DOCX was not found: $referencePath" }
    $reference = Open-DocxData $referencePath
    $checks['Equation markers match reference'] = ($data.EquationMarkers -eq $reference.EquationMarkers)
    $checks['OLE parts match reference'] = ($data.OleParts -eq $reference.OleParts)
    $checks['Media parts match reference'] = ($data.MediaParts -eq $reference.MediaParts)
    $referenceNs = [Xml.XmlNamespaceManager]::new($reference.Xml.NameTable); $referenceNs.AddNamespace('w', $wNs)
    $referenceParagraphs = @($reference.Xml.SelectNodes('//w:body//w:p', $referenceNs))
    $referenceDisplayEquationParagraphs = @($referenceParagraphs | Where-Object {
        Test-MathTypeDisplayEquation $_ $referenceNs $reference.DisplayEquationStyleIds
    })
    $checks['Display equation paragraphs match reference'] = ($displayEquationParagraphs.Count -eq $referenceDisplayEquationParagraphs.Count)
}

$failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
[pscustomobject]@{
    Docx = $docxPath
    Paragraphs = $paragraphs.Count
    Topics = $topics
    DetectionHeadings = $detections
    DetectionAnswerHeadings = $answers
    FeedbackHeadings = $feedback
    StepParagraphs = $stepParagraphs.Count
    StepParagraphsWithDirectBold = $boldStepParagraphs.Count
    DisplayEquationParagraphs = $displayEquationParagraphs.Count
    DisplayEquationParagraphsWithTabs = $displayEquationLeadingTabs.Count
    DisplayEquationParagraphsWithNonZeroSpacing = $displayEquationNonZeroSpacing.Count
    DisplayEquationParagraphsNotLeftAligned = $displayEquationNotLeftAligned.Count
    FormulaOnlyBodyParagraphsWithLeadingTab = $formulaOnlyBodyLeadingTabs.Count
    UnexpectedTemplateBodyDirectIndents = $unexpectedBodyIndents.Count
    RemovableBlankParagraphs = $removableBlankParagraphs.Count
    ColumnBreaks = $columnBreaks
    ExpectedColumnBreaks = $expectedColumnBreaks
    PageBreaks = $pageBreaks
    PageBreakBefore = $pageBreakBefore
    EquationMarkers = $data.EquationMarkers
    OLEParts = $data.OleParts
    MediaParts = $data.MediaParts
    Checks = (($checks.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ')
    Status = if ($failed.Count -eq 0) { 'PASS' } else { 'FAIL' }
} | Format-List

if ($failed.Count -gt 0) { exit 1 }
