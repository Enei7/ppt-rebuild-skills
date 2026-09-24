param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [double]$MaxMeanDelta = 1.0
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$nativeEquations = 0
$emptyNary = 0
$emptySlides = @()
$zip = [IO.Compression.ZipFile]::OpenRead($Deck)
try {
    foreach ($entry in $zip.Entries) {
        if ($entry.FullName -notmatch '^ppt/slides/slide\d+\.xml$') { continue }
        $stream = $entry.Open()
        try {
            $xml = New-Object System.Xml.XmlDocument
            $xml.Load($stream)
        } finally { $stream.Dispose() }
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
        $nativeEquations += $xml.SelectNodes('//m:oMath', $ns).Count
        $count = $xml.SelectNodes('//m:nary/m:e[not(*) and not(normalize-space())]', $ns).Count
        if ($count) {
            $emptyNary += $count
            $emptySlides += [int]([regex]::Match($entry.FullName, 'slide(\d+)\.xml$').Groups[1].Value)
        }
    }
} finally { $zip.Dispose() }
if ($emptyNary) { throw "Empty Office sum/integral bodies: $emptyNary on slides $($emptySlides -join ', ')" }

function Median([double[]]$Values) {
    $sorted = @($Values | Sort-Object)
    $mid = [int][Math]::Floor($sorted.Count / 2)
    if ($sorted.Count % 2) { return [double]$sorted[$mid] }
    return ([double]$sorted[$mid - 1] + [double]$sorted[$mid]) / 2
}

$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
$expected = 0
$found = 0
$compared = 0
$sumDelta = 0.0
$overflow = @()
$slideCount = 0
try {
    $presentation = $app.Presentations.Open($Deck, $true, $false, $false)
    $slideCount = $presentation.Slides.Count
    for ($page = 1; $page -le $presentation.Slides.Count; $page++) {
        $manifest = Get-Content -LiteralPath (Join-Path $Run ('pages\page_{0:D3}\manifest.json' -f $page)) -Raw -Encoding utf8 | ConvertFrom-Json
        $ids = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($formula in $manifest.formula_inventory) { [void]$ids.Add([string]$formula.id); $expected++ }
        $slide = $presentation.Slides.Item($page)
        $body = New-Object 'System.Collections.Generic.List[double]'
        foreach ($shape in $slide.Shapes) {
            if (-not $shape.HasTextFrame -or -not $shape.TextFrame2.HasText -or $ids.Contains([string]$shape.Name)) { continue }
            $size = [double]$shape.TextFrame2.TextRange.Font.Size
            if ($size -ge 12 -and $size -le 25 -and $shape.Top -lt $presentation.PageSetup.SlideHeight - 25 -and $shape.Width -ge 20) {
                $body.Add($size)
            }
        }
        $target = if ($body.Count) { Median $body.ToArray() } else { 0.0 }
        foreach ($formula in $manifest.formula_inventory) {
            try { $shape = $slide.Shapes.Item([string]$formula.id) } catch { throw "Missing equation: page $page / $($formula.id)" }
            $found++
            $range = $shape.TextFrame2.TextRange
            if ([double]$range.BoundWidth -gt [double]$shape.Width * 1.01) { $overflow += "$page/$($formula.id)" }
            if ($target -gt 0) { $sumDelta += [Math]::Abs([double]$range.Font.Size - $target); $compared++ }
        }
    }
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    $app.Quit()
}

$meanDelta = $sumDelta / [Math]::Max(1, $compared)
'slides={0} equations={1}/{2} native={3} empty_nary=0 mean_delta={4:N2}pt overflow={5}' -f $slideCount, $found, $expected, $nativeEquations, $meanDelta, $overflow.Count
if ($found -ne $expected) { throw 'Equation count mismatch' }
if ($nativeEquations -lt $expected) { throw 'Native Office equation count below formula inventory' }
if ($overflow.Count) { throw ('Equation overflow: ' + ($overflow -join ', ')) }
if ($meanDelta -gt $MaxMeanDelta) { throw "Mean font-size delta $meanDelta exceeds $MaxMeanDelta" }
