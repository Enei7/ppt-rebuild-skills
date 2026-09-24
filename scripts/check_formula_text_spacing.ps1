param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [double]$MinOverlapPt = 4,
    [double]$MaxCenterDeltaPt = 13,
    [double]$MinClearancePt = 2,
    [switch]$ReportTightGaps,
    [switch]$ReportBaseline,
    [double]$MaxBaselineDeltaPt = 5,
    [switch]$FailOnCollision
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Deck)) { throw "Deck not found: $Deck" }
if (-not (Test-Path -LiteralPath $Run)) { throw "Run not found: $Run" }
$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
$hits = New-Object 'System.Collections.Generic.List[object]'
$baselineHits = New-Object 'System.Collections.Generic.List[object]'
$collisions = 0
try {
    $presentation = $app.Presentations.Open($Deck, $true, $false, $false)
    for ($page = 1; $page -le $presentation.Slides.Count; $page++) {
        $manifest = Get-Content -LiteralPath (Join-Path $Run ('pages\page_{0:D3}\manifest.json' -f $page)) -Raw -Encoding utf8 | ConvertFrom-Json
        $formulaIds = New-Object 'System.Collections.Generic.HashSet[string]'
        $latexById = @{}
        foreach ($formula in $manifest.formula_inventory) {
            [void]$formulaIds.Add([string]$formula.id)
            $latexById[[string]$formula.id] = [string]$formula.latex
        }
        $slide = $presentation.Slides.Item($page)
        $prose = @()
        foreach ($shape in $slide.Shapes) {
            if ($formulaIds.Contains([string]$shape.Name) -or -not $shape.HasTextFrame -or -not $shape.TextFrame.HasText) { continue }
            $range = $shape.TextFrame2.TextRange
            $content = [string]$range.Text
            # ponytail: single-line bounds only; full-slide visual QA covers multiline text.
            if ([string]::IsNullOrWhiteSpace($content) -or $content.Contains("`r") -or $content.Contains("`n")) { continue }
            $prose += [pscustomobject]@{
                name = [string]$shape.Name; text = $content.Trim()
                left = [double]$range.BoundLeft
                right = [double]$range.BoundLeft + [double]$range.BoundWidth - 6
                centerY = [double]$range.BoundTop + [double]$range.BoundHeight / 2
            }
        }
        foreach ($id in $formulaIds) {
            $shape = $slide.Shapes.Item($id)
            $range = $shape.TextFrame2.TextRange
            $left = [double]$range.BoundLeft
            $right = $left + [double]$range.BoundWidth - 6
            $centerY = [double]$range.BoundTop + [double]$range.BoundHeight / 2
            foreach ($text in $prose) {
                if ([Math]::Abs($centerY - $text.centerY) -gt $MaxCenterDeltaPt) { continue }
                $overlap = [Math]::Min($right, $text.right) - [Math]::Max($left, $text.left)
                $collision = $overlap -ge $MinOverlapPt
                if (-not $collision -and -not ($ReportTightGaps -and $overlap -gt -$MinClearancePt)) { continue }
                if ($collision) { $collisions++ }
                $hits.Add([pscustomobject]@{
                    page = $page; kind = $(if ($collision) { 'collision' } else { 'tight' })
                    formula = $id; text_shape = $text.name
                    overlap_pt = [Math]::Round($overlap, 1)
                    text = $text.text.Substring(0, [Math]::Min(80, $text.text.Length))
                })
            }
            if ($ReportBaseline -and [double]$range.BoundHeight -le 28 -and $latexById[$id] -notmatch '\\(?:sum|int|prod|frac|dfrac|tfrac|lim)') {
                $nearest = $prose | ForEach-Object {
                    $gap = [Math]::Max(0, [Math]::Max($_.left - $right, $left - $_.right))
                    [pscustomobject]@{ text = $_; gap = $gap; delta = $centerY - $_.centerY; deltaAbs = [Math]::Abs($centerY - $_.centerY) }
                } | Where-Object { $_.gap -le 20 -and $_.deltaAbs -le 20 } | Sort-Object deltaAbs,gap | Select-Object -First 1
                if ($nearest -and [Math]::Abs($nearest.delta) -gt $MaxBaselineDeltaPt) {
                    $baselineHits.Add([pscustomobject]@{
                        page = $page; formula = $id; delta_pt = [Math]::Round($nearest.delta, 1)
                        text_shape = $nearest.text.name; text = $nearest.text.text.Substring(0, [Math]::Min(60, $nearest.text.text.Length))
                    })
                }
            }
        }
    }
    $hits | Sort-Object -Property @{Expression='overlap_pt';Descending=$true} | Format-Table -AutoSize | Out-String | Write-Host
    Write-Host "Formula/prose collision candidates: $collisions; tight gaps: $($hits.Count - $collisions)"
    if ($ReportBaseline) {
        $baselineHits | Sort-Object page | Format-Table -AutoSize | Out-String | Write-Host
        Write-Host "Compact inline baseline candidates: $($baselineHits.Count)"
    }
    if ($FailOnCollision -and $collisions -gt 0) { throw "$collisions formula/prose collisions require review" }
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    $app.Quit()
}
