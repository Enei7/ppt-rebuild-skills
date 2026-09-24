param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [Parameter(Mandatory)][string]$Output,
    [int]$FirstPage = 1,
    [double]$MaxExpand = 1.65
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Deck)) { throw "Deck not found: $Deck" }
if (-not (Test-Path -LiteralPath $Run)) { throw "Run not found: $Run" }
if ([IO.Path]::GetFullPath($Deck) -eq [IO.Path]::GetFullPath($Output)) { throw 'Output must differ from input' }
if (Test-Path -LiteralPath $Output) { throw "Output already exists: $Output" }
if ($FirstPage -lt 1 -or $MaxExpand -lt 1) { throw 'FirstPage must be positive and MaxExpand must be at least 1' }

function Median([double[]]$Values) {
    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) { return 0.0 }
    $mid = [int][Math]::Floor($sorted.Count / 2)
    if ($sorted.Count % 2) { return [double]$sorted[$mid] }
    return ([double]$sorted[$mid - 1] + [double]$sorted[$mid]) / 2
}

$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
$changed = 0
$expanded = 0
$belowTarget = 0
try {
    $presentation = $app.Presentations.Open($Deck, $false, $false, $false)
    for ($slideNo = 1; $slideNo -le $presentation.Slides.Count; $slideNo++) {
        $page = $FirstPage + $slideNo - 1
        $slide = $presentation.Slides.Item($slideNo)
        $manifestPath = Join-Path $Run ('pages\page_{0:D3}\manifest.json' -f $page)
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding utf8 | ConvertFrom-Json
        $ids = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($formula in $manifest.formula_inventory) { [void]$ids.Add([string]$formula.id) }
        $bodySizes = New-Object 'System.Collections.Generic.List[double]'
        $formulaSizes = New-Object 'System.Collections.Generic.List[double]'
        foreach ($shape in $slide.Shapes) {
            if (-not $shape.HasTextFrame -or -not $shape.TextFrame2.HasText) { continue }
            $size = [double]$shape.TextFrame2.TextRange.Font.Size
            if ($ids.Contains([string]$shape.Name)) { $formulaSizes.Add($size) }
            elseif ($size -ge 12 -and $size -le 25 -and $shape.Top -lt $presentation.PageSetup.SlideHeight - 25 -and $shape.Width -ge 20) {
                $bodySizes.Add($size)
            }
        }
        $target = Median $bodySizes.ToArray()
        if ($target -eq 0) { $target = Median $formulaSizes.ToArray() }
        if ($target -eq 0) { continue }
        $target = [Math]::Round($target * 2) / 2

        foreach ($formula in $manifest.formula_inventory) {
            try { $shape = $slide.Shapes.Item([string]$formula.id) } catch { continue }
            $range = $shape.TextFrame2.TextRange
            $oldSize = [double]$range.Font.Size
            if ([Math]::Abs($oldSize - $target) -lt 0.25 -and [double]$range.BoundWidth -le [double]$shape.Width * 0.98) { continue }
            $oldLeft = [double]$shape.Left
            $oldWidth = [double]$shape.Width
            $range.Font.Size = $target
            $needed = [double]$range.BoundWidth / 0.95
            if ($needed -gt $oldWidth) {
                $leftLimit = 8.0
                $rightLimit = [double]$presentation.PageSetup.SlideWidth - 8
                $center = $oldLeft + $oldWidth / 2
                foreach ($other in $slide.Shapes) {
                    if ($other.Id -eq $shape.Id) { continue }
                    if (-not (($other.HasTextFrame -and $other.TextFrame2.HasText) -or $other.Type -eq 13)) { continue }
                    $overlapY = [Math]::Min($shape.Top + $shape.Height, $other.Top + $other.Height) - [Math]::Max($shape.Top, $other.Top)
                    if ($overlapY -le 2) { continue }
                    if ($other.Left + $other.Width -le $center) {
                        $leftLimit = [Math]::Max($leftLimit, [double]($other.Left + $other.Width + 3))
                    } elseif ($other.Left -ge $center) {
                        $rightLimit = [Math]::Min($rightLimit, [double]($other.Left - 3))
                    }
                }
                $available = [Math]::Max($oldWidth, [Math]::Min($oldWidth * $MaxExpand, $rightLimit - $leftLimit))
                if ($available -gt $oldWidth + 0.5) {
                    $newWidth = [Math]::Min($needed, $available)
                    $newLeft = [Math]::Max($leftLimit, [Math]::Min($oldLeft - ($newWidth - $oldWidth) / 2, $rightLimit - $newWidth))
                    $shape.Left = $newLeft
                    $shape.Width = $newWidth
                    $expanded++
                }
            }
            for ($attempt = 0; $attempt -lt 5; $attempt++) {
                $bound = [double]$range.BoundWidth
                if ($bound -le [double]$shape.Width * 0.98) { break }
                $size = [double]$range.Font.Size
                if ($size -le 8) { break }
                $nextSize = [Math]::Max(8, [Math]::Floor($size * [double]$shape.Width * 0.95 / $bound))
                if ($nextSize -ge $size) { $nextSize = [Math]::Max(8, [Math]::Floor($size) - 1) }
                $range.Font.Size = [double]$nextSize
            }
            if ([double]$range.BoundWidth -gt [double]$shape.Width * 1.01) { throw "Equation does not fit: page $page / $($formula.id)" }
            if ([Math]::Abs([double]$range.Font.Size - $oldSize) -ge 0.25) { $changed++ }
            if ([double]$range.Font.Size -lt $target - 1) { $belowTarget++ }
        }
    }
    $presentation.SaveAs($Output, 24)
    'slides={0} changed={1} expanded={2} below_target={3}' -f $presentation.Slides.Count, $changed, $expanded, $belowTarget
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    $app.Quit()
}
