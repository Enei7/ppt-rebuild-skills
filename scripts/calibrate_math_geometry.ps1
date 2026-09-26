param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [Parameter(Mandatory)][string]$Output,
    [string]$Python = 'python',
    [double]$MaxFont = 44,
    [double]$MaxShiftPx = 200
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'native_shape_lookup.ps1')
if (-not (Test-Path -LiteralPath $Deck)) { throw "Deck not found: $Deck" }
if (-not (Test-Path -LiteralPath $Run)) { throw "Run not found: $Run" }
if ([IO.Path]::GetFullPath($Deck) -eq [IO.Path]::GetFullPath($Output)) { throw 'Output must differ from input' }
if (Test-Path -LiteralPath $Output) { throw "Output already exists: $Output" }
$measure = Join-Path $PSScriptRoot 'measure_math_geometry.py'
if (-not (Test-Path -LiteralPath $measure)) { throw "Missing helper: $measure" }
$work = Join-Path $Run ('math_calibration_' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $work -ErrorAction Stop | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$presentation = $null

function Export-FormulaOnly([string]$Phase) {
    $folder = Join-Path $work $Phase
    New-Item -ItemType Directory -Path $folder -ErrorAction Stop | Out-Null
    for ($page = 1; $page -le $presentation.Slides.Count; $page++) {
        $manifest = Get-Content -LiteralPath (Join-Path $Run ('pages\page_{0:D3}\manifest.json' -f $page)) -Raw -Encoding utf8 | ConvertFrom-Json
        $ids = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($formula in $manifest.formula_inventory) { [void]$ids.Add([string]$formula.id) }
        $slide = $presentation.Slides.Item($page)
        $hidden = New-Object 'System.Collections.Generic.List[object]'
        try {
            foreach ($shape in $slide.Shapes) {
                if ($ids.Contains([string]$shape.Name)) { continue }
                $hidden.Add(@($shape, $shape.Visible))
                $shape.Visible = 0
            }
            $slide.Export((Join-Path $folder ('page_{0:D3}.png' -f $page)), 'PNG', [int]$manifest.source.width_px, [int]$manifest.source.height_px)
        } finally {
            foreach ($entry in $hidden) { $entry[0].Visible = $entry[1] }
        }
        if ($page % 10 -eq 0) { Write-Host "$Phase rendered $page/$($presentation.Slides.Count)" }
    }
    $json = Join-Path $work ($Phase + '.json')
    & $Python $measure --run $Run --renders $folder --output $json | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Geometry measurement failed: $Phase" }
    foreach ($entry in (Get-Content -LiteralPath $json -Raw -Encoding utf8 | ConvertFrom-Json)) { Write-Output $entry }
}

try {
    $presentation = $app.Presentations.Open($Deck, $false, $false, $false)
    $scaled = 0
    $shifted = 0
    $targeted = New-Object 'System.Collections.Generic.HashSet[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $sourceJson = Join-Path $work 'source.json'
    & $Python $measure --run $Run --output $sourceJson | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Source geometry measurement failed' }
    $sourceRows = Get-Content -LiteralPath $sourceJson -Raw -Encoding utf8 | ConvertFrom-Json
    $fitted = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in $sourceRows) {
        $key = "$($row.page)/$($row.id)"
        $shape = Get-ExactSlideShape ($presentation.Slides.Item([int]$row.page)) ([string]$row.id)
        $range = $shape.TextFrame2.TextRange
        $pxPerPointX = [double]$row.source_size_px[0] / [double]$presentation.PageSetup.SlideWidth
        $pxPerPointY = [double]$row.source_size_px[1] / [double]$presentation.PageSetup.SlideHeight
        $boundWidth = [double]$range.BoundWidth * $pxPerPointX
        $boundHeight = [double]$range.BoundHeight * $pxPerPointY
        if ($boundWidth -le 0 -or $boundHeight -le 0) { $warnings.Add("${key}: empty native equation bounds"); continue }
        $ratio = [Math]::Min(1.0, [Math]::Min([double]$row.box_px[2] / $boundWidth, [double]$row.box_px[3] / $boundHeight))
        if ($ratio -ge 0.95) { continue }
        [void]$fitted.Add($key)
        [void]$targeted.Add($key)
        $oldSize = [double]$range.Font.Size
        $range.Font.Size = [double]([Math]::Round([Math]::Max(8, $oldSize * $ratio) * 2) / 2)
        if ($range.Font.Size -lt $oldSize) { $scaled++ }
        $sourceCenterX = ([double]$row.box_px[0] + [double]$row.box_px[2] / 2) / $pxPerPointX
        $sourceCenterY = ([double]$row.box_px[1] + [double]$row.box_px[3] / 2) / $pxPerPointY
        $dx = $sourceCenterX - ([double]$range.BoundLeft + [double]$range.BoundWidth / 2)
        $dy = $sourceCenterY - ([double]$range.BoundTop + [double]$range.BoundHeight / 2)
        $shape.Left = [double]$shape.Left + $dx
        $shape.Top = [double]$shape.Top + $dy
        if ([Math]::Abs($dx * $pxPerPointX) -ge 1 -or [Math]::Abs($dy * $pxPerPointY) -ge 1) { $shifted++ }
    }
    foreach ($row in $sourceRows) {
        if ($fitted.Contains("$($row.page)/$($row.id)")) { continue }
        $tex = [string]$row.latex
        $operator = $tex.Contains('\sum') -or $tex.Contains('\int') -or $tex.Contains('\prod')
        $wideDisplay = [int]$row.box_px[2] -ge 800 -and [int]$row.box_px[3] -ge 90 -and -not $tex.Contains('\frac')
        if ($tex.Contains('\frac') -or $tex.Contains('\dfrac') -or $tex.Contains('\tfrac')) { continue }
        if (-not ($operator -or $wideDisplay)) { continue }
        [void]$targeted.Add("$($row.page)/$($row.id)")
        if (-not $row.source_ink) {
            $warnings.Add("$($row.page)/$($row.id): no source ink in formula box")
            continue
        }
        $sourceBox = if ($null -ne $row.source_box_px) { $row.source_box_px } else { $row.box_px }
        $sourceInkRaw = if ($null -ne $row.source_ink_in_source_box) { $row.source_ink_in_source_box } else { $row.source_ink }
        $w = [double]$sourceBox[2]; $h = [double]$sourceBox[3]
        $touchesEdge = [double]$sourceInkRaw[0] -le 2 -or [double]$sourceInkRaw[1] -le 2 -or [double]$sourceInkRaw[2] -ge $w - 2 -or [double]$sourceInkRaw[3] -ge $h - 2
        $inkWidth = [double]$sourceInkRaw[2] - [double]$sourceInkRaw[0]
        $inkHeight = [double]$sourceInkRaw[3] - [double]$sourceInkRaw[1]
        if ($touchesEdge -and ($inkWidth -lt $w * 0.55 -or $inkHeight -lt $h * 0.55)) {
            $warnings.Add("$($row.page)/$($row.id): source crop clips formula ink")
            continue
        }
        $shape = Get-ExactSlideShape ($presentation.Slides.Item([int]$row.page)) ([string]$row.id)
        $range = $shape.TextFrame2.TextRange
        $pxPerPointX = [double]$row.source_size_px[0] / [double]$presentation.PageSetup.SlideWidth
        $pxPerPointY = [double]$row.source_size_px[1] / [double]$presentation.PageSetup.SlideHeight
        $sourceWidth = [double]($row.source_ink[2] - $row.source_ink[0])
        # Office's BoundWidth includes approximately 6 pt of text-frame padding.
        $visibleWidth = [Math]::Max(1, ([double]$range.BoundWidth - 6.0) * $pxPerPointX)
        $ratio = $sourceWidth / $visibleWidth
        if ($ratio -lt 0.65 -or $ratio -gt 3.2) {
            $warnings.Add("$($row.page)/$($row.id): implausible width ratio $ratio")
            continue
        }
        $oldSize = [double]$range.Font.Size
        $newSize = [Math]::Round([Math]::Min($MaxFont, [Math]::Max(8, $oldSize * $ratio)) * 2) / 2
        if ([Math]::Abs($newSize - $oldSize) -ge 0.25) {
            try { $range.Font.Size = [double]$newSize } catch { throw "Cannot size $($row.page)/$($row.id) from $oldSize to $newSize (ratio $ratio): $($_.Exception.Message)" }
            $scaled++
        }
        $sourceCenterX = [double]$row.box_px[0] + ([double]$row.source_ink[0] + [double]$row.source_ink[2]) / 2
        $sourceCenterY = [double]$row.box_px[1] + ([double]$row.source_ink[1] + [double]$row.source_ink[3]) / 2
        $inkCenterX = ([double]$range.BoundLeft + ([double]$range.BoundWidth - 6.0) / 2) * $pxPerPointX
        $inkCenterY = ([double]$range.BoundTop + [double]$range.BoundHeight / 2 + 1.6) * $pxPerPointY
        $dx = $sourceCenterX - $inkCenterX
        $dy = $sourceCenterY - $inkCenterY
        if ([Math]::Abs($dx) -gt $MaxShiftPx -or [Math]::Abs($dy) -gt $MaxShiftPx) {
            $warnings.Add("$($row.page)/$($row.id): shift exceeds $MaxShiftPx px ($dx,$dy)")
            continue
        }
        $newLeft = [double]$shape.Left + $dx / $pxPerPointX
        $newTop = [double]$shape.Top + $dy / $pxPerPointY
        # Math ink can be inside the page even when its large text frame extends past an edge.
        $shape.Left = $newLeft
        $shape.Top = $newTop
        if ([Math]::Abs($dx) -ge 1 -or [Math]::Abs($dy) -ge 1) { $shifted++ }
    }
    $presentation.SaveAs($Output, 24)
    & $Python (Join-Path $PSScriptRoot 'equationize.py') --audit-existing --input $Output | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Saved PowerPoint failed native-equation/XML audit; do not deliver it' }
    $final = Export-FormulaOnly 'final'
    $outliers = @($final | Where-Object {
        $targeted.Contains("$($_.page)/$($_.id)") -and (
        -not $_.source_ink -or -not $_.rendered_ink -or
        [Math]::Abs([double]$_.dx_px) -gt 15 -or [Math]::Abs([double]$_.dy_px) -gt 15 -or
        [double]$_.width_ratio -lt 0.8 -or [double]$_.width_ratio -gt 1.25
        )
    })
    $report = [ordered]@{ input=$Deck; output=$Output; scaled=$scaled; shifted=$shifted; equations=$final.Count; targeted=$targeted.Count; outliers=$outliers.Count; warnings=$warnings.ToArray(); outlier_formulas=$outliers }
    $reportPath = Join-Path $work 'report.json'
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding utf8
    'equations={0} targeted={1} scaled={2} shifted={3} outliers={4} warnings={5} report={6}' -f $final.Count, $targeted.Count, $scaled, $shifted, $outliers.Count, $warnings.Count, $reportPath
} catch {
    Write-Error ("Calibration failed at line {0}, page {1}, formula {2}: {3}" -f $_.InvocationInfo.ScriptLineNumber, $row.page, $row.id, $_.Exception.Message)
    throw
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
