param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [Parameter(Mandatory)][string]$Lines,
    [Parameter(Mandatory)][string]$Output,
    [string]$Python = 'python',
    [int[]]$Pages,
    [string]$Report
)

# Reflow only explicit, source-reviewed mixed prose/math line recipes.
# Formula font, color, AutoSize, and Office Math structure are never mutated.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'native_shape_lookup.ps1')
$Deck = (Resolve-Path -LiteralPath $Deck).Path
$Run = (Resolve-Path -LiteralPath $Run).Path
$Lines = (Resolve-Path -LiteralPath $Lines).Path
$Output = [IO.Path]::GetFullPath($Output)
if (Test-Path -LiteralPath $Output) { throw 'Use a fresh output path' }
if (-not $Report) { $Report = $Output + '.geometry.json' }
$Report = [IO.Path]::GetFullPath($Report)
if (Test-Path -LiteralPath $Report) { throw 'Use a fresh geometry report path' }
$failedPlanPath = $Output + '.preflight-plan.json'
$failedWidthsPath = $Output + '.preflight-widths.json'
foreach ($evidence in @($failedPlanPath, $failedWidthsPath)) {
    if (Test-Path -LiteralPath $evidence) { throw 'Use a fresh output path to preserve failed preflight evidence' }
}

$preflight = Join-Path $PSScriptRoot 'preflight_reflow_mixed_lines.py'
if (-not (Test-Path -LiteralPath $preflight)) { throw "Missing preflight helper: $preflight" }
$tempRoot = [IO.Path]::GetTempPath()
$planPath = Join-Path $tempRoot ('mixed-line-plan-' + [Guid]::NewGuid().ToString('N') + '.json')
$measurementPath = Join-Path $tempRoot ('mixed-line-widths-' + [Guid]::NewGuid().ToString('N') + '.json')
$app = $null
$presentation = $null
$records = [System.Collections.Generic.List[object]]::new()

function Get-SourceMap($Manifest) {
    if ($null -eq $Manifest.content_box) { throw 'Manifest has no content_box' }
    $leftPt = [double]$Manifest.content_box.left * 72.0
    $topPt = [double]$Manifest.content_box.top * 72.0
    $widthPt = [double]$Manifest.content_box.width * 72.0
    $heightPt = [double]$Manifest.content_box.height * 72.0
    if ($widthPt -le 0 -or $heightPt -le 0) { throw 'Manifest content_box is invalid' }
    return [pscustomobject]@{
        LeftPt = $leftPt
        TopPt = $topPt
        PxPerPointX = [double]$Manifest.source.width_px / $widthPt
        PxPerPointY = [double]$Manifest.source.height_px / $heightPt
    }
}

try {
    $preflightArgs = @($preflight, '--lines', $Lines, '--run', $Run, '--deck', $Deck, '--output', $planPath)
    if ($Pages) {
        foreach ($page in $Pages) { $preflightArgs += @('--page', [string]$page) }
    }
    & $Python @preflightArgs | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Mixed-line recipe preflight failed; PowerPoint was not opened' }
    $plan = Get-Content -LiteralPath $planPath -Raw -Encoding utf8 | ConvertFrom-Json

    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $app.Presentations.Open($Deck, $false, $false, $false)
    $states = [System.Collections.Generic.List[object]]::new()
    $measurements = [System.Collections.Generic.List[object]]::new()

    # Normalize prose and collect all live BoundWidth values before moving any
    # shape. Formula ranges are read only.
    foreach ($line in @($plan.lines)) {
        $page = [int]$line.page
        $slide = $presentation.Slides.Item($page)
        $manifestPath = Join-Path $Run ('pages\page_{0:D3}\manifest.json' -f $page)
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding utf8 | ConvertFrom-Json
        $map = Get-SourceMap $manifest
        $itemStates = [System.Collections.Generic.List[object]]::new()
        $measuredItems = [System.Collections.Generic.List[object]]::new()
        foreach ($item in @($line.items)) {
            $shape = Get-ExactSlideShape $slide ([string]$item.shape)
            if (-not $shape.HasTextFrame -or -not $shape.TextFrame.HasText) {
                throw "Line item has no text range: $page/$($item.shape)"
            }
            $range = $shape.TextFrame2.TextRange
            if ($null -ne $item.expected_text -and [string]$range.Text -cne [string]$item.expected_text) {
                throw "Exact expected text changed after preflight: $page/$($item.shape)"
            }
            if ([string]$item.kind -eq 'prose') {
                # ShapeToFitText changes the frame, never the requested font size.
                $shape.TextFrame2.WordWrap = 0
                $shape.TextFrame2.AutoSize = 1
                $range.Font.Size = [double]$item.prose_font_pt
            } elseif ([string]$item.kind -ne 'math') {
                throw "Unknown line item kind: $($item.kind)"
            }
            $widthPx = [double]$range.BoundWidth * [double]$map.PxPerPointX
            if ([double]::IsNaN($widthPx) -or [double]::IsInfinity($widthPx) -or $widthPx -le 0) {
                throw "Invalid live BoundWidth: $page/$($item.shape)"
            }
            $itemStates.Add([pscustomobject]@{
                Plan = $item
                Shape = $shape
                Range = $range
                WidthPx = $widthPx
            })
            $measuredItems.Add([ordered]@{shape=[string]$item.shape; width_px=$widthPx})
        }
        $states.Add([pscustomobject]@{
            Plan = $line
            Map = $map
            Items = $itemStates.ToArray()
        })
        $measurements.Add([ordered]@{id=[string]$line.id; items=$measuredItems.ToArray()})
    }

    ConvertTo-Json -InputObject @($measurements.ToArray()) -Depth 8 | Set-Content -LiteralPath $measurementPath -Encoding utf8
    & $Python $preflight --plan $planPath --measurements $measurementPath | Out-Host
    if ($LASTEXITCODE -ne 0) {
        [IO.File]::Copy($planPath, $failedPlanPath, $false)
        [IO.File]::Copy($measurementPath, $failedWidthsPath, $false)
        Write-Host "Preflight evidence: $failedPlanPath ; $failedWidthsPath"
        throw 'Measured line widths overflow or differ from the validated plan; no shapes were moved'
    }

    foreach ($state in $states) {
        $line = $state.Plan
        $map = $state.Map
        $cursor = [double]$line.start_x_px
        foreach ($itemState in @($state.Items)) {
            $item = $itemState.Plan
            $shape = $itemState.Shape
            $range = $itemState.Range
            $cursor += [double]$item.gap_before_px
            $targetLeftPt = [double]$map.LeftPt + $cursor / [double]$map.PxPerPointX
            $targetCenterPt = [double]$map.TopPt + ([double]$line.center_y_px + [double]$item.center_offset_px) / [double]$map.PxPerPointY
            $dxPt = $targetLeftPt - [double]$range.BoundLeft
            $dyPt = $targetCenterPt - ([double]$range.BoundTop + [double]$range.BoundHeight / 2.0)
            $shape.Left = [double]$shape.Left + $dxPt
            $shape.Top = [double]$shape.Top + $dyPt
            $records.Add([ordered]@{
                line_id = [string]$line.id
                page = [int]$line.page
                shape = [string]$shape.Name
                kind = [string]$item.kind
                text = [string]$range.Text
                left_px = $cursor
                width_px = [double]$itemState.WidthPx
                center_y_px = [double]$line.center_y_px + [double]$item.center_offset_px
                dx_px = $dxPt * [double]$map.PxPerPointX
                dy_px = $dyPt * [double]$map.PxPerPointY
            })
            $cursor += [double]$itemState.WidthPx
        }
        Write-Host ("Reflowed page {0} line '{1}', right={2:N1}px" -f $line.page, $line.id, $cursor)
    }

    $presentation.SaveAs($Output, 24)
    [ordered]@{
        schema_version = 1
        input = $Deck
        output = $Output
        recipe = $Lines
        line_count = [int]$plan.line_count
        records = $records.ToArray()
    } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Report -Encoding utf8
} finally {
    if ($null -ne $presentation) { $presentation.Close() }
    if ($null -ne $app -and $app.Presentations.Count -eq 0) { $app.Quit() }
    if (Test-Path -LiteralPath $planPath) { Remove-Item -LiteralPath $planPath -Force }
    if (Test-Path -LiteralPath $measurementPath) { Remove-Item -LiteralPath $measurementPath -Force }
}
