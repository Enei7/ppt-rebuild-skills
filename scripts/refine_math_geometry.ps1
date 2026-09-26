param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [Parameter(Mandatory)][string]$Targets,
    [Parameter(Mandatory)][string]$Output,
    [string]$Python = 'python',
    [int]$FirstPage = 1,
    [switch]$AlignOnly,
    [double]$MaxShiftPx = 200,
    [double]$MinScale = 0.5,
    [double]$MaxScale = 2.5
)

# Selected, visually reviewed outliers only. All COM work must run serially.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'native_shape_lookup.ps1')
$Deck = (Resolve-Path -LiteralPath $Deck).Path
$Run = (Resolve-Path -LiteralPath $Run).Path
$Output = [IO.Path]::GetFullPath($Output)
if (Test-Path -LiteralPath $Output) { throw 'Use a new output path' }
if ($FirstPage -lt 1 -or $MaxShiftPx -le 0 -or $MinScale -le 0 -or $MaxScale -lt $MinScale) { throw 'Invalid refinement limits' }
$targetsData = @(Get-Content -LiteralPath $Targets -Raw -Encoding utf8 | ConvertFrom-Json)
if ($targetsData.Count -eq 0) { throw 'Select at least one source-verified formula' }
$work = Join-Path $Run ('math_refinement_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null
$helper = Join-Path $PSScriptRoot 'measure_isolated_math.py'
$app = New-Object -ComObject PowerPoint.Application
$presentation = $null

function Export-Isolated([string]$Phase) {
    $folder = Join-Path $work $Phase
    New-Item -ItemType Directory -Path $folder | Out-Null
    ConvertTo-Json -InputObject $targetsData -Depth 5 | Set-Content -LiteralPath (Join-Path $folder 'targets.json') -Encoding utf8
    for ($index = 0; $index -lt $targetsData.Count; $index++) {
        $target = $targetsData[$index]
        $slide = $presentation.Slides.Item([int]$target.page - $FirstPage + 1)
        $manifestPath = Join-Path $Run ('pages\page_{0:D3}\manifest.json' -f [int]$target.page)
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding utf8 | ConvertFrom-Json
        $visible = @()
        try {
            foreach ($shape in $slide.Shapes) {
                $visible += ,@($shape, $shape.Visible)
                $shape.Visible = 0
            }
            (Get-ExactSlideShape $slide ([string]$target.id)).Visible = -1
            $slide.Export((Join-Path $folder ('formula_{0:D4}.png' -f $index)), 'PNG', [int]$manifest.source.width_px, [int]$manifest.source.height_px)
        } finally {
            foreach ($entry in $visible) { $entry[0].Visible = $entry[1] }
        }
    }
    $result = & $Python $helper --run $Run --renders $folder
    if ($LASTEXITCODE -ne 0) { throw "Isolated measurement failed: $Phase" }
    $result | Set-Content -LiteralPath (Join-Path $folder 'measurement.json') -Encoding utf8
    $result | ConvertFrom-Json
}

try {
    $presentation = $app.Presentations.Open($Deck, $false, $false, $false)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($target in $targetsData) {
        if (-not $seen.Add("$($target.page)/$($target.id)")) { throw 'Duplicate refinement target' }
        $slideIndex = [int]$target.page - $FirstPage + 1
        if ($slideIndex -lt 1 -or $slideIndex -gt $presentation.Slides.Count) { throw 'Target page is outside the deck' }
        $shape = Get-ExactSlideShape ($presentation.Slides.Item($slideIndex)) ([string]$target.id)
        if (-not $shape.HasTextFrame) { throw 'Selected formula is not a native text/math shape' }
    }
    $before = @(Export-Isolated 'before')
    if (-not $AlignOnly) {
        foreach ($row in $before) {
            if ($row.ratio -lt $MinScale -or $row.ratio -gt $MaxScale) { throw "Implausible scale for $($row.page)/$($row.id): $($row.ratio)" }
            $range = (Get-ExactSlideShape ($presentation.Slides.Item([int]$row.page - $FirstPage + 1)) ([string]$row.id)).TextFrame2.TextRange
            $size = [double]$range.Font.Size
            if ($size -le 0) { throw 'Mixed/invalid font sizes require manual review' }
            $newSize = [Math]::Round($size * [double]$row.ratio * 2) / 2
            if ($newSize -lt 4 -or $newSize -gt 96) { throw "Unsafe candidate point size: $newSize" }
            $range.Font.Size = [double]$newSize
        }
    }
    $scaled = if ($AlignOnly) { $before } else { @(Export-Isolated 'scaled') }
    foreach ($row in $scaled) {
        if ([Math]::Abs([double]$row.dx_px) -gt $MaxShiftPx -or [Math]::Abs([double]$row.dy_px) -gt $MaxShiftPx) { throw "Large shift needs source review: $($row.page)/$($row.id)" }
        $shape = Get-ExactSlideShape ($presentation.Slides.Item([int]$row.page - $FirstPage + 1)) ([string]$row.id)
        $pxX = [double]$row.source_size_px[0] / [double]$presentation.PageSetup.SlideWidth
        $pxY = [double]$row.source_size_px[1] / [double]$presentation.PageSetup.SlideHeight
        $shape.Left = [double]$shape.Left + [double]$row.dx_px / $pxX
        $shape.Top = [double]$shape.Top + [double]$row.dy_px / $pxY
    }
    $after = @(Export-Isolated 'after')
    $presentation.SaveAs($Output, 24)
    & $Python (Join-Path $PSScriptRoot 'equationize.py') --audit-existing --input $Output | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Saved refinement failed native-equation/XML audit' }
    [ordered]@{ input=$Deck; output=$Output; targets=$targetsData.Count; align_only=[bool]$AlignOnly; measurements=$after; visual_review_required=$true } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $work 'report.json') -Encoding utf8
    Write-Output "Refined $($targetsData.Count) formulas; evidence=$work; output=$Output"
} finally {
    if ($null -ne $presentation) { $presentation.Close() }
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
