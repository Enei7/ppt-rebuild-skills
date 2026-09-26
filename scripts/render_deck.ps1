param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Run,
    [Parameter(Mandatory)][string]$Folder
)

# Read-only export. Run serially with other desktop PowerPoint automation.
$ErrorActionPreference = 'Stop'
$Deck = (Resolve-Path -LiteralPath $Deck).Path
$Run = (Resolve-Path -LiteralPath $Run).Path
$Folder = [IO.Path]::GetFullPath($Folder)
if (Test-Path -LiteralPath $Folder) { throw 'Use a new render folder to preserve earlier evidence' }
$metadata = Get-Content -LiteralPath (Join-Path $Run 'deck_manifest.json') -Raw -Encoding utf8 | ConvertFrom-Json
$pages = @($metadata.pages)
New-Item -ItemType Directory -Path $Folder | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $app.Presentations.Open($Deck, $true, $false, $false)
    if ($presentation.Slides.Count -ne $pages.Count) { throw 'Deck/run page counts differ' }
    for ($index = 0; $index -lt $pages.Count; $index++) {
        $page = $pages[$index]
        $manifest = Get-Content -LiteralPath (Join-Path $Run $page.manifest) -Raw -Encoding utf8 | ConvertFrom-Json
        $width = [int]$manifest.source.width_px
        $height = [int]$manifest.source.height_px
        if ($width -le 0 -or $height -le 0) { throw "Invalid source dimensions on page $($index + 1)" }
        $presentation.Slides.Item($index + 1).Export((Join-Path $Folder ('page_{0:D3}.png' -f ($index + 1))), 'PNG', $width, $height)
        Write-Output "Rendered $($index + 1)/$($pages.Count)"
    }
} finally {
    if ($null -ne $presentation) { $presentation.Close() }
    # Do not close unrelated presentations in the shared desktop application.
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
