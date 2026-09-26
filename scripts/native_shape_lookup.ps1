# PowerPoint Shapes.Item(string) resolves names case-insensitively. Formula
# identities are case-sensitive: inline_G and inline_g must remain distinct.
function Get-ExactSlideShape {
    param([Parameter(Mandatory)]$Slide, [Parameter(Mandatory)][string]$Name)
    $matches = @($Slide.Shapes | Where-Object { [string]$_.Name -ceq $Name })
    if ($matches.Count -ne 1) {
        throw "Expected exactly one case-sensitive shape '$Name'; found $($matches.Count)"
    }
    return $matches[0]
}
