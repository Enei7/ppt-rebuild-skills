$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\scripts\native_shape_lookup.ps1')
$slide=[pscustomobject]@{Shapes=@(
    [pscustomobject]@{Name='inline_G';Value='upper'},
    [pscustomobject]@{Name='inline_g';Value='lower'}
)}
if((Get-ExactSlideShape $slide 'inline_G').Value -cne 'upper'){throw 'Uppercase identity lost'}
if((Get-ExactSlideShape $slide 'inline_g').Value -cne 'lower'){throw 'Lowercase identity lost'}
$failed=$false
try{Get-ExactSlideShape $slide 'INLINE_G' | Out-Null}catch{$failed=$true}
if(-not $failed){throw 'Missing exact name must fail'}
$slide.Shapes += [pscustomobject]@{Name='inline_g';Value='duplicate'}
$failed=$false
try{Get-ExactSlideShape $slide 'inline_g' | Out-Null}catch{$failed=$true}
if(-not $failed){throw 'Ambiguous exact name must fail'}
Write-Output 'PASS exact shape lookup'
