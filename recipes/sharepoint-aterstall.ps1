<#PUPPA
namn:        Återställ från papperskorgen
kategori:    SharePoint
typ:         ändrar
kräver:      pnp:Url
beskrivning: Återställer raderade objekt vars namn matchar, valfritt bara de som en viss person raderat.
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
param:       Filter | Del av filnamn (* = allt) |
param:       RaderadAv? | Bara det som raderats av (e-post) |
param:       Dagar | Raderat senaste X dagar | 7
#>
$fran = (Get-Date).AddDays(-[int]$Dagar)
$poster = Get-PnPRecycleBinItem -RowLimit 20000 | Where-Object {
    $_.DeletedDate -ge $fran -and ($Filter -eq '*' -or $_.LeafName -like "*$Filter*") -and (-not $RaderadAv -or $_.DeletedByEmail -eq $RaderadAv) }
Write-Host "Återställer $(@($poster).Count) objekt"
foreach ($p in $poster) {
    Restore-PnPRecycleBinItem -Identity $p -Force -ErrorAction Continue
    [pscustomobject]@{ Fil = $p.LeafName; Mapp = $p.DirName }
}
