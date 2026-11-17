<#PUPPA
namn:        Microsoft 365-grupper utan ägare
kategori:    Grupper
typ:         läs
kräver:      graph:Group.Read.All
beskrivning: Grupper/team som saknar ägare – ingen kan hantera dem.
#>
Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All | ForEach-Object {
    if (-not (Get-MgGroupOwner -GroupId $_.Id -Top 1)) { [pscustomobject]@{ Grupp = $_.DisplayName; Mail = $_.Mail; Skapad = $_.CreatedDateTime } }
}
