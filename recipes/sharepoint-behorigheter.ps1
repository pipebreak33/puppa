<#PUPPA
namn:        Behörigheter på en sajt
kategori:    SharePoint
typ:         läs
kräver:      pnp:Url
beskrivning: Vem har åtkomst? Webbplatsadministratörer, SharePoint-grupper och deras medlemmar.
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
#>
Get-PnPSiteCollectionAdmin | ForEach-Object {
    [pscustomobject]@{ Grupp = '(Webbplatsadministratör)'; Roll = 'Full kontroll'; Medlem = $_.Title; Konto = $_.Email }
}
foreach ($g in Get-PnPGroup) {
    $roll = (Get-PnPGroupPermissions -Identity $g -ErrorAction SilentlyContinue | ForEach-Object Name) -join ', '
    $m = @(Get-PnPGroupMember -Group $g -ErrorAction SilentlyContinue)
    if (-not $m) { [pscustomobject]@{ Grupp = $g.Title; Roll = $roll; Medlem = '(tom)'; Konto = '' }; continue }
    foreach ($u in $m) { [pscustomobject]@{ Grupp = $g.Title; Roll = $roll; Medlem = $u.Title; Konto = $u.Email } }
}
