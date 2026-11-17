<#PUPPA
namn:        Listor och bibliotek på en sajt
kategori:    SharePoint
typ:         läs
kräver:      pnp:Url
beskrivning: Synliga listor/bibliotek med antal objekt, senaste ändring och om behörigheterna är unika.
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
#>
Get-PnPList -Includes HasUniqueRoleAssignments | Where-Object { -not $_.Hidden } | Sort-Object Title |
    Select-Object Title, @{ n = 'Typ'; e = { if ($_.BaseTemplate -eq 101) { 'Bibliotek' } else { 'Lista' } } },
        ItemCount, @{ n = 'Ändrad'; e = { $_.LastItemUserModifiedDate.ToString('yyyy-MM-dd') } },
        @{ n = 'Unika beh.'; e = { $_.HasUniqueRoleAssignments } }
