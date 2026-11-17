<#PUPPA
namn:        Raderade sajter
kategori:    SharePoint
typ:         läs
kräver:      pnp-admin
beskrivning: Sajter i tenantens papperskorg (kan återställas inom 93 dagar).
#>
Get-PnPTenantDeletedSite | Sort-Object DeletionTime -Descending |
    Select-Object Url, @{ n = 'Raderad'; e = { $_.DeletionTime.ToString('yyyy-MM-dd') } }, DaysRemaining
