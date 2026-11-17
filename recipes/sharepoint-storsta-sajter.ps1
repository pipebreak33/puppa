<#PUPPA
namn:        Största sajterna
kategori:    SharePoint
typ:         läs
kräver:      pnp-admin
beskrivning: Topplista över sajter som tar mest lagring.
param:       Antal | Hur många? | 20
#>
Get-PnPTenantSite | Sort-Object StorageUsageCurrent -Descending | Select-Object -First ([int]$Antal) |
    Select-Object Title, Url, @{ n = 'GB'; e = { [math]::Round($_.StorageUsageCurrent / 1024, 2) } },
        @{ n = 'Kvot GB'; e = { [math]::Round($_.StorageQuota / 1024, 0) } }
