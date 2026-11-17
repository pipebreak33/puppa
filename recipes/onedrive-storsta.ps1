<#PUPPA
namn:        Största OneDrive-kontona
kategori:    OneDrive
typ:         läs
kräver:      pnp-admin
beskrivning: Topplista över OneDrive-konton efter lagring.
param:       Antal | Hur många? | 20
#>
Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" |
    Sort-Object StorageUsageCurrent -Descending | Select-Object -First ([int]$Antal) |
    Select-Object Owner, Url, @{ n = 'GB'; e = { [math]::Round($_.StorageUsageCurrent / 1024, 2) } }
