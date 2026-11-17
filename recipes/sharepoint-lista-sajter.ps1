<#PUPPA
namn:        Lista sajter
kategori:    SharePoint
typ:         läs
kräver:      pnp-admin
beskrivning: Alla sajter i tenanten med mall, lagring och senaste ändring.
param:       Filter? | Del av titel eller URL att filtrera på |
#>
Get-PnPTenantSite |
    Where-Object { -not $Filter -or $_.Url -like "*$Filter*" -or $_.Title -like "*$Filter*" } |
    Sort-Object Url |
    Select-Object Title, Url, Template,
        @{ n = 'GB'; e = { [math]::Round($_.StorageUsageCurrent / 1024, 2) } },
        @{ n = 'Ändrad'; e = { $_.LastContentModifiedDate.ToString('yyyy-MM-dd') } }
