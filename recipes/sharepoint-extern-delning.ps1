<#PUPPA
namn:        Extern delning per sajt
kategori:    SharePoint
typ:         läs
kräver:      pnp-admin
beskrivning: Vilka sajter som tillåter delning med gäster eller 'alla med länken'.
param:       Oppna | Visa bara sajter med delning påslagen? | Ja | Ja,Nej
#>
$t = Get-PnPTenant
Write-Host (T "Tenant level: $($t.SharingCapability)" "Tenantens nivå: $($t.SharingCapability)")
Get-PnPTenantSite | Where-Object { $Oppna -eq 'Nej' -or "$($_.SharingCapability)" -ne 'Disabled' } | Sort-Object Url |
    Select-Object Title, Url, @{ n = 'Delning'; e = { "$($_.SharingCapability)" } }
