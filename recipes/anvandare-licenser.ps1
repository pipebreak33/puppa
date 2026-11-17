<#PUPPA
namn:        Licensöversikt
kategori:    Användare
typ:         läs
kräver:      graph:Organization.Read.All
beskrivning: Köpta, använda och lediga licenser per SKU.
#>
Get-MgSubscribedSku -All | Where-Object { $_.PrepaidUnits.Enabled -gt 0 } | Sort-Object SkuPartNumber |
    Select-Object SkuPartNumber,
        @{ n = 'Köpta'; e = { $_.PrepaidUnits.Enabled } },
        @{ n = 'Använda'; e = { $_.ConsumedUnits } },
        @{ n = 'Lediga'; e = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits } }
