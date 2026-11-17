<#PUPPA
namn:        Ta bort webbplatsadministratör
kategori:    SharePoint
typ:         ändrar
kräver:      pnp:Url
beskrivning: Tar bort en användare som webbplatsadministratör från en sajt.
param:       Url | Sajtens URL | {rootUrl}/sites/
param:       Anvandare | Användare (UPN) |
#>
Remove-PnPSiteCollectionAdmin -Owners $Anvandare -ErrorAction Stop
Get-PnPSiteCollectionAdmin | Select-Object Title, Email
