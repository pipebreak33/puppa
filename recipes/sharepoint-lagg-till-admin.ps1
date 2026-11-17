<#PUPPA
namn:        Lägg till webbplatsadministratör
kategori:    SharePoint
typ:         ändrar
kräver:      pnp-admin
beskrivning: Gör en eller flera användare till webbplatsadministratörer på en sajt (även OneDrive).
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
param:       Anvandare | Användare (UPN, flera med komma) | {upn} | @spuser+
#>
$lista = @($Anvandare -split '[,;\s]+' | Where-Object { $_ })
Set-PnPTenantSite -Identity $Url -Owners $lista -ErrorAction Stop
Write-Host "Lade till: $($lista -join ', ')"
