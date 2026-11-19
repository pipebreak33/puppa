<#PUPPA
namn:        Ge åtkomst till någons OneDrive
kategori:    OneDrive
typ:         ändrar
kräver:      pnp-admin
beskrivning: Lägger till en administratör på en användares OneDrive, t.ex. vid offboarding.
param:       Anvandare | Vems OneDrive? (UPN) | | @spuser
param:       Till | Vem ska få åtkomst? (UPN) | {upn} | @spuser
#>
$url = (Get-PnPUserProfileProperty -Account $Anvandare -Properties PersonalUrl -ErrorAction Stop).PersonalUrl
if (-not $url) { throw (T "No OneDrive found for $Anvandare." "Ingen OneDrive hittad för $Anvandare.") }
Set-PnPTenantSite -Identity $url.TrimEnd('/') -Owners $Till -ErrorAction Stop
[pscustomobject]@{ OneDrive = $url; 'Ny admin' = $Till }
