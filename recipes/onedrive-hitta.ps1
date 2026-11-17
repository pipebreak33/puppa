<#PUPPA
namn:        Hitta någons OneDrive
kategori:    OneDrive
typ:         läs
kräver:      pnp-admin
beskrivning: Ger URL, lagring och ägare för en användares OneDrive.
param:       Anvandare | Användare (UPN) |
#>
$url = (Get-PnPUserProfileProperty -Account $Anvandare -Properties PersonalUrl -ErrorAction Stop).PersonalUrl
if (-not $url) { throw "Ingen OneDrive hittad för $Anvandare (har den någonsin öppnats?)." }
$s = Get-PnPTenantSite -Identity $url.TrimEnd('/')
[pscustomobject]@{ Användare = $Anvandare; Url = $s.Url; 'GB' = [math]::Round($s.StorageUsageCurrent / 1024, 2); Ägare = $s.Owner; Låst = "$($s.LockState)" }
