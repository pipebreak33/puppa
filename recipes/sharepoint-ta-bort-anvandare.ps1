<#PUPPA
namn:        Nollställ användares åtkomst på sajt
kategori:    SharePoint
typ:         ändrar
kräver:      pnp:Url
beskrivning: Tar bort användaren ur sajtens användarlista (alla grupper och direkta rättigheter). Bra när ett konto återskapats.
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
param:       Anvandare | Användare (UPN) | | @spuser
#>
$u = Get-PnPUser | Where-Object { $_.Email -eq $Anvandare -or $_.LoginName -like "*|$Anvandare" } | Select-Object -First 1
if (-not $u) { throw (T "Cannot find $Anvandare on the site." "Hittar inte $Anvandare på sajten.") }
Remove-PnPUser -Identity $u.Id -Force -ErrorAction Stop
Write-Host (T "Removed $($u.Title) ($($u.LoginName)) from $Url" "Tog bort $($u.Title) ($($u.LoginName)) från $Url")
