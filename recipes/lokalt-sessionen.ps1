<#PUPPA
namn:        Vad är jag ansluten till?
kategori:    Lokalt
typ:         läs
kräver:      ingen
beskrivning: Visar Exchange- och Graph-anslutningarna i Puppas PowerShell-session (PnP syns på statusraden).
#>
# PnP körs i en egen PowerShell-process (annars krockar MSAL-versionerna) – dess status syns på Puppas statusrad.
$rader = @([pscustomobject]@{ Tjänst = 'PnP'; Konto = '–'; Mål = 'egen process – se statusraden i Puppa' })
$exo = if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) { Get-ConnectionInformation | Select-Object -First 1 }
$rader += [pscustomobject]@{ Tjänst = 'Exchange'; Konto = if ($exo) { $exo.UserPrincipalName } else { '–' }; Mål = if ($exo) { $exo.TenantID } else { 'inte ansluten' } }
$mg = if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) { Get-MgContext }
$rader += [pscustomobject]@{ Tjänst = 'Graph'; Konto = if ($mg) { $mg.Account } else { '–' }; Mål = if ($mg) { $mg.TenantId } else { 'inte ansluten' } }
$rader
