<#PUPPA
namn:        Vad är jag ansluten till?
kategori:    Lokalt
typ:         läs
kräver:      ingen
beskrivning: Visar PnP-, Exchange-, Teams- och Graph-anslutningarna i Puppas PowerShell-session.
#>
$rader = @()
try { $c = Get-PnPConnection -ErrorAction Stop; $rader += [pscustomobject]@{ Tjänst = 'PnP'; Konto = "$($c.ConnectionType)"; Mål = $c.Url } }
catch { $rader += [pscustomobject]@{ Tjänst = 'PnP'; Konto = '–'; Mål = 'inte ansluten' } }
$exo = if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) { Get-ConnectionInformation | Select-Object -First 1 }
$rader += [pscustomobject]@{ Tjänst = 'Exchange'; Konto = if ($exo) { $exo.UserPrincipalName } else { '–' }; Mål = if ($exo) { $exo.TenantID } else { 'inte ansluten' } }
$mg = if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) { Get-MgContext }
$rader += [pscustomobject]@{ Tjänst = 'Graph'; Konto = if ($mg) { $mg.Account } else { '–' }; Mål = if ($mg) { $mg.TenantId } else { 'inte ansluten' } }
$rader
