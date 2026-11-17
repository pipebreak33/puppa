<#PUPPA
namn:        Lista grupper
kategori:    Grupper
typ:         läs
kräver:      graph:Group.Read.All
beskrivning: Alla grupper med typ (Microsoft 365, säkerhet, distribution) och om de är dynamiska.
param:       Filter? | Början av gruppnamnet |
#>
$p = @{ All = $true; Property = 'id,displayName,mail,groupTypes,securityEnabled,mailEnabled,membershipRule' }
if ($Filter) { $p.Filter = "startswith(displayName,'$($Filter -replace "'", "''")')" }
Get-MgGroup @p | Sort-Object DisplayName | Select-Object DisplayName, Mail,
    @{ n = 'Typ'; e = { if ($_.GroupTypes -contains 'Unified') { 'Microsoft 365' } elseif ($_.SecurityEnabled -and $_.MailEnabled) { 'E-postsäkerhet' } elseif ($_.SecurityEnabled) { 'Säkerhet' } else { 'Distribution' } } },
    @{ n = 'Dynamisk'; e = { [bool]$_.MembershipRule } }
