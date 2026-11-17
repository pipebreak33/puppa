<#PUPPA
namn:        Om tenanten
kategori:    Entra ID
typ:         läs
kräver:      graph:Organization.Read.All,Domain.Read.All,Policy.Read.All
beskrivning: Organisation, tenant-ID, domäner (verifierade, standard), säkerhetsstandarder och tekniska kontakter.
#>
$o = Get-MgOrganization | Select-Object -First 1
$sd = try { (Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction Stop).IsEnabled } catch { '?' }
$dom = Get-MgDomain -All
[pscustomobject]@{
    Namn = $o.DisplayName; 'Tenant-ID' = $o.Id; Skapad = $o.CreatedDateTime; Land = $o.CountryLetterCode
    Standarddomän = ($dom | Where-Object IsDefault).Id
    Domäner = ($dom | ForEach-Object { $_.Id + $(if (-not $_.IsVerified) { ' (ej verifierad)' }) + $(if ($_.AuthenticationType -eq 'Federated') { ' (federerad)' }) }) -join ', '
    Säkerhetsstandarder = if ($sd -eq $true) { 'på' } elseif ($sd -eq $false) { 'av' } else { $sd }
    'Synkas från AD' = [bool]$o.OnPremisesSyncEnabled; 'Senaste synk' = $o.OnPremisesLastSyncDateTime
    'Tekniska kontakter' = $o.TechnicalNotificationMails -join ', '
}
