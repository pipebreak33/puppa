<#PUPPA
namn:        Inloggningar för en användare
kategori:    Entra ID
typ:         läs
kräver:      graph:AuditLog.Read.All,Directory.Read.All
beskrivning: Senaste inloggningarna: app, lyckades/fel och varför, IP och plats, enhet, MFA och villkorlig åtkomst (kräver Entra ID P1).
param:       Anvandare | Användare (UPN) | | @user
param:       Dagar | Hur många dagar bakåt? (högst 30) | 7
param:       Visa | Vilka inloggningar? | Alla | Alla,Misslyckade
#>
$fran = (Get-Date).ToUniversalTime().AddDays(-[Math]::Min([int]$Dagar, 30)).ToString('yyyy-MM-ddTHH:mm:ssZ')
$f = "userPrincipalName eq '$($Anvandare -replace "'", "''")' and createdDateTime ge $fran"
if ($Visa -eq 'Misslyckade') { $f += ' and status/errorCode ne 0' }
Get-MgAuditLogSignIn -Filter $f -Top 500 | ForEach-Object {
    [pscustomobject]@{
        Tid = $_.CreatedDateTime.ToLocalTime(); App = $_.AppDisplayName
        Resultat = if ($_.Status.ErrorCode -eq 0) { 'OK' } else { "Fel $($_.Status.ErrorCode): $($_.Status.FailureReason)" }
        IP = $_.IPAddress; Plats = (@($_.Location.City, $_.Location.CountryOrRegion) | Where-Object { $_ }) -join ', '
        Enhet = (@($_.DeviceDetail.DisplayName, $_.DeviceDetail.OperatingSystem, $_.DeviceDetail.Browser) | Where-Object { $_ }) -join ' · '
        Klient = $_.ClientAppUsed; MFA = $_.AuthenticationRequirement -replace 'multiFactorAuthentication', 'ja' -replace 'singleFactorAuthentication', 'nej'
        'Villkorlig åtkomst' = $_.ConditionalAccessStatus
    }
}
